from fastapi import FastAPI, UploadFile, File
from fastapi.responses import JSONResponse
import os
import tempfile
from typing import List
import zipfile
import io
from processing import (
    organiser_fichiers_par_date,
    generate_individual_curves,
    generate_concat_subfolders,
    generate_concat_all
)

# Dictionnaire temporaire pour stocker les résultats par session
zipped_results_cache = {}

app = FastAPI(title="Curve Generation API")

def create_log_function(logs: list):
    def log_callback(msg: str):
        logs.append(msg)
    return log_callback


@app.post("/process-files/")
async def process_files(files: List[UploadFile] = File(...)):
    logs = []
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            input_dir = os.path.join(tmpdir, "input")
            output_dir = os.path.join(tmpdir, "output")
            os.makedirs(input_dir, exist_ok=True)
            os.makedirs(output_dir, exist_ok=True)

            # Enregistrer les fichiers Excel reçus
            for file in files:
                file_path = os.path.join(input_dir, file.filename)
                with open(file_path, "wb") as f:
                    f.write(await file.read())
                logs.append(f"Fichier reçu : {file.filename}")

            # Exécution des étapes
            log_callback = create_log_function(logs)
            organiser_fichiers_par_date(input_dir, output_dir, log_callback)
            generate_individual_curves(input_dir, output_dir, log_callback)
            generate_concat_subfolders(output_dir, log_callback)
            generate_concat_all(output_dir, log_callback)
            
            # Créer un fichier zip en mémoire
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for root, _, files in os.walk(output_dir):
                    for file in files:
                        full_path = os.path.join(root, file)
                        rel_path = os.path.relpath(full_path, output_dir)
                        zip_file.write(full_path, arcname=rel_path)
            zip_buffer.seek(0)

            # Stocker le zip dans le cache avec un ID unique
            import uuid
            session_id = str(uuid.uuid4())
            zipped_results_cache[session_id] = zip_buffer.read()

            # Inclure l'ID dans la réponse
            return {
                "status": "Traitement terminé",
                "log": logs,
                "download_url": f"/download-zip/{session_id}"
            }

    except Exception as e:
        return JSONResponse(status_code=500, content={
            "status": "Erreur",
            "message": str(e),
            "log": logs
        })


from fastapi.responses import StreamingResponse

@app.get("/download-zip/{session_id}")
def download_zip(session_id: str):
    if session_id not in zipped_results_cache:
        return JSONResponse(status_code=404, content={"error": "Session non trouvée."})
    
    zip_bytes = zipped_results_cache[session_id]
    return StreamingResponse(io.BytesIO(zip_bytes), media_type="application/zip", headers={
        "Content-Disposition": f"attachment; filename=results_{session_id}.zip"
    })