import os
import io
import zipfile
import uuid
from typing import List

import pandas as pd
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import JSONResponse, StreamingResponse

from processing import (
    process_folder,
    load_historique,
    save_historique,
    HISTORIQUE_DIR,
    HISTORIQUE_TEMPLATE,
)

app = FastAPI(title="PDF Stamping API")

zipped_cache: dict[str, bytes] = {}

INPUT_DIR = "input"
OUTPUT_DIR = "output"

os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)


def _mk_logger(logs: list[str]):
    def _log(msg: str):
        logs.append(msg)
    return _log


@app.post("/process-pdfs/")
async def process_pdfs(
    files: List[UploadFile] = File(...),
    matricule: str          = Form(...),
    note_tech:  bool = Form(False),
    num_nt:     str  = Form(""),
    gamme:      bool = Form(False),
    num_gs:     str  = Form(""),
    rvq_blanc:  bool = Form(False),
    rvq_orange: bool = Form(False),
    rvq_vert:   bool = Form(False),
    plan:       bool = Form(False),
):
    logs: list[str] = []

    try:
        # Nettoyage du dossier output
        for f in os.listdir(OUTPUT_DIR):
            os.remove(os.path.join(OUTPUT_DIR, f))

        # Sauvegarde des fichiers avec noms nettoyés
        for up in files:
            clean_name = up.filename.replace(" ", "_").replace('"', "").strip()
            dest = os.path.join(INPUT_DIR, clean_name)
            with open(dest, "wb") as f:
                f.write(await up.read())
            logs.append(f"Upload : {clean_name}")

        hist_path = os.path.join(
            HISTORIQUE_DIR,
            HISTORIQUE_TEMPLATE.format(matricule=matricule)
        )
        os.makedirs(HISTORIQUE_DIR, exist_ok=True)
        df_hist = load_historique(hist_path)

        opts = {
            "matricule":  matricule,
            "note_tech":  note_tech,
            "num_nt":     num_nt,
            "gamme":      gamme,
            "num_gs":     num_gs,
            "rvq_blanc":  rvq_blanc,
            "rvq_orange": rvq_orange,
            "rvq_vert":   rvq_vert,
            "plan":       plan,
        }

        results = process_folder(INPUT_DIR, OUTPUT_DIR, opts, log_callback=_mk_logger(logs))

        new_rows = [r["hist"] for r in results if "hist" in r]
        if new_rows:
            df_hist = pd.concat([df_hist, pd.DataFrame(new_rows)], ignore_index=True)
            save_historique(df_hist, hist_path)
            logs.append("Historique Excel mis à jour.")

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as z:
            for root, _, fs in os.walk(OUTPUT_DIR):
                for f in fs:
                    full = os.path.join(root, f)
                    arc = os.path.relpath(full, OUTPUT_DIR)
                    z.write(full, arc)
        zip_buf.seek(0)

        session_id = str(uuid.uuid4())
        zipped_cache[session_id] = zip_buf.read()

        return {
            "status": "success",
            "download_url": f"/download-zip/{session_id}",
            "logs": logs,
        }

    except Exception as exc:
        return JSONResponse(status_code=500, content={"error": str(exc), "logs": logs})


@app.get("/download-zip/{session_id}")
def download_zip(session_id: str):
    data = zipped_cache.get(session_id)
    if not data:
        return JSONResponse(status_code=404, content={"error": "invalid session"})
    return StreamingResponse(
        io.BytesIO(data),
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=rvq_results_{session_id}.zip"},
    )


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000, reload=True)
