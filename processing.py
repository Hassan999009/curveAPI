"""
processing.py
Module cœur pour le traitement des RVQ :
  1. Extraction des métadonnées PDF
  2. Ajout des tampons (Note Tech, Gamme Spéciale, RVQ Blanc/Orange/Vert, Plan Expédition)
  3. Mise à jour de l’historique Excel (protégé par mot de passe)
"""

import os
import io
import re
from datetime import datetime
from typing import List, Tuple, Dict, Optional

import fitz                    # PyMuPDF
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# --- Excel chiffré -------------------------------------------------------------
import msoffcrypto
from msoffcrypto import OfficeFile
from msoffcrypto.exceptions import DecryptionError
from openpyxl import load_workbook

# -------------------------------------------------------------------------------
# Configuration globale (adapter les chemins si nécessaire)
# -------------------------------------------------------------------------------
BASE_REP        = r"C:\Users\s623521\Desktop\Work\PDFTools boucle froide\fichers"
IMG_DIR         = os.path.join(BASE_REP, "images")

GAMME_SPECIAL_IMG  = os.path.join(IMG_DIR, "Gamme_speciale.png")
NOTE_TECH_IMG      = os.path.join(IMG_DIR, "Note_Technique.png")
RVQ_BLANC_IMG      = os.path.join(IMG_DIR, "livraison_directe.png")
RVQ_ORANGE_IMG     = os.path.join(IMG_DIR, "RVQ_ORANGE.png")
RVQ_VERT_IMG       = os.path.join(IMG_DIR, "RVQ_VERT.png")
TAMPON_PLAN_IMG    = os.path.join(IMG_DIR, "Tampon_plan_expedition.png")

HISTORIQUE_DIR      = os.path.join(BASE_REP, "Fichier excel")
HISTORIQUE_TEMPLATE = "Historique des RVQ fusionnés_{matricule}.xlsx"
HISTORIQUE_SHEET    = "Sheet1"
HISTORIQUE_PASSWORD = "SNECMA"

# -------------------------------------------------------------------------------
# REGEX pour extractions PDF
# -------------------------------------------------------------------------------
_RDV  = re.compile(r"Destinataire\s*:\s*(\d+)")
_REF  = re.compile(r"Quantité totale\s*:\s*([\d\-]+)")
_DEST = re.compile(r"Destinataire\s*:\s*(.*)\nRef\. article DTA", re.DOTALL)
_QTY  = re.compile(r"Quantité totale\s*:\s*.*?(\d+)\s*pièces", re.DOTALL)

# -------------------------------------------------------------------------------
# UTILITAIRES Excel chiffré
# -------------------------------------------------------------------------------
def _read_protected_excel(path: str, password: str) -> pd.DataFrame:
    try:
        with open(path, "rb") as f:
            off = OfficeFile(f)
            off.load_key(password=password)
            buf = io.BytesIO()
            off.decrypt(buf)
            buf.seek(0)
            return pd.read_excel(buf, sheet_name=HISTORIQUE_SHEET)
    except DecryptionError:
        return pd.read_excel(path, sheet_name=HISTORIQUE_SHEET)

def load_historique(path: str) -> pd.DataFrame:
    if not os.path.exists(path):
        return pd.DataFrame(columns=["Date", "Matricule", "N°Réf", "OF", "N°RVQ", "Qté"])
    return _read_protected_excel(path, HISTORIQUE_PASSWORD)

def save_historique(df: pd.DataFrame, path: str):
    tmp = path.replace(".xlsx", "_tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as wr:
        df.to_excel(wr, sheet_name=HISTORIQUE_SHEET, index=False)

    # Version simplifiée pour Linux (Render) sans win32com
    os.replace(tmp, path)

    # Nettoyage si besoin
    if os.path.exists(tmp):
        os.remove(tmp)

# -------------------------------------------------------------------------------
# EXTRACTION PDF
# -------------------------------------------------------------------------------
def extract_pdf_metadata(pdf_path: str) -> Optional[Tuple[str, str, str, str, str, str]]:
    try:
        doc = fitz.open(pdf_path)
        text = doc[0].get_text()
        doc.close()
    except Exception:
        return None

    filename = os.path.basename(pdf_path)
    of_match = re.match(r"(\d+)", filename)
    of = of_match.group(1) if of_match else None

    rdv  = _RDV.search(text)
    ref  = _REF.search(text)
    dest = _DEST.search(text)
    qty  = _QTY.search(text)

    rdv_val  = rdv.group(1)        if rdv  else ""
    ref_val  = ref.group(1)        if ref  else ""
    dest_val = dest.group(1).split("\n")[-1].strip() if dest else ""
    qty_val  = qty.group(1)        if qty  else ""

    if not all([rdv_val, ref_val, dest_val, of]):
        return None

    out_name = f"{ref_val} {rdv_val} OF{of} Qt-{qty_val}.pdf"
    return out_name, dest_val, rdv_val, ref_val, qty_val, of

# -------------------------------------------------------------------------------
# GÉNÉRATION DE TAMPONS
# -------------------------------------------------------------------------------
def _load_image(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()

def _stamp_with_text(base_img_path: str, text: str) -> bytes:
    img  = Image.open(base_img_path).convert("RGBA")
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("arial.ttf", 20)
    draw.text((10, 10), text, fill="black", font=font)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def apply_tampons(
    pdf_in:  str,
    pdf_out: str,
    *,
    note_tech: bool = False,
    num_nt:   str  = "",
    gamme:    bool = False,
    num_gs:   str  = "",
    rvq_blanc:   bool = False,
    rvq_orange:  bool = False,
    rvq_vert:    bool = False,
    plan:        bool = False,
):
    doc  = fitz.open(pdf_in)
    page = doc[0]
    w, h = page.rect.width, page.rect.height

    stamps: List[bytes] = []

    if note_tech and num_nt:
        stamps.append(_stamp_with_text(NOTE_TECH_IMG, num_nt))
    if gamme and num_gs:
        stamps.append(_stamp_with_text(GAMME_SPECIAL_IMG, f"DA-{num_gs}"))
    if rvq_blanc:
        stamps.append(_load_image(RVQ_BLANC_IMG))
    if rvq_orange:
        stamps.append(_load_image(RVQ_ORANGE_IMG))
    if rvq_vert:
        stamps.append(_load_image(RVQ_VERT_IMG))
    if plan:
        stamps.append(_load_image(TAMPON_PLAN_IMG))

    cur_y = h - 140
    for img_bytes in stamps:
        rect = fitz.Rect(w - 210, cur_y, w - 50, cur_y + 120)
        page.insert_image(rect, stream=img_bytes, keep_proportion=True)
        cur_y -= 130

    doc.save(pdf_out)
    doc.close()

# -------------------------------------------------------------------------------
# TRAITEMENT D’UN PDF
# -------------------------------------------------------------------------------
def process_one_pdf(pdf_path: str, dest_dir: str, opts: Dict) -> Optional[Dict]:
    meta = extract_pdf_metadata(pdf_path)
    if not meta:
        return None
    out_name, dest, rdv, ref_art, qty, of = meta
    os.makedirs(dest_dir, exist_ok=True)
    out_pdf = os.path.join(dest_dir, out_name)

    apply_tampons(
        pdf_in=pdf_path,
        pdf_out=out_pdf,
        note_tech   = opts.get("note_tech"),
        num_nt      = opts.get("num_nt", ""),
        gamme       = opts.get("gamme"),
        num_gs      = opts.get("num_gs", ""),
        rvq_blanc   = opts.get("rvq_blanc"),
        rvq_orange  = opts.get("rvq_orange"),
        rvq_vert    = opts.get("rvq_vert"),
        plan        = opts.get("plan"),
    )

    hist_line = {
        "Date":      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Matricule": opts.get("matricule"),
        "N°Réf":     ref_art,
        "OF":        of,
        "N°RVQ":     rdv,
        "Qté":       qty,
    }
    return {"hist": hist_line, "pdf_out": out_pdf}

# -------------------------------------------------------------------------------
# TRAITEMENT D’UN DOSSIER
# -------------------------------------------------------------------------------
def process_folder(
    inp: str,
    out_dir: str,
    opts: Dict,
    log_callback=None
) -> List[Dict]:
    results = []
    for fname in os.listdir(inp):
        if fname.lower().endswith(".pdf"):
            res = process_one_pdf(os.path.join(inp, fname), out_dir, opts)
            if res:
                results.append(res)
                if log_callback:
                    log_callback(f"OK : {fname} → {os.path.basename(res['pdf_out'])}")
            elif log_callback:
                log_callback(f"Échec sur : {fname}")
    return results
