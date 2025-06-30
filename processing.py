import os
import shutil
from datetime import datetime
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


def organiser_fichiers_par_date(dossier_source, dossier_destination, log_callback=None):
    """
    Étape 1 : Tirer automatiquement les sérials par campagne
    """
    if not os.path.isdir(dossier_source):
        log_callback(f"Le dossier source n'existe pas: {dossier_source}")
        return
    os.makedirs(dossier_destination, exist_ok=True)

    for fichier in os.listdir(dossier_source):
        if not (fichier.endswith('.xlsx') or fichier.endswith('.xls')):
            continue
        chemin_fichier = os.path.join(dossier_source, fichier)
        try:
            classeur = openpyxl.load_workbook(chemin_fichier, data_only=True)
        except Exception as e:
            log_callback(f"'{fichier}'")
            continue
        feuille = classeur.active
        valeur_date = feuille['B2'].value
        if valeur_date is None:
            log_callback(f"Pas de date dans {fichier}")
            continue
        if isinstance(valeur_date, datetime):
            date_extraite = valeur_date
        else:
            try:
                date_extraite = datetime.strptime(str(valeur_date), "%d/%m/%Y %H:%M")
            except ValueError:
                try:
                    date_extraite = datetime.strptime(str(valeur_date), "%d/%m/%Y")
                except ValueError:
                    date_extraite = None
        if date_extraite:
            nom_dossier = date_extraite.strftime("%Y-%m-%d")
        else:
            nom_dossier = str(valeur_date).replace('/', '_').replace(':', '_').replace(' ', '_')
        cible = os.path.join(dossier_destination, nom_dossier)
        os.makedirs(cible, exist_ok=True)
        try:
            shutil.move(chemin_fichier, os.path.join(cible, fichier))
            log_callback(f"{fichier} → {nom_dossier}")
        except Exception as e:
            log_callback(f"Erreur déplacement {fichier}: {e}")
    log_callback("Étape 1 terminée.")


def generate_individual_curves(input_dir, output_dir, log_callback=None):
    """
    Étape 2 : Affichage des courbes serial/serial pour chaque fichier
    """
    files = []
    for root, _, fs in os.walk(input_dir):
        for f in fs:
            if f.lower().endswith('.xlsx') and '_traite' not in f:
                files.append(os.path.join(root, f))
    total = len(files)
    for i, path in enumerate(files, 1):
        try:
            # Renommer
            base = os.path.splitext(path)[0]
            new_path = base + '_traite' + os.path.splitext(path)[1]
            os.rename(path, new_path)
            df = pd.read_excel(new_path, header=2)
            df_sel = df[['Temps', 'Effort[t]', 'Hauteur[mm]', 'Vitesse[mm/s]']].copy()
            df_sel['Temps'] = pd.to_numeric(df_sel['Temps'], errors='coerce') / 1000.0
            for col in ['Effort[t]', 'Hauteur[mm]', 'Vitesse[mm/s]']:
                df_sel[col] = pd.to_numeric(df_sel[col], errors='coerce')
            df_sel.rename(columns={'Temps': 'TPS EN SECONDE'}, inplace=True)
            # Création sortie
            rel = os.path.relpath(os.path.dirname(new_path), input_dir)
            target_dir = os.path.join(output_dir, rel)
            os.makedirs(target_dir, exist_ok=True)
            out_file = os.path.join(target_dir, f"Courbe_{os.path.basename(base)}.xlsx")
            writer = pd.ExcelWriter(out_file, engine='xlsxwriter')
            df_sel.to_excel(writer, sheet_name='Data', index=False)
            wb = writer.book
            ws = writer.sheets['Data']
            start_row = 1
            start_col = 8
            for idx, col in enumerate(['Effort[t]', 'Hauteur[mm]', 'Vitesse[mm/s]']):
                chart = wb.add_chart({'type': 'line'})
                chart.add_series({
                    'name': col,
                    'categories': f'=Data!$A$2:$A${len(df_sel)+1}',
                    'values': f'=Data!${chr(66+idx)}$2:${chr(66+idx)}${len(df_sel)+1}'
                })
                chart.set_title({'name': col})
                chart.set_x_axis({'name': 'TPS EN SECONDE'})
                chart.set_y_axis({'name': col.split(' ')[0]})
                ws.insert_chart(f'{chr(65+start_col)}{start_row}', chart, {'x_offset':15,'y_offset':15})
                start_row += 16
            writer.close()
            log_callback(f"Traitement {i}/{total} : {out_file}")
        except Exception as e:
            log_callback(f"Erreur sur {path}: {e}")
    log_callback("Étape 2 terminée.")


def generate_concat_subfolders(curves_dir, log_callback=None):
    """
    Étape 3 : Concaténation des courbes dans chaque sous-dossier
    """
    for root, _, files in os.walk(curves_dir):
        efforts, hauteurs, vitesses = [], [], []
        for f in files:
            if f.lower().endswith('.xlsx') and 'Courbe_' in f:
                df = pd.read_excel(os.path.join(root, f), sheet_name='Data')
                efforts.append((df['TPS EN SECONDE'], df['Effort[t]'], f))
                hauteurs.append((df['TPS EN SECONDE'], df['Hauteur[mm]'], f))
                vitesses.append((df['TPS EN SECONDE'], df['Vitesse[mm/s]'], f))
        if efforts:
            sub = os.path.basename(root)
            pdf = os.path.join(root, f'Courbes_concat_{sub}.pdf')
            with PdfPages(pdf) as pp:
                for data, title, ylabel in [(efforts, 'Effort[t]', 'Effort[t]'),
                                            (hauteurs, 'Hauteur[mm]', 'Hauteur[mm]'),
                                            (vitesses, 'Vitesse[mm/s]', 'Vitesse[mm/s]')]:
                    plt.figure(figsize=(10,6))
                    for x, y, lbl in data:
                        plt.plot(x, y, label=lbl)
                    plt.title(f'Courbe {title}')
                    plt.xlabel('Temps (s)')
                    plt.ylabel(ylabel)
                    plt.legend(loc='center left', bbox_to_anchor=(1,0.5), fontsize='small')
                    plt.grid(True)
                    plt.tight_layout()
                    pp.savefig()
                    plt.close()
            log_callback(f"PDF sous-dossier: {pdf}")
    log_callback("Étape 3 terminée.")


def generate_concat_all(curves_root, log_callback=None):
    """
    Étape 4 : Concaténation toutes campagnes confondues
    """
    parents = [os.path.join(curves_root, d) for d in os.listdir(curves_root)
               if os.path.isdir(os.path.join(curves_root, d))]
    for p in parents:
        e, h, v = [], [], []
        for root, _, files in os.walk(p):
            for f in files:
                if f.lower().endswith('.xlsx') and 'Courbe_' in f:
                    df = pd.read_excel(os.path.join(root, f), sheet_name='Data')
                    e.append((df['TPS EN SECONDE'], df['Effort[t]'], f))
                    h.append((df['TPS EN SECONDE'], df['Hauteur[mm]'], f))
                    v.append((df['TPS EN SECONDE'], df['Vitesse[mm/s]'], f))
        if e:
            pdf = os.path.join(p, f'{os.path.basename(p)}.pdf')
            with PdfPages(pdf) as pp:
                for data, title, ylabel in [(e, 'Effort[t]', 'Effort[t]'),
                                            (h, 'Hauteur[mm]', 'Hauteur[mm]'),
                                            (v, 'Vitesse[mm/s]', 'Vitesse[mm/s]')]:
                    plt.figure(figsize=(15,12))
                    for x, y, lbl in data:
                        plt.plot(x, y, label=lbl)
                    plt.title(f'Courbe {title}')
                    plt.xlabel('Temps (s)')
                    plt.ylabel(ylabel)
                    plt.legend(loc='center left', bbox_to_anchor=(1,0.5), fontsize='small')
                    plt.grid(True)
                    plt.tight_layout(rect=[0,0,0.75,1])
                    pp.savefig()
                    plt.close()
            log_callback(f"PDF global: {pdf}")
    log_callback("Étape 4 terminée.")
