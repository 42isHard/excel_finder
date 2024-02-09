import pandas as pd
import sys


def convertir_csv_en_xlsx(chemin_csv, chemin_xlsx):
    """Convertit un fichier CSV en fichier Excel (XLSX).
    python CSV_TO_XLSX.py /home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/marge_brute_CROSS.csv
    """
    df = pd.read_csv(chemin_csv)
    df.to_excel(chemin_xlsx, index=False)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py chemin_du_fichier.csv")
    else:
        chemin_csv = sys.argv[1]
        chemin_xlsx = chemin_csv.replace('.csv', '.xlsx')
        convertir_csv_en_xlsx(chemin_csv, chemin_xlsx)
        print(f"Fichier converti et enregistré sous : {chemin_xlsx}")
