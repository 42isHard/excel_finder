import pandas as pd
import sys
import os


def convertir_xlsx_en_csv(chemin_xlsx, chemin_csv):
    """
    Convertit un fichier XLSX en fichier CSV.

    Args:
    chemin_xlsx (str): Chemin du fichier XLSX à convertir.
    chemin_csv (str): Chemin où le fichier CSV sera enregistré.
    """
    df = pd.read_excel(chemin_xlsx)
    df.to_csv(chemin_csv, index=False)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python script.py chemin_du_fichier.xlsx")
    else:
        chemin_xlsx = sys.argv[1]
        chemin_csv = os.path.splitext(chemin_xlsx)[0] + '.csv'
        convertir_xlsx_en_csv(chemin_xlsx, chemin_csv)
        print(f"Fichier converti et enregistré sous : {chemin_csv}")
