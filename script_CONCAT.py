import pandas as pd
import os
from tqdm import tqdm

CHEMIN_DONNEES_SOURCE = "/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/"
CHEMIN_DONNEES_SORTIE = "/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/"
NOM_FICHIER_SORTIE = "marge_brute_CONCAT.csv"


def lister_fichiers(dossier):
    """Liste les fichiers Excel dans le dossier."""
    fichiers_a_traiter = []
    with os.scandir(dossier) as entries:
        for entry in entries:
            if entry.is_file() and (entry.name.startswith("Grand Livre FFE") or entry.name.startswith(
                    "Grand Livre FFSAS")) and entry.name.endswith(".xlsx"):
                fichiers_a_traiter.append(entry.name)
    return fichiers_a_traiter


def charger_feuilles_YTD(dossier, fichiers):
    """Charge les feuilles YTD avec la colonne 'Entité juridique'."""
    dataframes = []
    for fichier in tqdm(fichiers, desc="Chargement des fichiers", unit="fichier"):
        entite = "FFE" if "Grand Livre FFE" in fichier else "FFSAS"
        chemin_fichier_excel = os.path.join(dossier, fichier)
        try:
            xls = pd.read_excel(chemin_fichier_excel, sheet_name=None)
            for sheet_name, df in xls.items():
                if "YTD" in sheet_name:
                    df.insert(16, "Entité juridique", entite)
                    dataframes.append(df)
        except Exception as e:
            print(f"Erreur lors du chargement du fichier {fichier}: {e}")
    return dataframes


def modifier_dataframes(dataframes):
    """Modifie les dataframes."""
    for df in tqdm(dataframes, desc="Modification des dataframes", unit="dataframe"):
        df.loc[df.iloc[:, 10] == 'fixe', df.columns[10]] = 'CF'
        df.loc[df.iloc[:, 10] == 'variable', df.columns[10]] = 'CV'
        df.iloc[:, 9] = 'Mgb'
        df.insert(17, "Année", pd.to_datetime(df.iloc[:, 5]).dt.year)
        trimestres = {1: 'T1', 2: 'T2', 3: 'T3', 4: 'T4'}
        df.insert(18, "Trimestre", pd.to_datetime(df.iloc[:, 5]).dt.quarter.map(lambda x: trimestres.get(x, '')))


def sauvegarder_donnees(dossier_sortie, dataframes):
    """Sauvegarde les données dans un fichier CSV."""
    dataframe_concatene = pd.concat(dataframes, ignore_index=True)
    dataframe_concatene.to_csv(os.path.join(dossier_sortie, NOM_FICHIER_SORTIE), index=False)


def main():
    """Fonction principale."""
    fichiers = lister_fichiers(CHEMIN_DONNEES_SOURCE)
    dataframes = charger_feuilles_YTD(CHEMIN_DONNEES_SOURCE, fichiers)
    modifier_dataframes(dataframes)
    sauvegarder_donnees(CHEMIN_DONNEES_SORTIE, dataframes)
    print("\nTraitement terminé. Les données sont disponibles dans le fichier :")
    print(os.path.join(CHEMIN_DONNEES_SORTIE, NOM_FICHIER_SORTIE))


if __name__ == "__main__":
    main()
