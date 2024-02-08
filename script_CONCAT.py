import pandas as pd
import os
from tqdm import tqdm

# Définition des chemins d'accès aux données
CHEMIN_DONNEES_SOURCE = "/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/"
CHEMIN_DONNEES_SORTIE = "/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/"
NOM_FICHIER_SORTIE = "marge_brute.xlsx"


# Fonction pour lister les fichiers correspondants
def lister_fichiers(dossier):
    fichiers_a_traiter = []
    with os.scandir(dossier) as entries:
        for entry in entries:
            if entry.is_file() and (entry.name.startswith("GL_analytique_FFE") or entry.name.startswith(
                    "GL_analytique_FFSAS")) and entry.name.endswith(".xlsx"):
                fichiers_a_traiter.append(entry.name)
    return fichiers_a_traiter


# Fonction pour charger les feuilles YTD avec la colonne "Entité juridique"
def charger_feuilles_YTD(dossier, fichiers):
    dataframes = []
    for fichier in tqdm(fichiers, desc="Chargement des fichiers", unit="fichier"):
        entite = "FFE" if "GL_analytique_FFE" in fichier else "FFSAS"
        chemin_fichier_excel = os.path.join(dossier, fichier)
        xls = pd.read_excel(chemin_fichier_excel, sheet_name=None)
        for sheet_name, df in xls.items():
            if "YTD" in sheet_name:
                df.insert(16, "Entité juridique", entite)
                dataframes.append(df)
    return dataframes


# Fonction pour appliquer les modifications de dataframe
def modifier_dataframes(dataframes):
    for df in tqdm(dataframes, desc="Modification des dataframes", unit="dataframe"):
        # Modifier type de contrat
        df.loc[df.iloc[:, 10] == 'fixe', df.columns[10]] = 'CF'
        df.loc[df.iloc[:, 10] == 'variable', df.columns[10]] = 'CV'
        # Modifier valeur Mgb
        df.iloc[:, 9] = 'Mgb'
        # Ajouter année et trimestre
        df.insert(17, "Année", pd.to_datetime(df.iloc[:, 5]).dt.year)
        trimestres = {1: 'T1', 2: 'T2', 3: 'T3', 4: 'T4'}
        df.insert(18, "Trimestre", pd.to_datetime(df.iloc[:, 5]).dt.quarter.map(lambda x: trimestres.get(x, '')))


def sauvegarder_donnees(dossier_sortie, dataframes):
    dataframe_concatene = pd.concat(dataframes, ignore_index=True)
    dataframe_concatene.to_excel(os.path.join(dossier_sortie, NOM_FICHIER_SORTIE), index=False)


# Exécution des fonctions
fichiers = lister_fichiers(CHEMIN_DONNEES_SOURCE)
dataframes = charger_feuilles_YTD(CHEMIN_DONNEES_SOURCE, fichiers)
modifier_dataframes(dataframes)
sauvegarder_donnees(CHEMIN_DONNEES_SORTIE, dataframes)

# Affichage d'un message de fin
print("\nTraitement terminé. Les données sont disponibles dans le fichier :")
print(os.path.join(CHEMIN_DONNEES_SORTIE, NOM_FICHIER_SORTIE))
