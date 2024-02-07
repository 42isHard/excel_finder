# Importation des modules nécessaires
import pandas as pd
import os

# Définition des chemins d'accès aux données
CHEMIN_DONNEES_SOURCE = "/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/"
CHEMIN_DONNEES_SORTIE = "/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/"
NOM_FICHIER_SORTIE = "marge_brute.xlsx"


# Fonction pour lister les fichiers correspondants
def lister_fichiers(dossier):
    """
    Liste les fichiers Excel dans le dossier spécifié qui commencent par
    "GL_analytique_FFE" ou "GL_analytique_FFSAS" et se terminent par ".xlsx".

    Args:
        dossier (str): Le chemin d'accès au dossier à analyser.

    Returns:
        list: Une liste de noms de fichiers correspondant aux critères.
    """
    fichiers_a_traiter = []
    for fichier in os.listdir(dossier):
        if (fichier.startswith("GL_analytique_FFE") or fichier.startswith("GL_analytique_FFSAS")) and fichier.endswith(
                ".xlsx"):
            fichiers_a_traiter.append(fichier)
    return fichiers_a_traiter


# Fonction pour charger les feuilles YTD avec la colonne "Entité juridique"
def charger_feuilles_YTD(dossier, fichiers):
    """
    Charge les feuilles "YTD" des fichiers Excel spécifiés et ajoute une
    colonne "Entité juridique" avec la valeur "FFE" ou "FFSAS".

    Args:
        dossier (str): Le chemin d'accès au dossier contenant les fichiers Excel.
        fichiers (list): Une liste de noms de fichiers à charger.

    Returns:
        list: Une liste de DataFrames Pandas.
    """
    dataframes = []
    for fichier in fichiers:
        entite = "FFE" if "GL_analytique_FFE" in fichier else "FFSAS"
        chemin_fichier_excel = os.path.join(dossier, fichier)
        xls = pd.ExcelFile(chemin_fichier_excel)
        for sheet_name in xls.sheet_names:
            if "YTD" in sheet_name:
                df = pd.read_excel(chemin_fichier_excel, sheet_name=sheet_name)
                df.insert(16, "Entité juridique", entite)  # Ajoute la colonne à la 17ème position (index 16)
                dataframes.append(df)
    return dataframes


def sauvegarder_donnees(dossier_sortie, dataframes):
    """
    Concatène les DataFrames et les sauvegarde dans un fichier Excel nommé
    "donnees_concatenees.xlsx" dans le dossier de sortie spécifié.

    Args:
        dossier_sortie (str): Le chemin d'accès au dossier de sortie.
        dataframes (list): Une liste de DataFrames Pandas à concaténer.

    Returns:
        None: Enregistre le fichier Excel sur le disque.
    """
    dataframe_concatene = pd.concat(dataframes, ignore_index=True)
    nom_fichier_sortie = "donnees_concatenees.xlsx"
    dataframe_concatene.to_excel(os.path.join(dossier_sortie, NOM_FICHIER_SORTIE), index=False)


def modifier_type_contrat(dataframes):
    """
    Modifie la valeur de la 11e colonne en fonction du type de contrat.

    Args:
        dataframes (list): Une liste de DataFrames Pandas.

    Returns:
        None: Modifie les DataFrames en place.
    """
    for df in dataframes:
        df.loc[df.iloc[:, 10] == 'fixe', df.columns[10]] = 'CF'
        df.loc[df.iloc[:, 10] == 'variable', df.columns[10]] = 'CV'


def modifier_valeur_Mgb(dataframes):
    """
    Modifie la valeur de la 10e colonne pour qu'elle soit toujours "Mgb".

    Args:
        dataframes (list): Une liste de DataFrames Pandas.

    Returns:
        None: Modifie les DataFrames en place.
    """
    for df in dataframes:
        df.iloc[:, 9] = 'Mgb'


def ajouter_annee_date(dataframes):
    """
    Ajoute la colonne de l'année de la date de la 6e colonne.

    Args:
        dataframes (list): Une liste de DataFrames Pandas.

    Returns:
        None: Modifie les DataFrames en place.
    """
    for df in dataframes:
        df.insert(17, "Année", pd.to_datetime(df.iloc[:, 5]).dt.year)


def ajouter_trimestre_date(dataframes):
    """
    Ajoute la colonne du trimestre de la date de la 6e colonne au format T1, T2, T3, T4.

    Args:
        dataframes (list): Une liste de DataFrames Pandas.

    Returns:
        None: Modifie les DataFrames en place.
    """

    def map_trimestre(trimestre):
        trimestres = {1: 'T1', 2: 'T2', 3: 'T3', 4: 'T4'}
        return trimestres.get(trimestre, '')

    for df in dataframes:
        df.insert(18, "Trimestre", pd.to_datetime(df.iloc[:, 5]).dt.quarter.map(map_trimestre))


# Exécution des fonctions
fichiers = lister_fichiers(CHEMIN_DONNEES_SOURCE)
dataframes = charger_feuilles_YTD(CHEMIN_DONNEES_SOURCE, fichiers)

modifier_type_contrat(dataframes)
modifier_valeur_Mgb(dataframes)
ajouter_annee_date(dataframes)
ajouter_trimestre_date(dataframes)

sauvegarder_donnees(CHEMIN_DONNEES_SORTIE, dataframes)

# Affichage d'un message de fin
print("\nTraitement terminé. Les données sont disponibles dans le fichier :")
print(os.path.join(CHEMIN_DONNEES_SORTIE, NOM_FICHIER_SORTIE))
