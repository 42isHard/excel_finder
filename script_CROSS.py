import pandas as pd
from tqdm import tqdm

# Constantes pour les chemins des fichiers
CHEMIN_SOURCES_CSV = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/sources_CROSS.csv'
CHEMIN_MARGE_BRUTE_CSV = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/marge_brute_CONCAT.csv'
CHEMIN_MARGE_BRUTE_CROSS_CSV = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/marge_brute_CROSS.csv'


def charger_csv(chemin_fichier):
    """Fonction pour charger un fichier CSV."""
    return pd.read_csv(chemin_fichier, dtype=str)


def nettoyer_noms_colonnes(df):
    """Fonction pour nettoyer les noms de colonnes."""
    return df.rename(columns=str.strip)


def fusionner_sources_et_marge_brute(sources, marge_brute):
    """Fonction pour fusionner les données sources et la marge brute."""
    return pd.merge(marge_brute, sources, left_on='Section Analytique - Code', right_on='codes analytiques', how='left')


def copier_colonnes_specifiques(resultat, colonnes, df):
    """Fonction pour copier les colonnes spécifiques."""
    for colonne in tqdm(colonnes, desc="Copie des colonnes", unit="colonne"):
        df[colonne] = resultat[colonne]


def sauvegarder_csv(df, chemin_fichier):
    """Fonction pour sauvegarder un DataFrame en CSV."""
    df.to_csv(chemin_fichier, index=False)


def main():
    """Fonction principale."""
    # Charger les données sources et la marge brute
    sources = charger_csv(CHEMIN_SOURCES_CSV)
    marge_brute = charger_csv(CHEMIN_MARGE_BRUTE_CSV)

    # Nettoyer les noms de colonnes
    sources = nettoyer_noms_colonnes(sources)
    marge_brute = nettoyer_noms_colonnes(marge_brute)

    # Supprimer la colonne "Commentaire" si elle existe
    if "Commentaire" in marge_brute.columns:
        marge_brute.drop(columns=["Commentaire"], inplace=True)

    # Fusionner les données basée sur la correspondance des clés
    resultat = fusionner_sources_et_marge_brute(sources, marge_brute)

    # Colonnes à copier
    colonnes_a_copier = ["TYPE", "BL", "Groupe", "Compte", "GC", "Intérêt", "Session", "Date de fin", "Type de groupe",
                         "Secteur d'activité", "Titre du produit", "BL d'origine", "grand dispo ou pas",
                         "mode Formation", "regroupements Secteurs d'Activité", "pays facturation"]

    # Copier les colonnes spécifiques
    copier_colonnes_specifiques(resultat, colonnes_a_copier, marge_brute)

    # Sauvegarder le fichier modifié
    sauvegarder_csv(marge_brute, CHEMIN_MARGE_BRUTE_CROSS_CSV)


if __name__ == "__main__":
    main()
