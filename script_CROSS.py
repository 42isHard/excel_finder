import pandas as pd

# Chemins vers les fichiers
chemin_sources = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/sources.xlsx'
chemin_marge_brute = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/marge_brute.xlsx'

# Fonction pour nettoyer les noms de colonnes
def nettoyer_noms_colonnes(df):
    df.columns = df.columns.str.strip()
    return df

# Lire les noms des feuilles de calcul et trouver celle qui contient "BDD"
nom_feuille = None
with pd.ExcelFile(chemin_sources) as xls:
    for sheet_name in xls.sheet_names:
        if "BDD" in sheet_name:
            nom_feuille = sheet_name
            break

# Vérifier si une feuille correspondante a été trouvée
if nom_feuille:
    # Chargement des fichiers
    sources = pd.read_excel(chemin_sources, sheet_name=nom_feuille)
    marge_brute = pd.read_excel(chemin_marge_brute)

    # Nettoyer les noms de colonnes
    sources = nettoyer_noms_colonnes(sources)
    marge_brute = nettoyer_noms_colonnes(marge_brute)

    # Supprimer la colonne "Commentaire" si elle existe
    if "Commentaire" in marge_brute.columns:
        marge_brute.drop(columns=["Commentaire"], inplace=True)

    # Fusion des données basée sur la correspondance des clés
    resultat = pd.merge(marge_brute, sources, left_on='Section Analytique - Code', right_on='codes analytiques', how='left')

    # Sélection des colonnes à copier
    colonnes_a_copier = ["TYPE", "BL", "Groupe", "Compte", "GC", "Intérêt", "Session", "Date de fin", "Type de groupe", "Secteur d'activité", "Titre du produit", "BL d'origine", "grand dispo ou pas", "mode Formation", "regroupements Secteurs d'Activité", "pays facturation"]

    # Copier les colonnes spécifiques
    for colonne in colonnes_a_copier:
        marge_brute[colonne] = resultat[colonne]

    # Sauvegarder le fichier modifié
    marge_brute.to_excel(chemin_marge_brute, index=False)
else:
    print("Aucune feuille contenant 'BDD' n'a été trouvée dans le fichier.")
