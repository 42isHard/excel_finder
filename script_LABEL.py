import pandas as pd
from tqdm import tqdm

# Chemins des fichiers
chemin_source = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/SORTIE/marge_brute_CROSS.csv'
chemin_sortie = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/marge_brute_FINALE.csv'


def nettoyer_noms_colonnes(df):
    """Nettoie les noms de colonnes pour enlever les espaces superflus."""
    df.columns = df.columns.str.strip()
    return df


def categoriser(ligne):
    """Catégorise les données selon la première règle."""
    bl = str(ligne['BL']).strip() if pd.notna(ligne['BL']) else ""
    mode_formation = str(ligne['mode Formation']).strip() if pd.notna(ligne['mode Formation']) else ""
    if bl == "Regroupement Réglementaire" and mode_formation in ["Corporate MOOC", "Executive MOOC"]:
        return "DL REGLEMENTAIRE"
    elif bl == "Regroupement Réglementaire" and mode_formation in ["In house training", "Public training"]:
        return "PRÉSENTIEL RÉGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Public training":
        return "INTER PRÉSENTIEL HORS REGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation == "In house training":
        return "INTRA PRÉSENTIEL HORS REGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Online Certificate" and any(
            x in ligne['Titre du produit'] for x in ["hec", "cbs", "financ"]):
        return "FEO FINANCE"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Online Certificate" and not any(
            x in ligne['Titre du produit'] for x in ["hec", "cbs", "financ"]):
        return "FEO NON FINANCE"
    else:
        return "AUTRE DL"


def categoriser_nouvelle_bp(ligne):
    """Catégorise les données selon la deuxième règle."""
    bl = str(ligne['BL']).strip() if pd.notna(ligne['BL']) else ""
    mode_formation = str(ligne['mode Formation']).strip() if pd.notna(ligne['mode Formation']) else ""
    if bl == "Regroupement Réglementaire" and mode_formation in ["Corporate MOOC", "Executive MOOC"]:
        return "REGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation in ["Corporate MOOC", "Executive MOOC"]:
        return "AUTRES DISTANCIELS"
    elif mode_formation == "In house training":
        return "INTRA"
    elif mode_formation == "Public training":
        return "INTER"
    elif mode_formation in ["Skills First INTER", "Skills First INTRA"]:
        return "SKILL FIRST"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Online Certificate":
        return "EOC"
    else:
        return bl


# Chargement du fichier CSV
df = pd.read_csv(chemin_source)

# Nettoyer les noms de colonnes
df = nettoyer_noms_colonnes(df)

# Application de la première règle directement à la colonne 'BL' avec barre de progression
tqdm.pandas(desc="Catégorisation 1")
df['Mode de Formation Version BP'] = df.progress_apply(categoriser, axis=1)

# Application de la deuxième règle pour créer la nouvelle colonne temporairement avec barre de progression
tqdm.pandas(desc="Catégorisation 2")
df['Mode de formation new BP Temp'] = df.progress_apply(categoriser_nouvelle_bp, axis=1)

# Trouver l'index de la colonne 'Mode de Formation Version BP'
index_mode_formation_version_bp = df.columns.get_loc('Mode de Formation Version BP')

# Insérer la colonne 'Mode de formation new BP' à l'index désiré
df.insert(index_mode_formation_version_bp + 1, 'Mode de formation new BP', df['Mode de formation new BP Temp'])

# Supprimer la colonne temporaire
df.drop('Mode de formation new BP Temp', axis=1, inplace=True)

# Sauvegarde dans un nouveau fichier CSV
df.to_csv(chemin_sortie, index=False)
