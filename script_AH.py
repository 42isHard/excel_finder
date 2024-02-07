import pandas as pd

# Chemins des fichiers
chemin_source = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/2023.xlsx'
chemin_sortie = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/nouveau_fichier.xlsx'

# Chargement du fichier Excel
df = pd.read_excel(chemin_source)

# Nettoyer les noms de colonnes pour enlever les espaces superflus
df.columns = df.columns.str.strip()


def categoriser(ligne):
    """
    Fonction pour catégoriser la valeur de la colonne 'BL' selon les règles définies.

    Args:
        ligne: Une ligne du DataFrame Pandas.

    Returns:
        La catégorie de la valeur 'BL'.
    """
    bl = ligne['BL'].strip()  # Enlève les espaces superflus de la valeur de 'BL'
    mode_formation = ligne['mode de formation'].strip()  # Enlève les espaces superflus de 'mode de formation'

    # Applique les règles selon les conditions fournies
    if bl == "Regroupement Réglementaire" and mode_formation in ["Executive MOOC", "Corporate MOOC"]:
        return "REGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation in ["Executive MOOC", "Corporate MOOC"]:
        return "AUTRES DISTANCIELS"
    elif bl != "Regroupement Réglementaire" and mode_formation == "In house training":
        return "INTRA"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Public training":
        return "INTER"
    elif mode_formation in ["Skills First INTER", "Skills First INTRA"]:
        return "SKILL FIRST"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Online Certificate":
        return "EOC"
    else:
        return bl  # Retourne la valeur existante si aucune règle ne s'applique


def categoriser_nouvelle_bp(ligne):
    """
    Fonction pour catégoriser la valeur de la nouvelle colonne 'Mode de Formation Version BP'.

    Args:
        ligne: Une ligne du DataFrame Pandas.

    Returns:
        La catégorie de la valeur de la nouvelle colonne 'Mode de Formation Version BP'.
    """
    mode_formation_bp = ligne['Mode de Formation Version BP'].strip()
    mode_formation = ligne['mode de formation'].strip()
    titre_produit = ligne['Titre du produit'].lower()  # Convertir en minuscules pour la recherche
    bl = ligne['BL'].strip()

    # Appliquer les nouvelles règles
    if mode_formation_bp == "Réglementaire" and mode_formation in ["Corporate MOOC", "Executive MOOC"]:
        return "DL REGLEMENTAIRE"
    elif mode_formation_bp == "Réglementaire" and mode_formation == "In house training":
        return "INTRA PRÉSENTIEL RÉGLEMENTAIRE"
    elif mode_formation_bp == "Regroupement Réglementaire" and mode_formation == "Public training":
        return "INTER PRÉSENTIEL RÉGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation == "Public training":
        return "INTER PRÉSENTIEL HORS REGLEMENTAIRE"
    elif bl != "Regroupement Réglementaire" and mode_formation == "In house training":
        return "INTRA PRÉSENTIEL HORS REGLEMENTAIRE"
    elif mode_formation_bp == "EOC" and any(x in titre_produit for x in ["hec", "cbs", "financ"]):
        return "FEO FINANCE"
    elif mode_formation_bp == "EOC" and not any(x in titre_produit for x in ["hec", "cbs", "financ"]):
        return "FEO NON FINANCE"
    else:
        return "Autre"


# Application de la première règle directement à la colonne 'BL'
df['Mode de Formation Version BP'] = df.apply(categoriser, axis=1)

# Application de la deuxième règle pour créer la nouvelle colonne temporairement
df['Mode de formation new BP Temp'] = df.apply(categoriser_nouvelle_bp, axis=1)

# Trouver l'index de la colonne 'Mode de Formation Version BP'
index_mode_formation_version_bp = df.columns.get_loc('Mode de Formation Version BP')

# Insérer la colonne 'Mode de formation new BP' à l'index désiré
df.insert(index_mode_formation_version_bp + 1, 'Mode de formation new BP', df['Mode de formation new BP Temp'])

# Supprimer la colonne temporaire
df.drop('Mode de formation new BP Temp', axis=1, inplace=True)

# Sauvegarde dans un nouveau fichier Excel
df.to_excel(chemin_sortie, index=False)
