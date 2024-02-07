import pandas as pd

# Chemins des fichiers
chemin_source = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/2023.xlsx'
chemin_sortie = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/nouveau_fichier.xlsx'

# Chargement du fichier Excel
df = pd.read_excel(chemin_source)

# Nettoyer les noms de colonnes pour enlever les espaces superflus
df.columns = df.columns.str.strip()

# Fonction pour appliquer les nouvelles règles
def categoriser_nouvelle_bp(ligne):
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

# Appliquer la fonction et créer la nouvelle colonne temporairement
df['Mode de formation new BP Temp'] = df.apply(categoriser_nouvelle_bp, axis=1)

# Trouver l'index de la colonne 'Mode de Formation Version BP'
index_mode_formation_version_bp = df.columns.get_loc('Mode de Formation Version BP')

# Insérer la colonne 'Mode de formation new BP' à l'index désiré
df.insert(index_mode_formation_version_bp + 1, 'Mode de formation new BP', df['Mode de formation new BP Temp'])

# Supprimer la colonne temporaire
df.drop('Mode de formation new BP Temp', axis=1, inplace=True)

# Sauvegarde dans un nouveau fichier Excel
df.to_excel(chemin_sortie, index=False)
