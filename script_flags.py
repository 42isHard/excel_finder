import pandas as pd

# Chemins des fichiers
chemin_source = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/DATA/2023.xlsx'
chemin_sortie = '/home/laptopus/Bureau/SCRIPT_MARGE_BRUT/nouveau_fichier.xlsx'

# Chargement du fichier Excel
df = pd.read_excel(chemin_source)

# Nettoyer les noms de colonnes pour enlever les espaces superflus
df.columns = df.columns.str.strip()


# Fonction pour appliquer les règles et mettre à jour la colonne 'BL'
def categoriser(ligne):
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


# Application des règles directement à la colonne 'BL'
df['Mode de Formation Version BP'] = df.apply(categoriser, axis=1)
print("titi")
# Sauvegarde dans un nouveau fichier Excel
df.to_excel(chemin_sortie, index=False)
