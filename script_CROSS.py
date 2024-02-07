import os
import pandas as pd
from tqdm import tqdm

# Définir le chemin complet des fichiers
path_2023 = '/home/laptopus/Documents/SOURCES_EXCEL/2023.xlsx'
path_sources = '/home/laptopus/Documents/SOURCES_EXCEL/sources.xlsx'

# Charger les données depuis le fichier 2023.xlsx
print("Chargement des données depuis le fichier 2023.xlsx...")
df_2023 = pd.read_excel(path_2023)

# Initialiser un DataFrame vide pour stocker les résultats
result = pd.DataFrame()

# Convertir la colonne 'Section Analytique - Code' du fichier 2023.xlsx en chaîne
print("Conversion de la colonne 'Section Analytique - Code' en chaînes de caractères...")
df_2023['Section Analytique - Code'] = df_2023['Section Analytique - Code'].astype(str)

# Charger les données depuis le fichier sources.xlsx
print("Chargement des données depuis le fichier sources.xlsx...")
df_sources = pd.read_excel(path_sources, sheet_name=None)

# Initialiser la barre de progression
total_sheets = len(df_sources)
progress_bar = tqdm(total=total_sheets, desc='Traitement des feuilles')

# Parcourir chaque feuille du fichier sources.xlsx
for sheet_name, df_source in df_sources.items():
    print(f"\nTraitement de la feuille '{sheet_name}'...")

    # Convertir chaque colonne de la feuille en chaînes de caractères
    for column_name in df_source.columns:
        print(f"   Conversion de la colonne '{column_name}' en chaînes de caractères...")

        # Filtrer les lignes de df_2023 qui correspondent aux valeurs de df_source
        matching_rows = df_2023[df_2023['Section Analytique - Code'].isin(df_source[column_name])]

        # Afficher les données ajoutées à chaque ligne
        for index, row in matching_rows.iterrows():
            print(f"      Ajout de la ligne {index} - TROUVÉ")

        # Ajouter les résultats au DataFrame global
        result = pd.concat([result, matching_rows], ignore_index=True)

    # Mettre à jour la barre de progression
    progress_bar.update(1)

# Fermer la barre de progression
progress_bar.close()

# Enregistrer le résultat dans un nouveau fichier Excel
output_path = '/home/laptopus/Documents/SOURCES_EXCEL/output.xlsx'
result.to_excel(output_path, index=False)

print(f"\nLe résultat a été enregistré dans le fichier {output_path}.")
