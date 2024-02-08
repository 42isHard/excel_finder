import subprocess
from tqdm import tqdm
import time


def run_script(script_path, message):
    """
    Exécute un script Python situé au chemin spécifié et affiche un message approprié.
    """
    print(message)
    time.sleep(1)  # Petite pause pour lire le message
    try:
        subprocess.run(["python", script_path], check=True)
        print(f"✅ Exécution réussie : {script_path}")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erreur lors de l'exécution du script {script_path}: {e}")
    time.sleep(1)  # Petite pause après l'exécution


# Chemin du dossier contenant vos scripts
base_path = "/home/laptopus/PycharmProjects/Excel_automation"

# Liste des chemins de vos scripts
scripts = [
    f"{base_path}/script_CONCAT.py",
    f"{base_path}/script_CROSS.py",
    f"{base_path}/script_LABEL.py",
]

# Messages professionnels pour chaque script
messages = [
    "Démarrage du processus de concaténation des données.",
    "Début de l'opération de croisement des données.",
    "Lancement du processus de catégorisation des données."
]

# Exécution séquentielle des scripts avec barre de progression
for script, message in zip(tqdm(scripts, desc="Progression globale", unit="script"), messages):
    run_script(script, message)

print("Processus terminé. Tous les scripts ont été exécutés avec succès.")
