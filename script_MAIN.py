import subprocess
from tqdm import tqdm
import time


def run_script(script_path):
    """
    Exécute un script Python situé au chemin spécifié.
    """
    try:
        subprocess.run(["python", script_path], check=True)
        print(f"✅ Script exécuté avec succès : {script_path}")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erreur lors de l'exécution du script {script_path}: {e}")


# Chemin du dossier contenant vos scripts
base_path = "/home/laptopus/PycharmProjects/Excel_automation"

# Liste des chemins de vos scripts
scripts = [
    f"{base_path}/script_CONCAT.py",
    f"{base_path}/script_CROSS.py",
    f"{base_path}/script_LABEL.py",
]

# Messages encourageants
encouragements = [
    "🚀 Lancement du premier script, c'est parti !",
    "🔥 Deuxième script en cours, super travail !",
    "✨ Dernier script, presque terminé !"
]

# Exécution séquentielle des scripts avec barre de progression et messages
for i, script in enumerate(tqdm(scripts, desc="Progression globale", unit="script")):
    print(encouragements[i])
    run_script(script)
    time.sleep(1)  # Petite pause pour l'affichage

print("🎉 Tous les scripts ont été exécutés avec succès !")
