"""
Script de build et packaging pour ez_facture.
Usage : python build.py
"""
import subprocess
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).parent
DIST = ROOT / "dist"

DATA_FILES = ["clients.xlsx", "produits.xlsx", "config.xlsx"]
FOLDERS_TO_COPY = ["templates", "images", "assets"]
DIRS_TO_CREATE = ["brouillons", "pdf"]


def run_pyinstaller():
    print("==> PyInstaller...")
    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "main.spec", "--clean"],
        cwd=ROOT
    )
    if result.returncode != 0:
        print("ERREUR : PyInstaller a échoué.")
        sys.exit(1)


def copy_folders():
    print("==> Copie des dossiers...")
    for dirname in FOLDERS_TO_COPY:
        src = ROOT / dirname
        dst = DIST / dirname
        if src.exists():
            if dst.exists():
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
            print(f"    {dirname}/ copié")
        else:
            print(f"    {dirname}/ introuvable, ignoré")


def copy_data_files():
    print("==> Copie des fichiers de données...")
    for filename in DATA_FILES:
        src = ROOT / filename
        if src.exists():
            shutil.copy2(src, DIST / filename)
            print(f"    {filename} copié")
        else:
            print(f"    {filename} introuvable, ignoré")


def create_dirs():
    print("==> Création des dossiers...")
    for dirname in DIRS_TO_CREATE:
        target = DIST / dirname
        target.mkdir(exist_ok=True)
        print(f"    {dirname}/ créé")


if __name__ == "__main__":
    run_pyinstaller()
    copy_folders()
    copy_data_files()
    create_dirs()
    print("\nBuild terminé. Contenu de dist/:")
    for item in sorted(DIST.iterdir()):
        print(f"    {item.name}")
