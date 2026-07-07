from PyInstaller.utils.hooks import collect_data_files

# Inclure tous les fichiers du sous-dossier templates du package en16931
datas = collect_data_files('en16931', subdir='templates')
