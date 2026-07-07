# EzFacture

Application de facturation pour Windows. Génère des
factures PDF **Factur-X** à partir de modèles Excel, en embarquant un XML UBL
conforme aux spécifications de l'administration française pour la facturation électronique.

L'application fonctionne entièrement en
**mode local** (numéros de facture stockés sur disque).

## Prérequis

- **Windows** (obligatoire — l'app pilote Excel via COM/xlwings).
- **Microsoft Excel** desktop installé, version minimale recommandée : 2021.  
_Peut fonctionner avec des versions plus anciennes, à partir de 2010, mais nécessite dans ce cas des adaptations des formules dans les templates (`RECHERCHEX-> INDEX / EQUIV`)._
- **Python 3.13** (le projet est développé et testé sous 3.13).
- Les fichiers de données (fournis, à adapter) doivent être présents à la racine du projet :
  - `clients.xlsx` — liste des clients
  - `produits.xlsx` — catalogue produits
  - `config.xlsx` — configuration vendeur (SIRET, TVA, adresse…)

## Utilisation à partir des sources (avec un venv)

### 1. Cloner le dépôt

```powershell
git clone <url-du-depot> ezfacture
cd ezfacture
```

### 2. Créer et activer l'environnement virtuel

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

> Si l'activation est bloquée par la politique d'exécution PowerShell :
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` (une seule fois),
> puis réessayez. En invite `cmd`, utilisez `venv\Scripts\activate.bat`.

### 3. Installer les dépendances

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

Pour activer **en plus** le mode API (nécessite le sous-module privé) :

```powershell
pip install -r controller/backends/api/requirements.txt
```


### 5. Lancer l'application

Depuis la racine du projet, avec le venv activé :

```powershell
python main.py
```

### Lancer les tests (optionnel)

```powershell
pytest en16931/tests/
```

Les tests couvrent uniquement le module `en16931/` (le reste requiert Excel COM).

## Compilation en exécutable (PyInstaller)

Le build utilise [main.spec](main.spec), qui gère les imports cachés
(`xlwings`, `tzdata`), les données de la bibliothèque `factur-x` et le logo.
`pyinstaller` est déjà inclus dans `requirements.txt`.

### Option A — script de build automatisé (recommandé)

[build.py](build.py) lance PyInstaller **et** copie les fichiers nécessaires
à côté de l'exécutable :

```powershell
python build.py
```

Le script :
1. exécute `pyinstaller main.spec --clean` ;
2. copie les dossiers `templates/` et `images/` dans `dist/` ;
3. copie les fichiers de données `clients.xlsx`, `produits.xlsx`, `config.xlsx` ;
4. crée les dossiers de travail `brouillons/` et `pdf/`.

Le résultat se trouve dans `dist/` : lancez `dist/main.exe`.

### Option B — PyInstaller manuel

```powershell
pyinstaller main.spec --clean
```

Puis, **manuellement**, copiez à côté de `dist/main.exe` :

- les dossiers `templates/` et `images/`
- les fichiers `clients.xlsx`, `produits.xlsx`, `config.xlsx`
- créez les dossiers vides `brouillons/` et `pdf/`

> Sans ces fichiers, l'exécutable démarre mais ne trouve pas ses modèles ni ses
> données. L'option A automatise cette étape.


