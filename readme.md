# EzFacture

Application de facturation pour Windows. Génère des
factures PDF **Factur-X** à partir de modèles Excel, en embarquant un XML UBL
conforme aux spécifications de l'administration française pour la facturation électronique.

L'application fonctionne entièrement en
**mode local** (numéros de facture stockés sur disque).  

EzFacture garantit une numérotation **chronologique, continue et
sans rupture de séquence** des factures, sans jamais « réserver » un numéro qui pourrait
ensuite être perdu, voir [Numérotation des factures](#numérotation-des-factures).

## Prérequis

- **Windows** (obligatoire — l'app pilote Excel via COM/xlwings).
- **Microsoft Excel** desktop installé, version minimale recommandée : 2021.  
_Peut fonctionner avec des versions plus anciennes, à partir de 2010, mais nécessite dans ce cas des adaptations des formules dans les templates (`RECHERCHEX-> INDEX / EQUIV`)._
- **Python 3.13** (le projet est développé et testé sous 3.13).
- Les fichiers de données (fournis, à adapter) doivent être présents à la racine du projet :
  - `clients.xlsx` — liste des clients
  - `produits.xlsx` — catalogue produits
  - `config.xlsx` — configuration vendeur (SIRET, TVA, adresse…)

## Utilisation
* Télécharger la dernière version de l'éxécutable ici : https://github.com/serialiser/ezfacture/releases/.
* Extraire l'archive sur votre disque dur.  
* Si vous avez des travaux excel en cours, enregistrez vos document et quittez Excel.
* Lancer `ezfacture-1.0.0.exe`.  
* La documentation est disponible ici : https://www.ezfacture.fr/documentation.  


## Numérotation des factures

Les numéros validés sont stockés dans `invoices.jsonl`, un fichier **append-only**
(on n'ajoute qu'à la fin, jamais de modification ni de suppression). Chaque numéro
suit le format `FAC-AAAA-MM-NNNNN` (ou `AV-…` pour un avoir), où `NNNNN` est un
compteur strictement croissant : le prochain numéro est toujours `dernier + 1`.

### 1. Le numéro n'est écrit qu'en dernier, une fois tout le reste réussi

La validation d'une facture est **atomique** (transaction « tout ou rien »,
gérée par `Eztransaction`). Le numéro n'est **inscrit sur disque qu'à l'ultime
étape**, après que toutes les opérations risquées ont abouti :

1. calcul du prochain numéro (`dernier + 1`) — **lecture seule, aucune écriture** ;
2. génération du XML UBL et contrôle contre le schéma XSD ;
3. génération du PDF à partir de la feuille Excel ;
4. embarquement du XML dans le PDF (Factur-X) ;
5. **seulement alors**, écriture définitive du numéro dans `invoices.jsonl`.

Si **une seule** de ces étapes échoue (Excel fermé, XML invalide, erreur PDF…),
la transaction est annulée (rollback) : le numéro **n'a jamais été écrit**. Comme
rien n'a été consommé, **le même numéro sera réattribué à la tentative suivante**.

### 2. Protection contre les modifications

Pour rendre toute altération **détectable**, chaque entrée est protégée par une double chaîne de hachage :

- `file_hash_before` : SHA-256 du fichier **avant** l'ajout de cette entrée
  (chaînage — chaque ligne dépend de tout ce qui précède) ;
- `self_hash` : SHA-256 de `numéro + type + horodatage + file_hash_before`.

À cela s'ajoute un **sceau dans le registre Windows** (nombre d'entrées + dernier
`self_hash`), écrit à chaque validation. Au démarrage, `verify_local_file()`
recalcule toute la chaîne et la compare au sceau. Supprimer, modifier ou réordonner
une ligne casse le chaînage : l'application détecte l'incohérence, affiche une
erreur et **bloque la création de nouveaux documents** tant que le fichier n'est
pas rétabli.

## Développeurs - utilisation à partir des sources (avec un venv)

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
2. copie les dossiers `templates/`, `images/` et `assets/` dans `dist/` ;
3. copie les fichiers de données `clients.xlsx`, `produits.xlsx`, `config.xlsx` ;
4. crée les dossiers de travail `brouillons/` et `pdf/`.

Le résultat se trouve dans `dist/` : lancez l'exécutable
`dist/ezfacture-<version>.exe` (ex. `dist/ezfacture-1.0.0.exe`).

### Option B — PyInstaller manuel

```powershell
pyinstaller main.spec --clean
```

Puis, **manuellement**, copiez à côté de `dist/ezfacture-<version>.exe` :

- les dossiers `templates/`, `images/` et `assets/`
- les fichiers `clients.xlsx`, `produits.xlsx`, `config.xlsx`
- créez les dossiers vides `brouillons/` et `pdf/`

> Sans ces fichiers, l'exécutable démarre mais ne trouve pas ses modèles ni ses
> données. L'option A automatise cette étape.

## Licence

Distribué sous licence Apache, version 2.0. Voir [LICENSE.md](LICENSE.md).

