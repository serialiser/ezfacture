import win32com.client
import win32print
import winreg
import sys
from pathlib import Path
import logging
import json, os, hashlib
from config import LOCAL_FILE

logger = logging.getLogger(__name__)

_REGISTRY_KEY = r"Software\EzFacture"


def write_registry_seal(count: int, last_self_hash: str) -> None:
    """Écrit le sceau d'intégrité (nb de lignes + dernier self_hash) dans le registre Windows."""
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, _REGISTRY_KEY) as key:
        winreg.SetValueEx(key, "count", 0, winreg.REG_DWORD, count)
        winreg.SetValueEx(key, "last_hash", 0, winreg.REG_SZ, last_self_hash)


def _read_registry_seal():
    """Retourne (count, last_hash) depuis le registre, ou None si absent."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _REGISTRY_KEY) as key:
            count = winreg.QueryValueEx(key, "count")[0]
            last_hash = winreg.QueryValueEx(key, "last_hash")[0]
            return count, last_hash
    except FileNotFoundError:
        return None


def check_format_obj(data_type, valeur):
    """
    :param data_type: obj, type de format à vérifier ex. datetime
    :param valeur: obj, valeur à vérifier
    :return: True ou False
    """
    if not isinstance(valeur, data_type):
        return False
    return True


def check_format_str(data_type_str, valeur_str):
    """
    :param data_type_str: str, type de format à vérifier ex. "format spécifique"
    :param valeur_str: str, valeur à vérifier
    :return: True ou False
    """
    pass


def resource_path(relative_path: str) -> Path:
    # si on est dans un exe PyInstaller
    if getattr(sys, 'frozen', False):
        base_path = Path(sys.executable).resolve().parent
    else:
        # en dev : dossier où se trouve ton script .py
        base_path = Path(__file__).resolve().parent.parent
    
    return (base_path / relative_path).resolve()


def add_pdfa_srgb_output_intent(pdf_path):
    """Ajoute un OutputIntent sRGB (profil ICC embarqué) au PDF pour vaidité PDF/A-3

    Le PDF produit par Excel utilise DeviceRGB /
    DeviceGray sans OutputIntent, ce qui est interdit en PDF/A. En ajoutant un
    OutputIntent sRGB, ces espaces colorimétriques deviennent conformes.

    Sans effet si un OutputIntent est déjà présent.
    """
    try:
        import pikepdf
    except ImportError:
        logger.warning("pikepdf absent : OutputIntent PDF/A non ajouté.")
        return

    icc_path = resource_path("assets/sRGB.icc")
    try:
        icc_bytes = Path(icc_path).read_bytes()
    except OSError as e:
        logger.error(f"Profil ICC sRGB introuvable ({icc_path}) : {e}")
        return

    try:
        with pikepdf.open(str(pdf_path), allow_overwriting_input=True) as pdf:
            existing = pdf.Root.get("/OutputIntents")
            if existing is not None and len(existing) > 0:
                return  # déjà conforme

            icc_stream = pikepdf.Stream(pdf, icc_bytes)
            icc_stream.N = 3  # composants du profil (RGB)

            output_intent = pdf.make_indirect(pikepdf.Dictionary(
                Type=pikepdf.Name.OutputIntent,
                S=pikepdf.Name("/GTS_PDFA1"),
                OutputConditionIdentifier=pikepdf.String("sRGB"),
                Info=pikepdf.String("sRGB IEC61966-2.1"),
                DestOutputProfile=icc_stream,
            ))
            pdf.Root.OutputIntents = pikepdf.Array([output_intent])
            pdf.save(str(pdf_path))
    except Exception as e:
        logger.error(f"Échec de l'ajout de l'OutputIntent PDF/A sur {pdf_path} : {e}")


def close_all_protected_sheets():
    """
    Si Excel a déjà un / des fichier(s) ouvert en mode protégé au lancement de l'app, xlwings renvoie un objet ProtectedViewWindow au lieu d'un objet Book et on a des erreurs.
    On doit fermer ces fichiers avant de faire autre chose. 
    """
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
    except BaseException:
        return

    try:
        pv = excel.ProtectedViewWindows
        initial_count = pv.Count
        if initial_count == 0:
            return

        # parcourir en sens inverse (les index des éléments sont modifiés)
        for i in range(pv.Count, 0, -1):
            try:
                pv.Item(i).Close()
            except BaseException as e:
                logger.error(f"Impossible de fermer le fichier en mode protégé : {e}")

    except BaseException as e:
        logger.error(f"Erreur lors de la gestion des ProtectedViewWindows : {e}")
        return

def cleanup_draft_pdfs(pdf_dir: str = "pdf") -> None:
    """Supprime les fichiers DRAFT_* présents dans le répertoire pdf au démarrage."""
    pdf_path = Path(pdf_dir)
    if not pdf_path.exists():
        return
    for f in pdf_path.iterdir():
        if f.is_file() and f.name.startswith("DRAFT_"):
            try:
                f.unlink()
            except Exception as e:
                logger.error(f"Impossible de supprimer le brouillon pdf {f} : {e}")


def sha256(x: str) -> str:
    return hashlib.sha256(x.encode()).hexdigest()

def load_last_entry():
    """Renvoie le dernier numéro de facture + file_hash_before/self_hash de la ligne."""
    if not os.path.exists(LOCAL_FILE):
        return None, "0" * 64

    last_line = None
    with open(LOCAL_FILE, "r", encoding="utf8") as f:
        for line in f:
            last_line = line.strip()

    if not last_line:
        return None, "0" * 64

    entry = json.loads(last_line)
    return entry["number"], entry["self_hash"]


def load_last_entry_by_type(type_doc: str):
    """Renvoie le dernier numéro pour un type de document donné (facture, avoir...)."""
    if not os.path.exists(LOCAL_FILE):
        return None

    last_number = None
    with open(LOCAL_FILE, "r", encoding="utf8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            entry = json.loads(line)
            if entry.get("type") == type_doc:
                last_number = entry["number"]

    return last_number


def sha256_bytes(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def compute_file_hash(path: str) -> str:
    """Hash de tout le fichier JSONL. Hash des bytes vides si absent."""
    if not os.path.exists(path):
        return sha256_bytes(b"")
    with open(path, "rb") as f:
        return sha256_bytes(f.read())
    

def verify_local_file(path: str = LOCAL_FILE) -> bool:
    """
    Vérifie l'intégrité complète du fichier JSONL utilisant :
    - file_hash_before : hash cumulatif du fichier avant l'entrée
    - self_hash        : hash des champs de l'entrée
    - sceau registre   : count + last_hash pour détecter la suppression de la dernière ligne
    Retourne True si tout est valide, False sinon.
    """
    seal = _read_registry_seal()

    if not os.path.exists(path):
        if seal is not None and seal[0] > 0:
            logger.error(f"invoices.jsonl absent mais le registre indique {seal[0]} entrée(s).")
            return False
        return True

    try:
        with open(path, "rb") as f:
            full_data = f.read()
    except Exception as e:
        logger.error(f"Impossible de lire le fichier : {e}")
        return False

    offset = 0
    line_count = 0
    last_self_hash = None

    # splitlines(keepends=True) préserve les octets réels (\r\n ou \n)
    # ce qui garantit que offset reflète la position exacte dans full_data
    for line_idx, line_b in enumerate(full_data.splitlines(keepends=True)):
        line_stripped = line_b.decode("utf8").strip()

        expected_file_hash_before = sha256_bytes(full_data[:offset])

        try:
            entry = json.loads(line_stripped)
        except:
            logger.error(f"Ligne {line_idx+1} non valide JSON.")
            return False

        if entry.get("file_hash_before") != expected_file_hash_before:
            logger.error(
                f"Ligne {line_idx+1} : file_hash_before incorrect - attendu: {expected_file_hash_before} - trouvé : {entry.get('file_hash_before')}"
            )
            return False

        raw = (
            entry.get("number", "")
            + entry.get("type", "")
            + entry.get("timestamp", "")
            + entry.get("file_hash_before", "")
        )
        expected_self_hash = sha256(raw)

        if entry.get("self_hash") != expected_self_hash:
            logger.error(
                f"Ligne {line_idx+1} : self_hash incorrect - attendu: {expected_self_hash} - trouvé : {entry.get('self_hash')}"
            )
            return False

        offset += len(line_b)
        line_count += 1
        last_self_hash = entry["self_hash"]

    if seal is None:
        # Première exécution ou migration : initialiser le sceau sans bloquer
        write_registry_seal(line_count, last_self_hash or "0" * 64)
    else:
        sealed_count, sealed_hash = seal
        if line_count != sealed_count:
            logger.error(f"Sceau registre : {sealed_count} entrée(s) attendue(s), {line_count} trouvée(s).")
            return False
        if last_self_hash != sealed_hash:
            logger.error("Sceau registre : dernier self_hash différent.")
            return False

    return True

def disable_win_auto_printer_manage():
    key_path = r"Software\Microsoft\Windows NT\CurrentVersion\Windows"

    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        key_path,
        0,
        winreg.KEY_SET_VALUE
    ) as key:

        # 1 = désactive la gestion automatique
        winreg.SetValueEx(
            key,
            "LegacyDefaultPrinterMode",
            0,
            winreg.REG_DWORD,
            1
        )

def get_microsoft_print_to_pdf_printer():
    """
    Retourne le nom de l'imprimante utilisant le driver
    'Microsoft Print To PDF'.

    Retour :
        str | None
    """
    printers = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_LOCAL |
        win32print.PRINTER_ENUM_CONNECTIONS
    )

    for printer in printers:
        _, _, printer_name, _ = printer

        try:
            hprinter = win32print.OpenPrinter(printer_name)
            info = win32print.GetPrinter(hprinter, 2)

            driver_name = info["pDriverName"]

            win32print.ClosePrinter(hprinter)

            if driver_name == "Microsoft Print To PDF":
                return printer_name

        except Exception:
            pass

    return None

def get_default_printer_name():
    return win32print.GetDefaultPrinter()

def set_default_printer(printer_name):
    """Définit l'imprimante par défaut Windows sur l'imprimante pdf.
    Retourne True si l'opération a réussi, False sinon.
    """

    try:
        win32print.SetDefaultPrinter(printer_name)
        return True
    except Exception as e:
        return False
    
