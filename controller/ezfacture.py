import logging
import pywintypes

import xlwings as xw
from pathlib import Path
import shutil
from datetime import datetime
from controller.constantes import NAMES_FACT, NAMES_DEVIS, NAMES_AVOIR, DRAFT_PATH
from tools.utils import add_pdfa_srgb_output_intent

logger = logging.getLogger(__name__)


class EzfactureAlreadyExistsError(Exception):
    """
    Si on essaie de créer une deuxième instance de Ezfacture.
    """
    def __init__(self, message=None):
        if message is None:
            message = (
                "Plusieurs documents ne peuvent pas être édités en même temps.\n"
                "Pour créer un nouveau document, veuillez enregistrer et fermer "
                "le document en cours."
            )
        super().__init__(message)


class Ezfacture:
    """
    wb : workbook (classeur excel)
    """
    _instance = None
    type_modele_dict = {
        "Devis": "devis",
        "Facture": "facture",
        "Avoir": "avoir"
    }

    def __new__(cls, *args, **kwargs):
        """
        On ne peut créer qu'une seule instance
        """
        if cls._instance is not None:
            raise EzfactureAlreadyExistsError()
        cls._instance = super().__new__(cls)
        return cls._instance


    def __init__(self, type_doc, new_doc=True):
        """
        On crée une nouvelle instance quand on crée un nouveau doc à partir du template (facture, devis, avoir),
        ou quand on ouvre un fichier excel existant (brouillon ou devis).
        """
        self.wb = None
        self.type_doc = type_doc
        self.new_doc = new_doc
        self.state = 'Brouillon'
        self.doc_path = ""
        self.pdf_path = ""
        self.doc_name = ""
        self.onglet = None
        self.cell_date_name = ""  # nom de la cellule qui contient la date du document
        self.cell_numdoc_name = ""  # nom de la cellule qui contient le numéro de document
        self.cell_num_doc = None  # Cellule qui contient le numéro de document (class 'xlwings.main.Range')
        self.pdf_facture_num = ""  # Nom du fichier pdf

        if new_doc:
            prefixe = (
                "DRAFT_FAC_" if type_doc == "facture"
                else "DEV_" if type_doc == "devis"
                else "DRAFT_AV_" if type_doc == "avoir"
                else ""
            )

            self.doc_name = prefixe + str(int(datetime.now().timestamp()))
            self.doc_path = self.get_doc_path()
            shutil.copy(self.get_src(), self.doc_path)

            self.wb = xw.Book(self.doc_path)

            try:
                self.onglet = self.wb.sheets[type_doc]
                self.onglet_config = self.wb.sheets["config"]
            except BaseException as e:
                raise ValueError(f"Erreur avec l'onglet : {e}")
            self.onglet.activate()

            # Supprimer les autres onglets
            for val in Ezfacture.type_modele_dict.values():
                if val == type_doc:
                    continue
                self.wb.sheets[val].delete()

        else:
            # Implémentation via le controleur
            pass

    def close(self):
        if self.wb is not None:
            self.wb.close()
            self.wb = None
        type(self)._instance = None

    def check_cells(self):
        """
        Vérifie que l'onglet contient bien les cellules nommées nécessaires au fonctionnement
        :return: liste de noms manquants ou liste vide si ok
        """
        defined_names = [name.name for name in self.onglet.names]  # names contient une liste d'objets Name
        names_to_check = (
            NAMES_FACT if self.type_doc == 'facture' else
            NAMES_DEVIS if self.type_doc == 'devis' else
            NAMES_AVOIR if self.type_doc == 'avoir' else
            []
        )
        results = {name: (name in defined_names) for name in names_to_check}
        missing_names = []
        for name, exists in results.items():
            if not exists:
                missing_names.append(name)
        return missing_names

    @staticmethod
    def get_src():
        """
        :return: raw string du chemin du template
        """
        if Path('./templates/template.xlsx').exists():
            return r'templates/template.xlsx'
        else:
            raise FileNotFoundError(f"Le fichier n'existe pas.")

    def get_doc_path(self):
        """
        :return: raw string du chemin du brouillon de document
        """
        chemin_brouillons = Path(DRAFT_PATH)
        chemin_brouillons.mkdir(parents=True, exist_ok=True)
        
        return chemin_brouillons / f"{self.doc_name}.xlsx"

    def cell_error(self, onglet, cell_address, error_message):
        onglet[cell_address].font.color = "#cc0000"
        if onglet[cell_address].note:
            onglet[cell_address].note.text = error_message
        else:
            # pas d'autre méthode pour créer une note (= commentaire), on utilise donc cette syntaxe
            onglet[cell_address].api.AddComment(error_message)
        self.wb.save()

    def make_pdf(self, path=None):
        """
        Enregistre l'onglet en pdf
        :param path: optionnel, nom de dossier
        """
        if self.cell_num_doc.value:
            self.pdf_facture_num = self.cell_num_doc.value + '.pdf'
        else:
            # Ceci peut se produire si un nouveau document vient d'être créé, et excel est fermé sans enregistrer le doc -> pas de numéro de doc.
            # Il existe ensuite dans le dossier /brouillons un doc sans numéro de doc.
            raise ValueError("Le numéro de facture ne peut pas être vide.")
        if path:
            pdf_path = Path(path)
            pdf_path.mkdir(parents=True, exist_ok=True)
        else:
            pdf_path = Path('.')
        fact_path = pdf_path / self.pdf_facture_num

        try:
            self.onglet.to_pdf(path=fact_path, show=False)  # Attention ne pas ouvrir le pdf ici, sinon on ne peut pas le supprimer si rollback
            # Ajout d'un OutputIntent sRGB pour la conformité PDF/A-3 (Factur-X)
            add_pdfa_srgb_output_intent(fact_path)
            # on stocke le chemin du pdf pour pouvoir facilement créer ensuite le doc facturX
            self.pdf_path = fact_path

        except pywintypes.com_error as e:
            logger.error(f"Il semble que l'imprimante n'est pas connectée : {e}.")

        except BaseException as e:
            logger.exception(e)
            raise ValueError(e)

    def delete_draft(self):
        """Supprime le fichier brouillon Excel après validation réussie."""
        draft_path = Path(self.doc_path)
        if draft_path.exists():
            try:
                draft_path.unlink()
            except Exception as e:
                logger.error(f"Impossible de supprimer le brouillon {draft_path} : {e}")

    def unmake_pdf(self):
        """
        Annule la création du PDF en supprimant le fichier existant.
        Ne fait rien si aucun PDF n'a été créé.
        """
        pdf_path = getattr(self, "pdf_path", None)

        if not pdf_path:
            return

        try:
            pdf_file = Path(pdf_path)

            if pdf_file.exists():
                pdf_file.unlink()
        except BaseException as e:
            raise ValueError(f"Impossible de supprimer le PDF : {e}")

        # On annule aussi les attributs liés
        self.pdf_path = None
        self.pdf_facture_num = None

    def is_wb_open(self):
        """Si on peut accéder à l'attribut sheets du workbook alors il est ouvert"""
        try:
            _ = self.wb.sheets
            return True
        except BaseException as e:
            return False
