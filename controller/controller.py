import os
import sys
import gc
import webbrowser
from pathlib import Path
from threading import Thread
import xlwings as xw
import pywintypes
import pandas as pd
from datetime import datetime
import logging
import json
from facturx import generate_from_file, xml_check_xsd
from tools.utils import check_format_obj, resource_path, close_all_protected_sheets, disable_win_auto_printer_manage, set_default_printer, get_microsoft_print_to_pdf_printer, cleanup_draft_pdfs
from controller import Ezfacture, EzfactureAlreadyExistsError, Eztransaction
from controller.backends import get_numbering_backend
from controller.constantes import DRAFT_PATH
from models import Facture, Entity, InvoiceLine, Devis
from en16931.tax import FR_FRANCHISE_EN_BASE

from controller.constantes import RANGE_PRINT_AREA

logger = logging.getLogger(__name__)


def check_ui(f):
    """
    Décorateur à utiliser avant une action sur un doc, pour vérifier qu'excel n'a pas été fermé en-dehors de l'app ezfacture.
    """
    def wrapper(self, *args, **kwargs):
        """
        Le wrapper est utilisé sur des méthodes de la classe Controller
        donc on lui passe le paramètre self
        """
        if self.doc.is_wb_open():
            return f(self, *args, **kwargs)
        else:
            self.view.block_ui()
            self.view.show_feedback(txt="L'action ne peut pas être réalisée car le document a été fermé.\n"
                                    "Merci de redémarrer et réouvrir le document.",
                                    message_type="error")
            self.view.show_restart()
            self.view.set_actions({"restart": self.restart})
            return None
    return wrapper


def check_opened_doc(f):
    """
    Décorateur à utiliser avant de créer ou d'ouvrir un document.
    """
    def wrapper(self, *args, **kwargs):

        # On s'assure qu'il n'y a pas déjà un doc excel ouvert
        if self.doc and self.doc.is_wb_open():
            is_closed = self.fermer()
            
            if is_closed:
                # S'il existe une instance Ezfacture il faut la détruire
                if self.doc:
                    type(self.doc)._instance = None
                    gc.collect()

        # s'il existe une instance Ezfacture active, excel ayant été fermé en-dehors de l'app
        elif self.doc:
            type(self.doc)._instance = None
            gc.collect()

        return f(self, *args, **kwargs)
    return wrapper


class Controller:

    def __init__(self, model, view):
        self.doc = None  # Instance Ezfacture
        self.model = model
        self.view = view
        self.backend = get_numbering_backend(self)  # LocalBackend ou ApiBackend
        self.template = ""
        self.file_name = ""
        self.cell_date_name = ""  # Nom de l'onglet date en fonction du template

        # On relie les commandes des boutons et menus (view) aux fonctions du controleur
        actions = {
            "connexion": self.connexion,
            "save": self.save_draft_and_pdf,
            "valider": self.validate,
            "transformer": self.transform,
            "aide": self.aide,
        }
        self.view.set_actions(actions)
        self.view.set_menu_nouveau(self.create_doc, [key for key in Ezfacture.type_modele_dict])
        self.view.set_menu_ouvrir(self.ouvrir, self.load_docs_paths())

        close_all_protected_sheets()
        cleanup_draft_pdfs()

        # On désactive "Laisser Windows gérer mon imprimante par défaut"
        # Sinon erreurs pour générer le pdf depuis excel
        disable_win_auto_printer_manage()

    @check_opened_doc
    def create_doc(self, choice, numClient=None, lignes=None):

        try:
            self.template = Ezfacture.type_modele_dict.get(choice)
            self.doc = Ezfacture(self.template)
            self.view.delete_messages("feedback")

            self.init_doc(numClient=numClient, lignes=lignes)

        except EzfactureAlreadyExistsError as e:
            self.view.show_feedback(txt=str(e), message_type="error")

    @check_opened_doc
    def ouvrir(self, choice):
        """
        Crée une instance Ezfacture et ouvre un fichier excel existant
        """

        type_doc = self.get_type_doc(choice)

        # on réinitialise le bouton ouvrir sinon il contient le chemin du doc (trop long)
        self.view.reset_menu_ouvrir("Ouvrir")
        
        try:
            self.doc = Ezfacture(type_doc, new_doc=False)
            self.doc.doc_path = choice
            self.doc.wb = xw.Book(self.doc.doc_path)
            self.view.delete_messages("feedback")

        except EzfactureAlreadyExistsError as e:
            self.view.show_feedback(txt=str(e), message_type="error")
            return

        try:
            self.doc.onglet = self.doc.wb.sheets[type_doc]
            self.doc.onglet_config = self.doc.wb.sheets["config"]
        except ValueError as e:
            raise ValueError(f"Erreur avec l'onglet : {e}")
        
        self.doc.onglet.activate()

        self.init_doc()

    def fermer(self):
        fermer_doc = self.view.open_alert_close("Enregistrer le document et fermer ?")
        if fermer_doc:
            if self.doc is not None:
                self.save_draft()
                self.doc.close()
                self.doc = None
                gc.collect()

            self.view.enable_ui()
            
            # On recharge le menu Ouvrir afin que le document fermé s'y trouve ensuite
            self.view.set_menu_ouvrir(self.ouvrir, self.load_docs_paths())
        return fermer_doc


    def init_doc(self, numClient=None, lignes=None):

        self.load_data()

        self.model = (
            Facture(invoice_type_code='380') if self.doc.type_doc == 'facture' else
            Facture(invoice_type_code='381') if self.doc.type_doc == 'avoir' else
            Devis() if self.doc.type_doc == 'devis' else
            ''
        )

        # ------------------- DEVIS ------------------ #
        if self.doc.type_doc == 'devis':
            self.view.buttons["valider"].configure(text="Transformer en facture")
            self.view.buttons["valider"].configure(command=self.transform)
        else:
            self.view.buttons["valider"].configure(text="Valider le document")
            self.view.buttons["valider"].configure(command=self.validate)
        # -------------------------------------------- #


        self.doc.cell_date_name = (
            'date_facture' if self.doc.type_doc in ('facture', 'avoir') else
            'dev_date_devis' if self.doc.type_doc == 'devis' else
            ''
        )

        NOM_CELL_NUM_DOC = {
            'facture': 'fact_num_fact',
            'devis': 'dev_num_devis',
            'avoir': 'av_num_avoir'
        }

        # Nom de la cellule qui contient le numéro de doc
        self.doc.cell_numdoc_name = NOM_CELL_NUM_DOC.get(self.doc.type_doc, '')

        # Cellule qui contient le numéro de doc
        self.doc.cell_num_doc = self.doc.onglet[self.doc.cell_numdoc_name]

        try:
            self.load_logo()
        except BaseException as e:
            logger.error(f"Erreur de chargement du logo : {e}.")
            self.view.show_feedback(txt=f"Erreur de chargement du logo.", message_type="error", stack=True)

        # ============== Cas de l'ouverture d'un doc existant
        if not self.doc.new_doc:
            try:
                self.doc.doc_name = self.get_value(self.doc.onglet, self.doc.cell_numdoc_name)
            except ValueError as e:
                self.view.show_feedback(
                    txt=f"Ce brouillon n'a pas été enregistré correctement, merci de le supprimer : {self.doc.doc_path}",
                    message_type="error")
                self.view.block_ui()

        # ============== Cas d'un nouveau document
        else:
            # On écrit le numéro de brouillon dans la cellule numéro de document
            self.doc.cell_num_doc.value = self.doc.doc_name

            # On copie les valeurs obligatoires de l'onglet config
            keys_seller = [
            "seller_name", "seller_tax_scheme_id", "seller_country", "seller_party_legal_entity_id",
            "seller_registration_name", "seller_mail", "seller_address"
            ]
            for key in keys_seller:
                try:
                    self.doc.onglet[key].value = self.get_value(self.doc.onglet_config, key)
                except pywintypes.com_error as e:
                    self.view.show_feedback(txt=f"Erreur avec la cellule {key}.", message_type="error")
                    logger.error(f"Erreur avec la cellule {key} : {e}.")

            # Valeurs facultatives
            self.doc.onglet['seller_cp_city'].value = str(int(self.doc.onglet_config['seller_postalzone'].value)) + ' ' + self.doc.onglet_config['seller_city'].value

            if self.doc.onglet_config['seller_address2'].value != None:
                self.doc.onglet['seller_address'].value += "\n" + self.doc.onglet_config['seller_address2'].value

            self.doc.onglet['seller_phone'].value = self.doc.onglet_config['seller_phone'].value
            self.doc.onglet['RCS'].value = self.doc.onglet_config['RCS'].value

            # Libellés pour une meilleure lisibilité
            self.doc.onglet['seller_party_legal_entity_id'].value = f"Siret : {self.doc.onglet['seller_party_legal_entity_id'].value}"
            self.doc.onglet['seller_tax_scheme_id'].value = f"TVA : {self.doc.onglet['seller_tax_scheme_id'].value}"
            self.doc.onglet['RCS'].value = f"RCS : {self.doc.onglet['RCS'].value}"

            # On enregistre immédiatement le document car s'il est fermé sans enregistrement il n'aura pas de numéro
            self.save_draft()

            # On recharge le menu Ouvrir
            self.view.set_menu_ouvrir(self.ouvrir, self.load_docs_paths())

        # Vérifications dans le doc
        self.doc.missing_cell_names = self.doc.check_cells()
        # En cas d'erreur dans le template
        if self.doc.missing_cell_names:
            erreur_cellules = ", ".join(self.doc.missing_cell_names)
            self.view.show_feedback(
                txt=f"Nom(s) de cellule manquants : {erreur_cellules}. "
                    f"Veuillez corriger le template ou utiliser le template original.",
                message_type="error")

        else:
            self.view.show_infos(
                doc=self.doc.type_doc, 
                etat=self.doc.state, 
                fichier=self.doc.doc_name, 
                date=self.doc.onglet[self.doc.cell_date_name].value.strftime('%Y-%m-%d'), 
                numero=''
            )
            self.view.enable_ui()

        # -------------- Transformation devis en facture ---------------- #
        if numClient:
            self.doc.onglet['num_client'].value = numClient

            lo = self.doc.onglet.api.ListObjects
            table = lo.Item("lignes_facture")

            if lignes:
                # on supprime les lignes vides du template
                while table.ListRows.Count > 0:
                    table.ListRows(1).Delete()
                # on alimente les valeurs du devis (ref qtite, remise - les autres valeurs sont récupérées de l'onglet produit)
                for l in lignes:
                    new_row = table.ListRows.Add()
                    ref_prod = new_row.Range.Columns(1)
                    ref_prod.Value = l[0]
                    qtite = new_row.Range.Columns(3)
                    qtite.Value = l[2]
                    remise = new_row.Range.Columns(7)
                    remise.Value = l[6]

    @staticmethod
    def get_type_doc(choice):
        if "DRAFT_FAC" in choice:
            return "facture"
        elif "DRAFT_DEV" in choice or "DEV_" in choice:
            return "devis"
        elif "DRAFT_AV" in choice:
            return "avoir"
        else:
            return None
            
    def load_docs_paths(self):
        docs_paths = []
        docs_names = []
        for root, dir, files in os.walk(DRAFT_PATH):
            for file in files:
                if file.startswith(("DRAFT_", "DEV_")):
                    docs_paths.append(os.path.join(root, file))
                    docs_names.append(file)
        return docs_paths

    def get_from_config(self, valeur):
        pass

    @staticmethod
    def named_cell_exists(key, onglet):
        """
        Une première vérification est faite à l'initialisation, mais l'utilisateur peut supprimer des cellules après.
        :param key: str : nom de la cellule
        :param onglet: obj : onglet
        :return: True ou False
        """
        # On reconstitue le nom de la cellule à partir de la clé
        key = onglet.name + "!" + key  # ex. : facture!date_echeance
        cell_exists = False
        try:    
            cell_exists = (key == onglet[key].name.name)
        except BaseException as e:
            logger.exception(e)
        return cell_exists


    @staticmethod
    def get_address(cell_name):
        """
        :param cell_name: onglet[key]
        :return: str adresse de la cellule
        """
        pos = cell_name.address
        return pos.replace("$", "")

    def get_value(self, onglet, key, data_type=None, required=True):
        """
        Récupère une valeur depuis une cellule excel.
        :param data_type (str): type de donnée (ex. date)
        :param required (bool): champ obligatoire, défaut True
        :return: valeur de la cellule (str)
        """
        # on vérifie que la cellule existe
        if not self.named_cell_exists(key, onglet):
            self.view.show_feedback(txt=f"Erreur dans l'onglet {onglet} : la clé {key} n'existe pas.",
                                   message_type="error", stack=True)
            self.view.block_ui()
            raise ValueError(f"La clé {key} n'existe pas..")
        else:
            try:
                # on vérifie que la cellule n'est pas vide
                val = onglet[key].value
                # une chaîne composée uniquement d'espaces est considérée vide
                if isinstance(val, str):
                    val = val.strip()
                if required is True:
                    if val is None or val == "":
                        self.show_error_excel(key, "Valeur obligatoire")
                        raise ValueError(f"La cellule '{self.get_address(onglet[key])}' ({key}) contient une valeur vide ou nulle.")

                if data_type == "date" and val not in (None, ""):
                    if not check_format_obj(datetime, val):
                        self.show_error_excel(key, "Format incorrect")
                        raise ValueError(f"La donnée de la cellule '{self.get_address(onglet[key])}' ({key}) n'est pas au bon format.")

                if data_type == "str":
                    return str(val)
                return val
            except BaseException as e:
                raise ValueError(e)

    # Noms de pays courants -> code ISO 3166-1 alpha-2 (CountryID Factur-X, BT-40/BT-55)
    _COUNTRY_MAP = {
        "france": "FR", "belgique": "BE", "belgium": "BE", "allemagne": "DE",
        "germany": "DE", "espagne": "ES", "spain": "ES", "italie": "IT",
        "italy": "IT", "luxembourg": "LU", "suisse": "CH", "switzerland": "CH",
        "pays-bas": "NL", "netherlands": "NL", "portugal": "PT",
        "royaume-uni": "GB", "united kingdom": "GB",
    }

    @classmethod
    def normalize_country(cls, val):
        """Retourne un code pays ISO 3166-1 alpha-2 (ex. 'France' -> 'FR')."""
        if val is None:
            return val
        s = str(val).strip()
        if len(s) == 2:
            return s.upper()
        return cls._COUNTRY_MAP.get(s.lower(), s)

    @classmethod
    def normalize_party_kwargs(cls, kw):
        """Normalise sur place les données d'une partie (vendeur/acheteur).

        - ``country`` -> code ISO 2 lettres ;
        - ``postalzone`` -> chaîne sans décimale (Excel renvoie souvent 44000.0).
        """
        if "country" in kw:
            kw["country"] = cls.normalize_country(kw["country"])

        pz = kw.get("postalzone")
        if isinstance(pz, float) and pz.is_integer():
            kw["postalzone"] = str(int(pz))
        elif pz is not None:
            kw["postalzone"] = str(pz).strip()
        return kw

    @staticmethod
    def increment_fact_number(invoice_number):
        parts = invoice_number.split('-')
        last_number = int(parts[-1])
        last_number += 1
        parts[-1] = f"{last_number:05d}"
        return '-'.join(parts)
    

    """
    Enregistre le brouillon excel dans le répertoire brouillons et créé le pdf
    :return: None
    """
    @check_ui
    def save_draft_and_pdf(self):
        self.view.show_feedback(txt="-> Création de la prévisualisation...")

        # Enregistrement du doc excel
        self.view.show_feedback(txt=f"-> Enregistrement du brouillon...", stack=True)

        self.doc.wb.save()

        self.view.show_infos(
                doc=self.doc.type_doc, 
                etat=self.doc.state, 
                fichier=self.doc.doc_name, 
                date=self.doc.onglet[self.doc.cell_date_name].value.strftime('%Y-%m-%d'), 
                numero=''
            )
        
        if self.gen_pdf():
            self.open_pdf(self.doc.pdf_path)
            self.show_success(devis=True)

    """
    Enregistre le brouillon excel dans le répertoire brouillons
    :return: None
    """
    @check_ui
    def save_draft(self):
        self.doc.wb.save()

    """
    Vérification de valeur nulles ou vides
    :return: True ou False
    """
    @staticmethod
    def check_empty_values(*args):
        for value in args:
            if value is None or str(value).strip() == '':
                return False
        return True
    
    """
    Vérifie si toutes les valeurs d'une ligne sont vides
    :return: True ou False
    """
    @staticmethod
    def all_empty(*args):
        for value in args:
            if value is not None and str(value).strip() != '':
                return False
        return True
    
    """
    Retourne la catégorie de TVA (str) au niveau des lignes
    Les types possibles en entrée sont : B (spécifique ezfacture), S, E, AE, K, G, O, L, M
    L'ajout d'un code B permet de différencier les biens des services
    """
    def get_tax_code(self, tva_code):
        if tva_code not in ("B", "S", "E", "AE", "K", "G", "O", "L", "M"):
            self.view.show_feedback(txt=f"Erreur : code de TVA erroné : {tva_code}", message_type="error", stack=True)
            raise ValueError(f"Code de TVA erroné : {tva_code}")
        else:
            return tva_code

    @staticmethod
    def to_standard_cat(letter):
        return "S" if letter == "B" else letter
    
        
    """
    Crée un nom de taxe définissant un type de taxe unique, en concaténant la catégorie et le pourcentage
    """
    def get_tax_name(self, cat, percent):
        if cat is None or percent is None:
            return None
        return cat + str(percent)
    
    """
    :param type_prestation: str, "B" (bien), "S" (service), "M" (mixte)
    :param acompte: bool, True si facture d'acompte, False si facture définitive
    :param paid: bool, True si la facture est déjà payée, sinon False
    """
    def calculate_profile(self, type_presta, acompte, paid):
        if type_presta is None:
            return None

        if acompte and not paid:
            number = "1"   # acompte non payé
        elif acompte and paid:
            number = "2"   # acompte déjà payé
        elif not acompte and paid:
            number = "2"   # facture définitive déjà payée
        elif not acompte and not paid:
            number = "4"   # facture définitive non payée
        else:
            return None    # cas impossible

        return f"{type_presta}{number}"
    
    @staticmethod
    def get_type_presta(uniques_tax_cat):
        """
        :param uniques_tax_cat: set, ensemble des catégories (uniques) de taxes de la facture
        """
        if "S" in uniques_tax_cat and "B" in uniques_tax_cat:
            return "M"
        elif "B" not in uniques_tax_cat:
            return "S"
        elif "S" not in uniques_tax_cat:
            return "B"
        else:
            return None


    """
    Valide les données et crée la facture au format pdf en incorporant les données factur-x.
    :param xml_file: bool, optionnel, pour écrire le xml dans un fichier séparé en plus du xml incorporé
    :param path_xml: string, optionnel, chemin du fichier xml
    """
    @check_ui
    def validate(self, xml_file=False, path_xml=None, ):

        confirm = self.view.open_alert_close("Etes-vous sûr ? \nLe document ne pourra plus être modifié ni supprimé.")
        if not confirm:
            return

        # ================== AVOIR ==================== #
        if self.doc.type_doc == "avoir":
            # BR-FR-CO-05 : un avoir doit référencer la facture d'origine (BT-25)
            # ET sa date (BT-26). Les deux cellules sont donc obligatoires.
            self.model.invoice_reference = self.get_value(self.doc.onglet, "reference_facture")
            self.model.invoice_reference_date = self.get_value(
                self.doc.onglet, "date_facture_reference", data_type="date"
            )

        # ================= FACTURE =================== #

        self.view.show_feedback(txt="Création de la facture...", stack=True)

        # Créer le document xml à partir des données excel
        self.model.currency = "EUR"

        keys_buyer = [
            "buyer_name", "buyer_tax_scheme_id", "buyer_country", "buyer_party_legal_entity_id",
            "buyer_registration_name", "buyer_mail", "buyer_address", "buyer_postalzone",
            "buyer_city"
        ]

        # Clés à extraire de l'onglet_config
        config_keys_seller = ["seller_name", "seller_tax_scheme_id", "seller_country", "seller_party_legal_entity_id",
                              "seller_registration_name", "seller_mail", "seller_address", "seller_postalzone", "seller_city"]

        try:
            # Construction des arguments pour Entity Seller - Suppression du préfixe 'seller_' dans les clés (noms des cellules excel)
            kwargs_seller = {
                key.replace("seller_", ""): self.get_value(self.doc.onglet_config, key) for key in config_keys_seller
            }

            # Construction des arguments pour Entity Buyer - Suppression du préfixe 'buyer_' dans les clés (noms des cellules excel)
            kwargs_buyer = {
                key.replace("buyer_", ""): self.get_value(self.doc.onglet, key) for key in keys_buyer
            }

            kwargs_seller["tax_scheme"] = kwargs_buyer["tax_scheme"] = "VAT"

            # Normalisation des données pour la conformité Factur-X (CountryID ISO,
            # code postal sans décimale, etc.)
            self.normalize_party_kwargs(kwargs_seller)
            self.normalize_party_kwargs(kwargs_buyer)

        except BaseException:
            msg = "Erreur lors de la création du document."
            self.view.show_feedback(txt=msg, message_type="error", stack=True)
            logger.exception(msg)
            raise ValueError(msg)

        else:
            try:
                # Création des instances (Entity) seller & buyer
                self.model.seller_party = Entity(**kwargs_seller)
                self.model.buyer_party = Entity(**kwargs_buyer)
            except BaseException:
                msg = "Erreur lors de la création des instances seller & buyer."
                self.view.show_feedback(txt=msg, message_type="error", stack=True)
                logger.exception(msg)
                raise ValueError(msg)

            try:
                self.model.issue_date = self.get_value(self.doc.onglet, "date_facture", data_type="date")
                if self.doc.type_doc == 'facture':
                    self.model.due_date = self.get_value(self.doc.onglet, "fact_date_echeance", data_type="date")
                    livraison = self.get_value(self.doc.onglet, "fact_date_livraison", data_type="date", required=False)
                else:
                    livraison = None
                # Date de livraison (BT-72) : cellule fact_date_livraison, à défaut la date de facture
                self.model.delivery_date = livraison or self.model.issue_date
            except BaseException as e:
                self.view.show_feedback(txt=str(e), message_type="error", stack=False)
                logger.exception(f"Erreur dates : {e}")
                raise ValueError(e)

        tbl = None
        if self.doc.type_doc == 'facture':
            tbl = self.doc.onglet.tables['lignes_facture']
        elif self.doc.type_doc == 'avoir':
            tbl = self.doc.onglet.tables['lignes_avoir']
        else:
            tbl = None

        data = tbl.range.options(ndim=2).value

        # headers = data[0]
        rows = data[1:]

        lines = []
        unique_tax_cat_list = set()

        # Régime de franchise en base de TVA (art. 293 B du CGI) : statut vendeur
        # global (lu dans regime_tva). Quand il est actif, toutes les lignes sont
        # exonérées (catégorie E) avec la mention légale et le code VATEX français,
        # quels que soient la catégorie et le taux saisis dans la ligne.
        regime_tva_val = str(self.doc.onglet_config['regime_tva'].value or "")
        franchise_293b = "293" in regime_tva_val
        fr_cat, fr_code, fr_reason = FR_FRANCHISE_EN_BASE

        # ============= Lignes de factures ================= #

        for row in rows:

            if self.all_empty(row[0], row[2], row[3], row[4]):
                msg = "La facture doit contenir au moins une ligne, et aucune ligne vide."
                self.view.show_feedback(txt=msg, message_type="error")
                raise ValueError(msg)

            # En franchise, la catégorie (row[5]) et le taux (row[4]) sont ignorés :
            # on n'exige donc que la désignation, la quantité et le prix.
            required_cells = (row[0], row[2], row[3]) if franchise_293b \
                else (row[0], row[2], row[3], row[4], row[5])
            if not self.check_empty_values(*required_cells):
                msg = "Erreur dans les lignes de la facture : certaines données sont manquantes."
                self.view.show_feedback(txt=msg, message_type="error")
                raise ValueError(msg)

            # Les catégories de tax ezfacture contiennent une catégorie supplémentaire "B" permettant de différencier les biens
            # dans la 6ème colonne (masquée) du tableau lignes_facture (row[5]).
            # On doit remettre cette valeur à S dans le xml.

            quantity=row[2]
            unit_code="EA"
            price=row[3]
            item_name=row[0]
            currency="EUR"
            discount=row[6]

            if franchise_293b:
                # Ligne exonérée : catégorie E, taux 0, mention + code VATEX 293 B.
                ez_cat = fr_cat
                tax_percent = 0
                exemption_reason = fr_reason
                exemption_reason_code = fr_code
            else:
                ez_cat = self.get_tax_code(row[5])
                tax_percent = row[4]
                exemption_reason = None
                exemption_reason_code = None

            tax_name = self.get_tax_name(ez_cat, tax_percent)  # On définit le nom de la taxe arbitrairement en concaténant la catégorie ezfacture et le pourcentage
            unique_tax_cat_list.add(ez_cat)  # pour définir le profil B, S ou M

            # Si la config est tva sur les débits et qu'on est sur une ligne de services, on alimente DueDateTypeCode dans le xml
            if ez_cat == "S" and regime_tva_val.endswith("5"):
                tax_debits = True
            else:
                tax_debits = False

            try:
                line = InvoiceLine(
                    quantity=quantity,
                    unit_code=unit_code,
                    price=price,
                    item_name=item_name,
                    currency=currency,
                    tax_category=self.to_standard_cat(ez_cat),  # La catégorie B, spécifique Ezfacture n'est pas dans la norme -> on remet la valeur initiale "S"
                    tax_percent=tax_percent,
                    tax_name=tax_name,
                    discount=discount,
                    tax_debits=tax_debits,
                    exemption_reason=exemption_reason,
                    exemption_reason_code=exemption_reason_code
                )
                lines.append(line)

            except BaseException:
                msg = "Erreur lors de l'ajout d'une ligne."
                self.view.show_feedback(txt=msg, message_type="error", stack=True)
                logger.exception(msg)
                raise ValueError(msg)
            
        self.model.add_lines_from(lines)

        # try:
        #     self.model.total_ht = self.get_value(self.doc.onglet, "total_ht")
        # except BaseException:
        #     msg = "Erreur de total HT."
        #     self.view.show_feedback(txt=msg, message_type="error")
        #     logger.exception(msg)
        #     raise ValueError(msg)
        
        try:
            # le champ acompte existe dans le template avoir mais n'est pas utilisé (doit être = 0)
            self.model.paid_amount = self.get_value(self.doc.onglet, "acompte", required=False) or 0
        except BaseException as e:
            self.view.show_feedback(txt=f"Erreur : {e}", message_type="error")
            logger.exception(e)
            raise ValueError(e)
        
        try:
            self.model.payable_amount = self.get_value(self.doc.onglet, "net_a_payer")
        except BaseException as e:
            self.view.show_feedback(txt=f"Erreur : {e}", message_type="error")
            logger.exception(e)
            raise ValueError(e)
    
        self.view.show_feedback(txt="-> Création des données xml...", stack=True)
        
        # Calcul du profil BT-23 (cadre de facturation)
        type_presta = self.get_type_presta(unique_tax_cat_list)
        acompte = True if self.doc.onglet['acompte'].value > 0 else False
        paid = True if self.doc.onglet['net_a_payer'].value == 0 else False
        self.model.profile_id = self.calculate_profile(type_presta, acompte, paid)

        # Mentions légales françaises obligatoires (BR-FR-05 / BT-22),
        # lues dans les cellules nommées PMD / PMT / AAB de l'onglet config.
        try:
            self.model.notes = [
                {"code": code, "content": self.get_value(self.doc.onglet_config, code)}
                for code in ("PMD", "PMT", "AAB")
            ]
        except BaseException as e:
            self.view.show_feedback(txt=str(e), message_type="error", stack=True)
            logger.exception(f"Erreur mentions légales : {e}")
            raise ValueError(e)

        # Note libre de l'utilisateur (BT-22), lue dans la cellule note_facture / note_avoir
        # de l'onglet document. Champ facultatif, sans code sujet (BT-21).
        try:
            note_key = "note_facture" if self.doc.type_doc == "facture" else "note_avoir"
            note_libre = self.get_value(self.doc.onglet, note_key, required=False)
            # on ignore le texte indicatif laissé par défaut dans le template
            if note_libre and note_libre != "Note de facture":
                self.model.notes.append({"code": "", "content": note_libre})
        except BaseException as e:
            self.view.show_feedback(txt=str(e), message_type="error", stack=True)
            logger.exception(f"Erreur note document : {e}")
            raise ValueError(e)

        # Validation et création du document
        try:
            with Eztransaction() as tx:

                # 1. Récupération du numéro (lecture seulement)
                self.model.invoice_id = self.doc.cell_num_doc.value = tx.do(step_fn=lambda: self.backend.get_number(type_doc=self.doc.type_doc))

                # 2. Création du xml Factur-X
                tx.do(
                    step_fn=lambda: self.model.save(...),
                    rollback_fn=lambda: self.model.unsave()
                )

                # 3. Validation XML
                tx.do(step_fn=lambda: xml_check_xsd(self.model.xml))

                # 4. Génération PDF
                self.view.show_feedback(txt="-> Création du fichier pdf...", stack=True)
                tx.do(
                    step_fn=lambda: self.gen_pdf(),
                    rollback_fn=lambda: self.doc.unmake_pdf()
                )

                # 5. Incorporation du XML dans le PDF
                self.view.show_feedback(txt="-> Incorporation du xml...", stack=True)
                tx.do(step_fn=lambda: generate_from_file(pdf_file=self.doc.pdf_path, xml=self.model.xml, flavor='factur-x', check_schematron=True))

                # 6. Réservation du numéro (LOCAL : PREPARE / API : création serveur)
                prepared_entry_or_void = tx.do(
                    step_fn=lambda: self.backend.reserve(
                        self.model.invoice_id,
                        self.doc.type_doc
                    ),
                    rollback_fn=lambda: self.backend.cancel(self.model.invoice_id)
                )

                # 7. COMMIT FINAL (LOCAL : écriture JSONL / API : no-op)
                tx.do(
                    step_fn=lambda: None,
                    commit_fn=lambda: self.backend.commit(prepared_entry_or_void)
                )

            self.doc.state = 'Validé'
            self.show_success()
            self.doc.close()
            self.doc.delete_draft()

        except BaseException as e:
            msg = f"Erreur lors de la création du document : {e}"
            self.view.show_feedback(txt=msg, message_type="error", stack=True)
            self.view.delete_messages("infos")
            logger.exception(msg)
            raise ValueError(msg)

        
        # Ouverture du document pdf
        self.open_pdf(self.doc.pdf_path)
        self.doc = None  # on réinitialise


    """
    Transforme un devis en facture
    """
    @check_ui
    def transform(self):
        num_client = self.get_value(self.doc.onglet, "num_client")
        tbl = self.doc.onglet.tables['lignes_devis']
        data = tbl.range.options(ndim=2).value
        rows = data[1:]
        # passer numclient et rows
        self.create_doc("Facture", numClient=num_client, lignes=rows)


    def show_success(self, devis=False):
        self.view.show_feedback(txt="Document créé !", message_type="success")
        if not devis:
            self.view.block_boutons(["save", "valider"])
            num=self.model.invoice_id
        else:
            self.doc.state = "Enregistré"
            num = ""
        self.view.show_infos(
            doc=self.doc.type_doc, 
            etat=self.doc.state, 
            fichier=self.doc.pdf_facture_num, 
            date=self.doc.onglet[self.doc.cell_date_name].value.strftime('%Y-%m-%d'), 
            numero=num
        )
        self.view.reset_menu_nouveau("Nouveau")

    def aide(self):
        webbrowser.open("https://www.ezfacture.fr/documentation")

    def restart(self):
        self.doc.close()
        self.doc = None
        self.view.close()
        gc.collect()
        python = sys.executable
        os.execl(python, python, *sys.argv)  # Redémarre le script

    def connexion(self):
        self.view.buttons["connexion"].configure(state="disabled", text="Démarrage en cours...")
        Thread(target=self._connexion_worker, daemon=True).start()

    def _connexion_worker(self):
        def _reset_btn():
            self.view.buttons["connexion"].configure(
                state="normal", text="Démarrer",
                fg_color="#3498db", hover_color="#2980b9"
            )

        try:
            # login() : authentification API, ou simple déverrouillage en local.
            if self.backend.login():
                def _on_success():
                    self.view.menu_nouveau.configure(state="normal")
                    self.view.menu_ouvrir.configure(state="normal")
                    self.view.enable_boutons(["aide"])
                    self.view.buttons["connexion"].configure(text="Ezfacture prêt...", fg_color="#242424")
                    self.view.buttons["connexion"].configure(state="disabled")
                    self.view.delete_messages("feedback")
                    if not self.backend.integrity_ok():
                        self.view.show_feedback(txt="Le fichier des factures est corrompu suite à une modification. \nIl est nécessaire de le restaurer à partir d'une sauvegarde.", message_type="error")
                self.view.after(0, _on_success)
            else:
                msg = "Problème lors du démarrage."
                logger.error(msg)
                def _on_failure():
                    self.view.block_ui()
                    _reset_btn()
                    self.view.show_feedback(txt=msg, message_type="error")
                self.view.after(0, _on_failure)

        except BaseException:
            msg2 = "Problème lors du démarrage de l'application."
            logger.exception(msg2)
            def _on_error():
                self.view.block_ui()
                _reset_btn()
                self.view.show_feedback(txt=msg2, message_type="error")
            self.view.after(0, _on_error)

    """
    Génère le fichier pdf
    """
    def gen_pdf(self):

        # Avant de faire le pdf on essaie de basculer sur l'imprimante pdf. Pour ça on a appelé au préalable disable_win_auto_printer_manage().
        # Décommenter la ligne ci-dessous et set_default_printer(default_printer) plus loin pour rebasculer sur l'imprimante par défaut à la fin 
        # default_printer = get_default_printer_name()
        pdf_printer = get_microsoft_print_to_pdf_printer()
        if set_default_printer(pdf_printer):
            logger.info(f"Imprimante par défaut définie sur {pdf_printer}")

        try:
            # couleurs fond et polices pour l'impression
            self.doc.onglet.range(RANGE_PRINT_AREA).font.color = "#000000"

            if self.doc.type_doc in ('facture', 'avoir'):
                self.doc.onglet.range('adresse_client').color = "#ffffff"
                self.doc.onglet.range('date_facture').color = "#ffffff"
                self.doc.onglet.range('acompte').color = "#ffffff"
                self.doc.onglet.range('num_client').color = "#ffffff"

            if self.doc.type_doc == 'facture':
                self.doc.onglet.tables['lignes_facture'].data_body_range.color = "#ffffff"
                self.doc.onglet.range('fact_ref_commande').color = "#ffffff"
                self.doc.onglet.range('fact_date_livraison').color = "#ffffff"
                self.doc.onglet.range('fact_date_echeance').color = "#ffffff"
                if self.doc.onglet['note_facture'].value == "Note de facture":
                    self.doc.onglet['note_facture'].value = ""

            if self.doc.type_doc == 'avoir':
                self.doc.onglet.tables['lignes_avoir'].data_body_range.color = "#ffffff"
                self.doc.onglet.range('reference_facture').color = "#ffffff"
                self.doc.onglet.range('date_facture_reference').color = "#ffffff"
                if self.doc.onglet['note_avoir'].value == "Note de facture":
                    self.doc.onglet['note_avoir'].value = ""

            if self.doc.type_doc == 'devis':
                self.doc.onglet.range('num_client').color = "#ffffff"
                self.doc.onglet.range('dev_date_devis').color = "#ffffff"
                self.doc.onglet.range('dev_date_valid').color = "#ffffff"
                self.doc.onglet.range('description_devis').color = "#ffffff"
                self.doc.onglet.range('adresse_client').color = "#ffffff"
                self.doc.onglet.tables['lignes_devis'].data_body_range.color = "#ffffff"

        except pywintypes.com_error as e:
            logger.error(f"Erreur lors de la préparation du document pdf : {e}.")

        try:
            self.doc.make_pdf('./pdf')

            # # On essaie de rebasculer sur l'imprimante par défaut.
            # if set_default_printer(default_printer):
            #     logger.info(f"Imprimante par défaut redéfinie sur {default_printer}")
            # else:
            #     logger.warning(f"Impossible de rebasculer sur l'imprimante par défaut {default_printer}")

        except pywintypes.com_error as e:
            logger.error(f"Erreur COM lors de la création du PDF : {e}")
            self.view.show_feedback(
                txt="Erreur lors de la création du PDF",
                message_type="error", stack=True
            )
            return False

        except BaseException as e:
            logger.exception(e)
            self.view.show_feedback(
                txt="Erreur lors de la création du PDF. Si le document est ouvert, merci de le fermer avant d'enregistrer.",
                message_type="error", stack=True
            )
            return False

        return True

    def show_error_excel(self, cellule, message):
        """
        Affiche un message d'erreur sur l'interface et dans excel
        :param cellule: nom ou coordonnées de la cellule
        :param message: texte à afficher
        :return: None
        """
        self.doc.cell_error(self.doc.onglet, cellule, message)
        self.view.show_feedback(txt="Données incorrectes, merci de vérifier le document.", message_type="error", stack=True)


    def load_logo(self):
        image_path = resource_path("images/logo.png")
        cell = self.doc.onglet.range('logo')
        pic = self.doc.onglet.pictures.add(
            str(image_path),
            left=cell.left,
            top=cell.top
        )

        pic.height = cell.height


    def open_pdf(self, path):
        try:
            os.startfile(path)
        except OSError as e:
            logger.error(f"Impossible d'ouvrir le PDF : {e}.")
            self.view.show_feedback(txt="Impossible d'ouvrir le fichier PDF", message_type="error", stack=True)


    def load_data(self):
        try:
            sheet_clients = self.doc.wb.sheets["clients"]
            df_clients = pd.read_excel(r'./clients.xlsx', sheet_name='clients')
            table_clients = sheet_clients.tables['table_clients']
            table_clients.data_body_range.value = df_clients.values.tolist()

            sheet_produits = self.doc.wb.sheets["produits"]
            df_produits = pd.read_excel(r'./produits.xlsx', sheet_name='produits')
            table_produits = sheet_produits.tables['table_produits']
            table_produits.data_body_range.value = df_produits.values.tolist()

            sheet_config = self.doc.wb.sheets["config"]
            df_config = pd.read_excel(r'./config.xlsx', sheet_name='config')
            sheet_config['A1'].options(index=False, header=True).value = df_config

        except BaseException as e:
            msg= "Erreur lors du chargement des données, merci de vérifier la présence des fichiers clients.xlsx, produits.xlsx et config.xlsx dans le répertoire de l'application."
            logger.error(f"{msg} : {e}.")
            self.view.show_feedback(txt=msg, message_type="error", stack=True)
            raise ValueError(msg)


