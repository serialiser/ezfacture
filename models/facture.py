from en16931 import Invoice, Entity, InvoiceLine
import os
import tempfile


class Facture(Invoice):
    def __init__(self, invoice_id=None, invoice_type_code=None, currency="EUR", profile_id=None):
        super().__init__()
        self.invoice_id = invoice_id
        self.invoice_type_code = invoice_type_code
        self.invoice_reference = None  # Pour les avoirs
        self.currency = currency
        self.profile_id = profile_id
        self._xml = ""

    @property
    def xml(self):
        return self._xml

    def save(self, path=None, to_file=False):
        """
        Enregistre les données de la facture en xml.
        Réécriture de cette méthode pour enregistrer de manière sécurisée
        et plus performante qd le xml existe déjà
        """
        if path is None:
            path = 'invoice_{}.xml'.format(self.invoice_id)

        self._xml = self._original_xml if self._original_xml is not None else self.to_xml()

        if to_file:
            os.makedirs(path, exist_ok=True)

            output_path = os.path.join(path, f"{self.invoice_id}.xml")

            # Création d'un fichier temporaire puis remplacement de l'ancien fichier
            temp_fd, temp_path = tempfile.mkstemp(dir=path)
            try:
                with os.fdopen(temp_fd, 'w', encoding='utf-8') as f:
                    f.write(self.xml)
                os.replace(temp_path, output_path)
            except BaseException as e:
                os.remove(temp_path)
                raise e

    def unsave(self):
        self._xml=""