from abc import ABC, abstractmethod


class NumberingBackend(ABC):
    """Interface d'un backend de numérotation / authentification.

    Le backend reçoit le Controller afin d'accéder au document Excel
    (``controller.doc.onglet_config``), à la vue et au modèle en cours.

    Deux implémentations existent :
    - ``LocalBackend`` (public) : numéros stockés localement dans invoices.jsonl ;
    - ``ApiBackend`` (module privé) : numéros gérés par l'API ezfacture (OAuth2).
    """

    #: True si le backend nécessite une authentification réseau.
    requires_auth = False

    def __init__(self, controller):
        self.controller = controller

    @abstractmethod
    def login(self):
        """Authentifie / initialise l'accès.

        :return: un token (API) ou True (local) en cas de succès, sinon une valeur falsy.
        """
        raise NotImplementedError

    def integrity_ok(self):
        """Vérification d'intégrité post-login (fichier local). True par défaut."""
        return True

    @abstractmethod
    def get_number(self, type_doc="facture"):
        """Retourne le prochain numéro de document (lecture seule, aucune écriture)."""
        raise NotImplementedError

    @abstractmethod
    def reserve(self, number, type_doc="facture"):
        """Réserve le numéro.

        :return: une ``entry`` à committer (local) ou None (API).
        """
        raise NotImplementedError

    @abstractmethod
    def commit(self, entry):
        """Valide définitivement la réservation (écriture locale ; no-op côté API)."""
        raise NotImplementedError

    @abstractmethod
    def cancel(self, number):
        """Annule la réservation (utilisé lors d'un rollback de transaction)."""
        raise NotImplementedError
