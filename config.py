"""
Configuration publique de l'application.

Le mode de numérotation est choisi automatiquement par
``controller.backends.get_numbering_backend`` :
- backend API si le sous-module privé ``controller/backends/api`` est présent ;
- backend local sinon.

On peut forcer un mode via :
- la variable d'environnement ``EZFACTURE_MODE`` = ``local`` / ``api`` ;
- ou la constante ``MODE_API`` ci-dessous.
"""
import logging

# None = automatique (selon la présence du backend API)
# True = force le mode API ; False = force le mode local
MODE_API = True

LOCAL_FILE = "invoices.jsonl"


# Levels : DEBUG, INFO, WARNING, ERROR, CRITICAL
# LEVEL = logging.WARNING if PROD else logging.DEBUG
LEVEL = logging.DEBUG
HANDLERS = [logging.FileHandler("log.log", encoding='utf-8')]

if LEVEL == logging.DEBUG:
    HANDLERS.append(logging.StreamHandler())


def setup_logger():
    logging.basicConfig(
        level=LEVEL,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=HANDLERS,
        force=True  # Force la réinitialisation de la conf pour éviter les conflits si logging.basicConfig() a déjà été configuré
    )
