import os
import logging

from controller.backends.local import LocalBackend

logger = logging.getLogger(__name__)


def get_numbering_backend(controller):
    """Sélectionne le backend de numérotation à utiliser.

    Priorité :
      1. Override explicite :
         - variable d'environnement ``EZFACTURE_MODE`` = ``local`` / ``api`` ;
         - ou constante ``MODE_API`` de config.py (True force l'API, False force le
           local, None = auto).
      2. Sinon, détection automatique : backend API si le sous-module privé
         ``controller/backends/api`` est présent, backend local sinon.

    Ainsi, un clone public (sans accès au sous-module) bascule automatiquement en
    mode local.
    """
    try:
        from config import MODE_API
    except ImportError:
        MODE_API = None

    mode = (os.getenv("EZFACTURE_MODE") or "").strip().lower() or None

    force_local = mode == "local" or MODE_API is False
    force_api = mode == "api" or MODE_API is True

    if force_local:
        logger.info("Backend de numérotation : local (forcé).")
        return LocalBackend(controller)

    try:
        from controller.backends.api import ApiBackend
    except ImportError:
        if force_api:
            raise RuntimeError(
                "Mode API demandé mais le module privé "
                "'controller/backends/api' est absent."
            )
        logger.info("Module API absent → backend de numérotation : local.")
        return LocalBackend(controller)

    logger.info("Backend de numérotation : API.")
    return ApiBackend(controller)


__all__ = ["get_numbering_backend", "LocalBackend"]
