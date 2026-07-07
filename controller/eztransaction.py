import logging

logger = logging.getLogger(__name__)


class EztransactionError(Exception):
    pass


class Eztransaction:
    """
    Transaction avec actions :
    - step_fn : exécutée immédiatement
    - rollback_fn : exécutée si exception dans la transaction
    - commit_fn : exécutée seulement si aucun rollback et transaction OK
    """

    def __init__(self):
        self.rollback_actions = []
        self.commit_actions = []

    def do(self, step_fn, rollback_fn=None, commit_fn=None):
        result = step_fn()

        # rollback exécuté en cas d'échec
        if rollback_fn:
            self.rollback_actions.append(rollback_fn)

        # commit exécuté seulement si toute la transaction réussit
        if commit_fn:
            self.commit_actions.append(commit_fn)

        return result

    def rollback(self):
        for action in reversed(self.rollback_actions):
            try:
                action()
            except Exception as e:
                logger.exception(f"Erreur durant un rollback: {e}")

    def commit(self):
        for action in self.commit_actions:
            try:
                action()
            except Exception as e:
                logger.exception(f"Erreur durant commit: {e}")
                raise

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, tb):
        if exc_type:
            logger.error("Transaction échouée → rollback")
            self.rollback()
            return False  # laisse remonter l'erreur
        else:
            try:
                self.commit()
            except Exception:
                logger.error("Commit échoué → rollback")
                self.rollback()
                raise
            return False
