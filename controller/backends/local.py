import json
import logging
import shutil
from datetime import datetime
from pathlib import Path

from config import LOCAL_FILE
from tools.utils import (
    load_last_entry_by_type,
    compute_file_hash,
    sha256,
    write_registry_seal,
    verify_local_file,
)
from controller.backends.base import NumberingBackend

logger = logging.getLogger(__name__)


class LocalBackend(NumberingBackend):
    """Numérotation locale.

    Les numéros sont stockés en append-only dans ``invoices.jsonl``, protégés par
    une chaîne de hash (``file_hash_before`` / ``self_hash``) et un sceau dans le
    registre Windows. Aucun accès réseau.
    """

    requires_auth = False

    def login(self):
        # Pas d'authentification en local : on considère l'accès toujours ouvert.
        # L'intégrité du fichier est vérifiée séparément via ``integrity_ok``.
        return True

    def integrity_ok(self):
        return verify_local_file(LOCAL_FILE)

    def get_number(self, type_doc="facture"):
        prefix_map = {'facture': 'FAC', 'avoir': 'AV'}
        prefix = prefix_map.get(type_doc, 'FAC')
        today = datetime.now()
        year_month_prefix = f"{prefix}-{today.year}-{today.month:02d}"

        last_number = load_last_entry_by_type(type_doc)

        if last_number is None:
            cell_name = "first_avoir_number" if type_doc == "avoir" else "first_invoice_number"
            first_num = int(self.controller.doc.onglet_config[cell_name].value)
            new_num = f"{year_month_prefix}-{first_num:05d}"
        else:
            last_seq = int(last_number.split('-')[-1])
            new_num = f"{year_month_prefix}-{last_seq + 1:05d}"

        return new_num

    def reserve(self, number, type_doc="facture"):
        """PREPARE — ne rien écrire, retourne l'enveloppe à committer."""
        file_hash_before = compute_file_hash(LOCAL_FILE)

        entry = {
            "number": number,
            "type": type_doc,
            "timestamp": datetime.utcnow().isoformat(),
            "file_hash_before": file_hash_before,
        }

        # SELF HASH = hash(number + type + timestamp + file_hash_before)
        raw = (
            entry["number"]
            + entry["type"]
            + entry["timestamp"]
            + entry["file_hash_before"]
        )
        entry["self_hash"] = sha256(raw)

        return entry

    def commit(self, entry):
        """COMMIT — écriture append-only dans le JSONL."""
        raw = (
            entry["number"]
            + entry["type"]
            + entry["timestamp"]
            + entry["file_hash_before"]
        )
        if sha256(raw) != entry["self_hash"]:
            raise ValueError("Hash incohérent : entrée altérée avant commit")

        with open(LOCAL_FILE, "a", encoding="utf8") as f:
            f.write(json.dumps(entry) + "\n")

        with open(LOCAL_FILE, "r", encoding="utf8") as f:
            count = sum(1 for line in f if line.strip())
        write_registry_seal(count, entry["self_hash"])

        logger.info(f"Facture {entry['number']} validée et écrite dans {LOCAL_FILE}")

        # Sauvegarde de sécurité (best-effort : n'invalide jamais le commit).
        self._backup()

    def _backup(self):
        """Copie ``invoices.jsonl`` vers le répertoire indiqué dans la cellule
        ``config!backup``.

        Best-effort : toute erreur est journalisée et signalée à l'utilisateur
        mais n'interrompt pas la validation, la facture étant déjà écrite.
        """
        try:
            onglet_config = self.controller.doc.onglet_config

            if self.controller.named_cell_exists("backup", onglet_config):
                dest_dir = onglet_config["backup"].value
            else:
                dest_dir = None

            if dest_dir is None or str(dest_dir).strip() == "":
                # Cellule 'backup' absente ou vide
                dest_dir = "."

            dest_dir = Path(str(dest_dir).strip())
            dest_dir.mkdir(parents=True, exist_ok=True)
            dest_file = dest_dir / (Path(LOCAL_FILE).name + ".bak")
            shutil.copy2(LOCAL_FILE, dest_file)
            logger.info(f"Sauvegarde de {LOCAL_FILE} vers {dest_file}")

        except Exception as e:
            msg = f"Échec de la sauvegarde de {LOCAL_FILE} : {e}"
            logger.warning(msg)
            try:
                self.controller.view.show_feedback(
                    txt=f"Attention : {msg}", message_type="error", stack=True
                )
            except Exception:
                pass

    def cancel(self, number):
        # En local, rien à annuler : le numéro n'est jamais écrit avant validation.
        return True
