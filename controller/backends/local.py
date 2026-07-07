import json
import logging
from datetime import datetime

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

    def cancel(self, number):
        # En local, rien à annuler : le numéro n'est jamais écrit avant validation.
        return True
