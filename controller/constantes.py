from pathlib import Path
import sys

NAMES_FACT = ['facture!fact_date_echeance', 'facture!date_facture', 'facture!fact_date_livraison',
              'facture!num_client', 'facture!fact_num_fact', 'facture!fact_ref_commande']
NAMES_DEVIS = ['devis!dev_num_devis', 'devis!dev_date_devis', 'devis!dev_date_valid', 'devis!num_client']
NAMES_AVOIR = ['avoir!av_num_avoir', 'avoir!date_facture']

NOM_CELL_NUM_DOC = {
    'facture': 'fact_num_fact',
    'devis': 'dev_num_devis',
    'avoir': 'av_num_avoir'
}

if getattr(sys, 'frozen', False):
    _base = Path(sys.executable).parent
else:
    _base = Path(__file__).parent.parent

DRAFT_PATH = (_base / "brouillons").resolve()

RANGE_PRINT_AREA = '$A$1:$I$54'
