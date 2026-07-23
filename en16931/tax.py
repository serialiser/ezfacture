"""
Class for representing a Tax category.
"""
from en16931.utils import parse_float


CATEGORIES = {'AE', 'L', 'M', 'E', 'S', 'Z', 'G', 'O', 'K'}

# Catégories de TVA exonérées / non soumises qui EXIGENT un motif d'exonération
# (BT-120 = texte, BT-121 = code VATEX) dans le récapitulatif de TVA (BG-23),
# sous peine d'échec de la validation EN16931 (règles BR-AE-10, BR-E-10,
# BR-IC-10, BR-G-10, BR-O-10). Une seule des deux valeurs suffit ; on
# fournit un libellé par défaut, et le code VATEX quand il est univoque.
# Format : catégorie -> (code_vatex_ou_None, libellé_par_défaut).
DEFAULT_EXEMPTION_REASONS = {
    'AE': ('VATEX-EU-AE', 'Autoliquidation'),
    'K':  ('VATEX-EU-IC', 'Exonération de TVA, article 262 ter I du CGI'),
    'G':  ('VATEX-EU-G',  'Exportation hors de l’Union européenne'),
    'O':  ('VATEX-EU-O',  'Non soumis à la TVA'),
    'E':  (None,          'Exonéré de TVA'),
}

# Motif d'exonération spécifique à la FRANCHISE EN BASE de TVA (art. 293 B du
# CGI). Ce cas relève de la catégorie 'E' (exonéré) mais NE peut PAS être le
# défaut de 'E' (qui couvre aussi d'autres exonérations) : il doit être posé
# explicitement. Code BT-121 propre à la France + mention légale (BT-120).
# Source : spécifications DGFiP / code list Factur-X (VATEX-FR-FRANCHISE).
FR_FRANCHISE_EN_BASE = ('E', 'VATEX-FR-FRANCHISE',
                        'TVA non applicable, article 293 B du CGI')

# Catégories pour lesquelles le taux de TVA (BT-119 au récap, BT-152 en ligne)
# NE DOIT PAS être présent (règles EN16931 BR-O-5 / BR-O-8, « Non soumis à la
# TVA »). Les autres catégories exonérées (E, AE, K, G, Z) portent un taux 0.
NO_RATE_CATEGORIES = {'O'}


class Tax:
    """Tax class.

    It representas a tax to apply globally or to a concrete
    invoice line.

    Only categories of taxes enabled by the EN16931 standard are
    supported. See the documentation of :meth:`category` property
    for more details.

    You can create Tax objects directly:

    >>> t = Tax(0.21, "S", "IVA")

    Or specify the relevant attributes when building
    :class:`InvoiceLines` or :class:`Invoice`

    """

    def __init__(self, percent, category, name, comment="",
                 exemption_reason=None, exemption_reason_code=None):
        """Initialize a Tax object.

        Parameters
        ----------
        category: string.
            A string representing the category of the Tax.
            It must be one of 'AE', 'L', 'M', 'E', 'S', 'Z',
            'G', 'O', or 'K'.

        percent: float.
            The percentage of the Tax. Can be 0.

        name: string.
            Arbitrary name to identify the Tax.

        comment: string.
            A comment on the tax.

        exemption_reason: string (optional).
            Free-text VAT exemption reason (BT-120). For exempt / not
            subject categories (E, AE, K, G, O), it defaults to the entry
            in :data:`DEFAULT_EXEMPTION_REASONS` when not provided.

        exemption_reason_code: string (optional).
            VATEX exemption reason code (BT-121). Same default behaviour
            as ``exemption_reason``.

        Notes
        -----
        A tax is compared to other Tax objects by equality of their
        percentage, category, and name.

        """
        self.category = category
        self.name = name
        self.comment = comment
        pct = parse_float(percent)
        if pct > 1 or pct < -1:
            self.percent = pct / 100
        else:
            self.percent = pct
        default_code, default_reason = DEFAULT_EXEMPTION_REASONS.get(
            category, (None, None))
        self.exemption_reason = (exemption_reason if exemption_reason is not None
                                 else default_reason)
        self.exemption_reason_code = (exemption_reason_code
                                      if exemption_reason_code is not None
                                      else default_code)

    @property
    def category(self):
        """Property: The category of the Tax.

        Parameters
        ----------
        category: string.
            A string representing the category of the Tax.
            It must be one of 'AE', 'L', 'M', 'E', 'S', 'Z',
            'G', 'O', or 'K'.

        Raises
        ------
        ValueError: if the category is not valid.

        """
        return self._category

    @category.setter
    def category(self, category):
        """Sets the category of the Tax.
        """
        if category not in CATEGORIES:
            msg = "Category {} not valid. Use one of {}"
            raise ValueError(msg.format(category, CATEGORIES))
        self._category = category

    @property
    def has_rate(self):
        """False si la catégorie interdit la présence d'un taux de TVA.

        Sert à ne pas émettre ``RateApplicablePercent`` pour la catégorie
        'O' (non soumis à la TVA), conformément aux règles BR-O-5 / BR-O-8.
        """
        return self.category not in NO_RATE_CATEGORIES

    @property
    def code(self):
        """An identification code of the tax.
        """
        return "{}_{}".format(self.percent, self.category)

    def __eq__(self, other):
        """
        A tax is compared to other Tax objects by equality of their
        percentage, category, and name.
        """
        if other is None:
            return False
        return (self.percent == other.percent and
                self.category == self.category and
                self.name == self.name)

    def __hash__(self):
        return hash(self.code)

    def __repr__(self):
        return "Tax {}: {} {} {}".format(self.category, self.percent,
                                         self.name, self.comment)
