import pytest
from collections.abc import Hashable

from en16931.tax import Tax, FR_FRANCHISE_EN_BASE
from en16931.invoice_line import InvoiceLine


class TestTaxes:

    def test_initialization(self):
        t = Tax(0.21, "S", "IVA")
        assert t

    def test_hashable(self):
        t = Tax(0.21, "S", "IVA")
        assert isinstance(t, Hashable)

    def test_percent_less_than_one(self):
        t = Tax(0.21, "S", "IVA")
        assert t.percent == 0.21

    def test_percent_more_than_one(self):
        t = Tax(21, "S", "IVA")
        assert t.percent == 0.21

    def test_percent_string(self):
        t = Tax("21", "S", "IVA")
        assert t.percent == 0.21

    def test_cmp_with_None(self):
        t = Tax("21", "S", "IVA")
        assert not (t == None)

    def test_value_error_bad_percent(self):
        with pytest.raises(ValueError):
            t = Tax("asdf", "S", "IVA")

    def test_value_error_bad_category(self):
        with pytest.raises(ValueError):
            t = Tax("21", "asd", "IVA")

    def test_standard_category_has_no_exemption_reason(self):
        t = Tax(0.21, "S", "IVA")
        assert t.exemption_reason is None
        assert t.exemption_reason_code is None

    def test_intracommunity_defaults(self):
        t = Tax(0, "K", "K0")
        assert t.exemption_reason == "Exonération de TVA, article 262 ter I du CGI"
        assert t.exemption_reason_code == "VATEX-EU-IC"

    def test_reverse_charge_defaults(self):
        t = Tax(0, "AE", "AE0")
        assert t.exemption_reason == "Autoliquidation"
        assert t.exemption_reason_code == "VATEX-EU-AE"

    def test_exempt_has_text_but_no_code(self):
        t = Tax(0, "E", "E0")
        assert t.exemption_reason == "Exonéré de TVA"
        assert t.exemption_reason_code is None

    def test_standard_category_has_rate(self):
        assert Tax(0.21, "S", "IVA").has_rate is True

    def test_not_subject_category_has_no_rate(self):
        t = Tax(0, "O", "O0")
        assert t.has_rate is False
        assert t.exemption_reason == "Non soumis à la TVA"
        assert t.exemption_reason_code == "VATEX-EU-O"

    def test_exemption_reason_can_be_overridden(self):
        t = Tax(0, "K", "K0", exemption_reason="Motif sur mesure",
                exemption_reason_code="VATEX-EU-79-C")
        assert t.exemption_reason == "Motif sur mesure"
        assert t.exemption_reason_code == "VATEX-EU-79-C"

    def test_franchise_en_base_constant(self):
        category, code, reason = FR_FRANCHISE_EN_BASE
        assert category == "E"
        assert code == "VATEX-FR-FRANCHISE"
        assert reason == "TVA non applicable, article 293 B du CGI"

    def test_invoice_line_forwards_exemption_reason_to_tax(self):
        category, code, reason = FR_FRANCHISE_EN_BASE
        line = InvoiceLine(quantity=1, unit_code="EA", price=100,
                           item_name="prestation", currency="EUR",
                           tax_percent=0, tax_category=category, tax_name="E0",
                           exemption_reason=reason, exemption_reason_code=code)
        assert line.tax.exemption_reason == reason
        assert line.tax.exemption_reason_code == code

    def test_invoice_line_without_reason_uses_category_default(self):
        line = InvoiceLine(quantity=1, unit_code="EA", price=100,
                           item_name="marchandise", currency="EUR",
                           tax_percent=0, tax_category="K", tax_name="K0")
        assert line.tax.exemption_reason_code == "VATEX-EU-IC"
