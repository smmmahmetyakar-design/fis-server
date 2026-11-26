"""e-Fatura UBL XML okuyucu: sentetik fatura ile (gerçek müşteri verisi yok)."""
import io
import zipfile

from app import ubl_fatura as u

XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
 xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
 xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
<cbc:ProfileID>TICARIFATURA</cbc:ProfileID><cbc:ID>ABC2026000000001</cbc:ID>
<cbc:UUID>11111111-2222-3333-4444-555555555555</cbc:UUID>
<cbc:IssueDate>2026-09-30</cbc:IssueDate><cbc:InvoiceTypeCode>TEVKIFAT</cbc:InvoiceTypeCode>
<cbc:DocumentCurrencyCode>TRY</cbc:DocumentCurrencyCode>
<cac:AccountingSupplierParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="VKN">1234567890</cbc:ID></cac:PartyIdentification>
<cac:PartyName><cbc:Name>DENEME NAKLIYE LTD STI</cbc:Name></cac:PartyName></cac:Party></cac:AccountingSupplierParty>
<cac:AccountingCustomerParty><cac:Party><cac:PartyName><cbc:Name>ALICI AS</cbc:Name></cac:PartyName></cac:Party></cac:AccountingCustomerParty>
<cac:TaxTotal><cbc:TaxAmount>2100</cbc:TaxAmount>
<cac:TaxSubtotal><cbc:TaxableAmount>10000</cbc:TaxableAmount><cbc:TaxAmount>2000</cbc:TaxAmount><cbc:Percent>20</cbc:Percent>
<cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>
<cac:TaxSubtotal><cbc:TaxableAmount>1000</cbc:TaxableAmount><cbc:TaxAmount>100</cbc:TaxAmount><cbc:Percent>10</cbc:Percent>
<cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal>
</cac:TaxTotal>
<cac:WithholdingTaxTotal><cbc:TaxAmount>400</cbc:TaxAmount><cac:TaxSubtotal><cbc:TaxAmount>400</cbc:TaxAmount>
<cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>624</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:WithholdingTaxTotal>
<cac:LegalMonetaryTotal><cbc:TaxInclusiveAmount>13100</cbc:TaxInclusiveAmount><cbc:PayableAmount>12700</cbc:PayableAmount></cac:LegalMonetaryTotal>
<cac:InvoiceLine><cbc:LineExtensionAmount>10000</cbc:LineExtensionAmount><cac:Item><cbc:Name>YUK TASIMA</cbc:Name></cac:Item></cac:InvoiceLine>
</Invoice>'''


def _zip(*dosyalar):
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        for ad, veri in dosyalar:
            z.writestr(ad, veri)
    return b.getvalue()


def test_tek_xml_alis():
    f = u.oku_bytes(XML, "alis", "a.xml")[0]
    assert f["fatura_no"] == "ABC2026000000001" and f["tarih"] == "2026-09-30"
    assert f["cari_ad"] == "DENEME NAKLIYE LTD STI" and f["vkn"] == "1234567890"
    assert f["kalemler"] == [{"oran": 20, "matrah": 10000.0, "kdv": 2000.0},
                             {"oran": 10, "matrah": 1000.0, "kdv": 100.0}]
    assert f["tevkifat"] == 400.0 and f["tevkifat_kod"] == "624" and f["tur"] == "TEVKIFAT"
    assert f["toplam"] == 12700.0 and f["kaynak"] == "xml" and not f["eksik"] and not f["uyari"]


def test_satis_yonunde_karsi_taraf_alici():
    assert u.oku_bytes(XML, "satis")[0]["cari_ad"] == "ALICI AS"


def test_gib_zip_html_ile_ve_cift_fatura():
    z = _zip(("f.xml", XML), ("f.html", b"<html>fatura</html>"), ("kopya/f.xml", XML))
    assert u.ubl_mu(z)
    assert len(u.oku_bytes(z, "alis")) == 1          # aynı UUID bir kez
    assert u.html_gorunum(z) == b"<html>fatura</html>"
    assert u.fatura_nolari(z) == ["ABC2026000000001", "ABC2026000000001"]


def test_office_zip_fatura_sayilmaz():
    xlsx = _zip(("[Content_Types].xml", b"<Types/>"), ("xl/workbook.xml", b"<workbook/>"))
    assert not u.ubl_mu(xlsx)
