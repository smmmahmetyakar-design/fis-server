"""e-Fatura XML'inin görüntüsü: gömülü XSLT ile HTML, şablon yoksa sade sayfa, ZIP'teki HTML.
python tests/test_efatura_gorunum.py — sahte veri."""
import base64
import io
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app import ubl_fatura  # noqa: E402

XSLT = b"""<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:inv="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
  xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
  xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2">
  <xsl:output method="html" encoding="UTF-8"/>
  <xsl:template match="/">
    <html><body><h1>SABLON FATURA <xsl:value-of select="//inv:Invoice/cbc:ID"/></h1>
    <p><xsl:value-of select="//cac:AccountingSupplierParty//cac:PartyName/cbc:Name"/></p>
    </body></html>
  </xsl:template>
</xsl:stylesheet>"""


def _ubl(xslt: bytes | None) -> bytes:
    ek = ""
    if xslt:
        ek = f"""<cac:AdditionalDocumentReference><cbc:ID>1</cbc:ID><cbc:IssueDate>2026-11-05</cbc:IssueDate>
          <cbc:DocumentType>XSLT</cbc:DocumentType><cac:Attachment>
          <cbc:EmbeddedDocumentBinaryObject characterSetCode="UTF-8" encodingCode="Base64" filename="ORN.xslt"
            mimeCode="application/xml">{base64.b64encode(xslt).decode()}</cbc:EmbeddedDocumentBinaryObject>
          </cac:Attachment></cac:AdditionalDocumentReference>"""
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
  xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
  xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2">
  <cbc:ProfileID>TICARIFATURA</cbc:ProfileID><cbc:ID>ORN2026000000001</cbc:ID>
  <cbc:UUID>11111111-2222-3333-4444-555555555555</cbc:UUID><cbc:IssueDate>2026-11-05</cbc:IssueDate>
  <cbc:InvoiceTypeCode>SATIS</cbc:InvoiceTypeCode><cbc:Note>Sahte test faturası</cbc:Note>
  <cbc:DocumentCurrencyCode>TRY</cbc:DocumentCurrencyCode>{ek}
  <cac:AccountingSupplierParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="VKN">1234567890</cbc:ID>
    </cac:PartyIdentification><cac:PartyName><cbc:Name>ÖRNEK SATICI A.Ş.</cbc:Name></cac:PartyName></cac:Party>
  </cac:AccountingSupplierParty>
  <cac:AccountingCustomerParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="VKN">9876543210</cbc:ID>
    </cac:PartyIdentification><cac:PartyName><cbc:Name>ÖRNEK ALICI LTD.</cbc:Name></cac:PartyName></cac:Party>
  </cac:AccountingCustomerParty>
  <cac:TaxTotal><cbc:TaxAmount currencyID="TRY">200.00</cbc:TaxAmount><cac:TaxSubtotal>
    <cbc:TaxableAmount currencyID="TRY">1000.00</cbc:TaxableAmount><cbc:TaxAmount currencyID="TRY">200.00</cbc:TaxAmount>
    <cbc:Percent>20</cbc:Percent><cac:TaxCategory><cac:TaxScheme><cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme>
    </cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal>
  <cac:LegalMonetaryTotal><cbc:LineExtensionAmount currencyID="TRY">1000.00</cbc:LineExtensionAmount>
    <cbc:TaxInclusiveAmount currencyID="TRY">1200.00</cbc:TaxInclusiveAmount>
    <cbc:PayableAmount currencyID="TRY">1200.00</cbc:PayableAmount></cac:LegalMonetaryTotal>
  <cac:InvoiceLine><cbc:ID>1</cbc:ID><cbc:InvoicedQuantity unitCode="C62">10</cbc:InvoicedQuantity>
    <cbc:LineExtensionAmount currencyID="TRY">1000.00</cbc:LineExtensionAmount>
    <cac:TaxTotal><cac:TaxSubtotal><cbc:Percent>20</cbc:Percent><cac:TaxCategory><cac:TaxScheme>
      <cbc:TaxTypeCode>0015</cbc:TaxTypeCode></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal>
    <cac:Item><cbc:Name>DENEME MALZEMESİ</cbc:Name></cac:Item>
    <cac:Price><cbc:PriceAmount currencyID="TRY">100.00</cbc:PriceAmount></cac:Price></cac:InvoiceLine>
</Invoice>""".encode("utf-8")


def test_gomulu_xslt():
    h = ubl_fatura.html_gorunum(_ubl(XSLT)).decode("utf-8")
    assert "SABLON FATURA ORN2026000000001" in h and "ÖRNEK SATICI A.Ş." in h, h
    # yerel dosya okumaya çalışan şablon çalıştırılmaz (sade sayfaya düşer), dosya içeriği sızmaz
    kotu = XSLT.replace(b"</body>", b"<p><xsl:value-of select=\"document('/etc/passwd')\"/></p></body>")
    h = ubl_fatura.html_gorunum(_ubl(kotu)).decode("utf-8")
    assert "root:" not in h and "ORN2026000000001" in h, h


def test_sablonsuz_sade_sayfa():
    h = ubl_fatura.html_gorunum(_ubl(None)).decode("utf-8")
    for parca in ("ORN2026000000001", "05.11.2026", "ÖRNEK SATICI A.Ş.", "ÖRNEK ALICI LTD.",
                  "DENEME MALZEMESİ", "1.200,00", "Sahte test faturası"):
        assert parca in h, parca


def test_zip_icindeki_html():
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("f.xml", _ubl(XSLT))
        z.writestr("f.html", "<html><body>ENTEGRATOR GORUNUMU</body></html>")
    assert b"ENTEGRATOR GORUNUMU" in ubl_fatura.html_gorunum(b.getvalue())
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("f.xml", _ubl(XSLT))
    assert b"SABLON FATURA" in ubl_fatura.html_gorunum(b.getvalue())


if __name__ == "__main__":
    test_gomulu_xslt(); test_sablonsuz_sade_sayfa(); test_zip_icindeki_html()
    print("3/3 kontrol gecti")
