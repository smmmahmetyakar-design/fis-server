#!/usr/bin/env python3
"""
Fiş PDF Aktarım Aracı (Receipt PDF Transfer Tool)
FastAPI web application for converting receipt images/PDFs to accounting CSV

Port: 8093
URL: http://dellserver:8093 or http://100.74.86.128:8093
"""

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import os
import csv
import tempfile
from pathlib import Path
from typing import List, Dict, Tuple
import shutil
from datetime import datetime
import re

# Try to import OCR
try:
    import pytesseract
    from PIL import Image
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# Try to import PDF support
try:
    from pdf2image import convert_from_path
    HAS_PDF = True
except ImportError:
    HAS_PDF = False

app = FastAPI(title="Fiş PDF Aktarım Aracı", version="1.0.0")

# Configuration
UPLOAD_DIR = Path("/tmp/fis-uploads")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR = Path("/tmp/fis-outputs")
OUTPUT_DIR.mkdir(exist_ok=True)

# Allowed file types
ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.pdf', '.JPG', '.JPEG', '.PNG', '.PDF'}

# CSV field headers
CSV_HEADERS = [
    'EVRAK TARİHİ',
    'EVRAK NO',
    'TCKN/VKN',
    'SOYADI ÜNVAN',
    'ADI DEVAMI',
    'TUTAR',
    'KDV ORANI',
    'KDV TUTARI',
    'TOPLAM TUTAR'
]


class ReceiptProcessor:
    """Process receipt images and extract data."""

    def __init__(self):
        self.receipts: List[Dict] = []
        self.skipped: List[Tuple[str, str]] = []
        self.extraction_notes: List[Dict] = []

    def process_file(self, file_path: str, filename: str) -> None:
        """
        Process a single uploaded file. A PDF may contain MULTIPLE separate
        receipts/invoices (one per page) - each page is OCR'd and parsed as
        its OWN receipt row. Combining all pages into one blob of text before
        parsing was the earlier bug: it mixed numbers from different receipts
        together into garbage values.
        """
        try:
            if not HAS_OCR:
                self.skipped.append((filename, "OCR not available - install pytesseract"))
                return

            if filename.lower().endswith('.pdf'):
                if not HAS_PDF:
                    self.skipped.append((filename, "PDF support not available - install pdf2image"))
                    return

                try:
                    images = convert_from_path(file_path)
                except Exception as e:
                    self.skipped.append((filename, f"Failed to convert PDF: {str(e)[:50]}"))
                    return

                if not images:
                    self.skipped.append((filename, "PDF has no pages"))
                    return

                total_pages = len(images)
                for page_num, image in enumerate(images, start=1):
                    page_text = pytesseract.image_to_string(image, lang='tur+eng')
                    if total_pages > 1:
                        page_label = f"{filename} (sayfa {page_num}/{total_pages})"
                    else:
                        page_label = filename
                    self._process_page_text(page_text, page_label)
            else:
                image = Image.open(file_path)
                text = pytesseract.image_to_string(image, lang='tur+eng')
                self._process_page_text(text, filename)

        except Exception as e:
            self.skipped.append((filename, f"Error: {str(e)[:50]}"))

    def _process_page_text(self, text: str, label: str) -> None:
        """Parse the OCR text of ONE page/image as ONE receipt and record the result."""
        if not text.strip():
            self.skipped.append((label, "Text extraction failed - too blurry or unreadable"))
            return

        receipt, field_count, found_fields = self._parse_text(text, label)

        clean_text = text.strip()
        if len(clean_text) > 700:
            text_preview = clean_text[:400].replace('\n', ' | ') + " [...] " + clean_text[-300:].replace('\n', ' | ')
        else:
            text_preview = clean_text.replace('\n', ' | ')

        if receipt:
            self.receipts.append(receipt)
            # Debug note for successful extractions too, so we can verify
            # parsed values against the raw OCR text without needing a
            # failure to see what OCR actually read.
            self.extraction_notes.append({
                'file': label,
                'field_count': field_count,
                'values': {k: v for k, v in receipt.items() if v},
                'ocr_preview': text_preview
            })
        else:
            missing = [f for f in CSV_HEADERS if f not in found_fields]
            reason = f"{field_count}/9 alan bulundu. Eksik: {', '.join(missing)}. OCR metni: \"{text_preview}\""
            self.skipped.append((label, reason))

    def _parse_text(self, text: str, source: str) -> Dict | None:
        """
        Parse OCR text into receipt fields using regex patterns.

        Extracts 9 required fields from Turkish receipts:
        - EVRAK TARİHİ (receipt date DD/MM/YYYY)
        - EVRAK NO (receipt number)
        - TCKN/VKN (tax ID)
        - SOYADI ÜNVAN (vendor name/company)
        - ADI DEVAMI (vendor name continuation)
        - TUTAR (amount before tax)
        - KDV ORANI (VAT rate %)
        - KDV TUTARI (VAT amount)
        - TOPLAM TUTAR (total amount)
        """

        receipt = {field: '' for field in CSV_HEADERS}
        field_count = 0
        found_fields = []

        # Pattern 1: Date in DD/MM/YYYY format (handles both dashes and slashes)
        date_match = re.search(r'\b(\d{1,2})[-/](\d{1,2})[-/](\d{4})\b', text)
        if date_match:
            day, month, year = date_match.groups()
            # Validate date range
            if 1 <= int(day) <= 31 and 1 <= int(month) <= 12 and 2000 <= int(year) <= 2099:
                receipt['EVRAK TARİHİ'] = f"{day:>02s}/{month:>02s}/{year}"
                field_count += 1
                found_fields.append('EVRAK TARİHİ')

        # Pattern 2: Tax ID (TCKN/VKN) - 10 or 11 digit number
        vkn_match = re.search(r'\b(\d{10,11})\b', text)
        if vkn_match:
            receipt['TCKN/VKN'] = vkn_match.group(1)
            field_count += 1
            found_fields.append('TCKN/VKN')

        # Pattern 3: Receipt number. Try SPECIFIC, reliable labels first
        # ("Fiş No", "Fis No" - OCR often drops the cedilla, "Evrak No",
        # "Belge No"). A bare "No" is tried only as a last resort, because
        # it also matches an address street number like "NO:52/206" - and
        # since that appears EARLIER in the text (in the header/address
        # block) than the real "FİŞ NO" line, a naive re.search() for a
        # bare "No" alternative would grab the address number instead.
        # Require digits only for the captured value here: on some receipts
        # OCR reading order puts an unrelated label ("SAAT 16:28") right
        # after "Fiş No:" instead of the actual number (which prints
        # further down/right on the physical receipt), and a permissive
        # [A-Za-z0-9]+ capture would grab that unrelated word instead of
        # failing cleanly and falling through to the digit-only fallbacks
        # below.
        no_match = re.search(
            r'(?:Belge\s*No|Evrak\s*No|Fi[şs]\s*No)[:\s]*(\d{2,10})',
            text, re.IGNORECASE
        )
        if no_match:
            receipt['EVRAK NO'] = no_match.group(1).strip()
            field_count += 1
            found_fields.append('EVRAK NO')
        else:
            # Fallback 1: the "Fiş/Evrak/Belge No" keyword was found, but not
            # immediately followed by digits - OCR reading order sometimes
            # puts an unrelated label ("SAAT 16:28") right after it instead,
            # with the real number printed a bit further down/right. Search
            # a nearby window after the keyword instead of the whole
            # document (which risks grabbing a phone number or address
            # number that happens to appear earlier in the text). Require
            # 3+ digits and reject anything glued to a ":" so a time like
            # "16:28" (2-digit components) is never mistaken for it.
            found_no = False
            kw_match = re.search(r'(?:Belge|Evrak|Fi[şs])\s*No', text, re.IGNORECASE)
            if kw_match:
                window = text[kw_match.end():kw_match.end() + 80]
                for cand in re.finditer(r'(?<![:\d])(\d{3,8})(?![:\d])', window):
                    receipt['EVRAK NO'] = cand.group(1)
                    field_count += 1
                    found_fields.append('EVRAK NO')
                    found_no = True
                    break

            if not found_no:
                # Fallback 2: a bare "No" keyword, but skip it if what
                # follows looks like part of an address (contains "/", as
                # in "52/206")
                no_bare = re.search(r'\bNo[:\s]*([A-Za-z0-9\-/]{2,20})', text, re.IGNORECASE)
                if no_bare and '/' not in no_bare.group(1):
                    receipt['EVRAK NO'] = no_bare.group(1).strip()
                    field_count += 1
                    found_fields.append('EVRAK NO')
                    found_no = True

            if not found_no:
                # Fallback 3 (last resort): a standalone 3-8 digit number
                # anywhere in the document
                no_fallback = re.search(r'(?:^|\s)(\d{3,8})(?:\s|$)', text, re.MULTILINE)
                if no_fallback:
                    receipt['EVRAK NO'] = no_fallback.group(1)
                    field_count += 1
                    found_fields.append('EVRAK NO')

        # Pattern 4: Vendor name - typically appears early in receipt
        # Look for lines with company/shop names (usually after header/date)
        lines = text.split('\n')
        vendor_found = False
        for i, line in enumerate(lines):
            line_clean = line.strip()
            # Skip very short lines and common noise
            if 3 < len(line_clean) < 100 and not re.match(r'^\d+[.,]?\d*$', line_clean):
                # Check if line looks like vendor name (not just numbers or common receipt headers)
                if not any(header in line_clean.upper() for header in ['TOPLAM', 'TUTAR', 'KDV', 'EVRAK', 'TARİH']):
                    receipt['SOYADI ÜNVAN'] = line_clean[:50]  # First 50 chars
                    found_fields.append('SOYADI ÜNVAN')
                    # Try to get continuation from next line if it exists
                    if i + 1 < len(lines) and 3 < len(lines[i + 1].strip()) < 50:
                        next_line = lines[i + 1].strip()
                        if not any(c.isdigit() for c in next_line[:5]):  # Doesn't start with numbers
                            receipt['ADI DEVAMI'] = next_line[:50]
                            found_fields.append('ADI DEVAMI')
                    vendor_found = True
                    field_count += 1
                    break

        # Pattern 5: Amounts - look for currency values
        # Turkish format: number with comma as decimal (1.234,56) or just
        # (1234,56 - no thousands separator at all, common on smaller
        # receipts). Allows a stray OCR space before the decimals, e.g.
        # "650, 00".
        # The trailing (?!\d) is critical: without it, a number like
        # "48,900" (which is really 48.900 with the thousands DOT misread
        # as a comma by OCR) gets wrongly truncated to a "valid-looking"
        # match "48,90" by grabbing only the first 2 of the 3 digits after
        # the comma. Rejecting any match immediately followed by another
        # digit forces us to correctly recognize this as NOT a clean
        # decimal amount, rather than silently extracting the wrong value.
        # Allow an optional stray OCR space after the thousands-separator
        # dot too (e.g. "4. 720,00"), not just before the final decimal
        # digits - otherwise the leading "4." thousands group gets dropped
        # entirely and "4.720,00" (4720) is silently misread as "720,00"
        # (720), a full order-of-magnitude error.
        # The thousands-group separator is accepted as EITHER "." (correct
        # Turkish notation) OR "," (a very common OCR misread of the dot on
        # thermal receipts) - but the final decimal separator right before
        # the last 2 digits must always be a literal ",". This is still
        # strict about digit counts (exactly 3 per thousands group, exactly
        # 2 decimals) and the trailing (?!\d) guard, so it doesn't swallow
        # an unrelated run of digits - it only widens which punctuation
        # character is tolerated as the grouping separator.
        AMOUNT_TOKEN = r'(?:\d{1,3}(?:[.,]\s?\d{3})+,\s?\d{2}|\d{1,6},\s?\d{2})(?!\d)'
        # KDV is frequently OCR'd as "KDY", "KOV", or with a stray space
        # inserted inside it ("K DY") on blurry/crumpled thermal receipts.
        KDV_KW = r'K\s*D\s*[VY]|VAT|Vergi'
        # TOPLAM is frequently truncated/garbled down to just "TOP".
        TOPLAM_KW = r'Genel\s*Toplam|TOPLAM\s*TUTAR|Toplam|Total|TOP\b'

        # NOTE: deliberately NOT matching "Ara Toplam" here. On some
        # receipts "ARATOPLAM" is the VAT-INCLUSIVE subtotal (same value as
        # TOP/TOPLAM), not the pre-tax base - treating it as TUTAR caused
        # TUTAR to wrongly equal TOPLAM, blocking the KDV reconciliation
        # below from ever subtracting the tax out. Only an explicit
        # "Tutar" label is trusted for the pre-tax amount; otherwise it's
        # derived from TOPLAM - KDV further down.
        amount_pattern = rf'(?:Tutar)[^\d]{{0,8}}({AMOUNT_TOKEN})'
        amount_matches = re.findall(amount_pattern, text, re.IGNORECASE)

        total_pattern = rf'(?:{TOPLAM_KW})[^\d]{{0,8}}({AMOUNT_TOKEN})'
        total_matches = re.findall(total_pattern, text, re.IGNORECASE)

        # Cross-check: OCR frequently inserts stray digit+comma noise right
        # next to a keyword (e.g. "TOP ¥5, 650,00" gets misread as "5, 65"
        # instead of "650,00"), so a keyword-adjacent match is NOT trusted
        # blindly. The actual total amount tends to print 2-3 times on the
        # receipt (item line, totals line, payment line) and a properly
        # repeated value is stronger evidence than a keyword-adjacent one -
        # it takes priority whenever it exists.
        bare_amounts = re.findall(rf'\b({AMOUNT_TOKEN})\b', text)
        if bare_amounts:
            normalized_bare = [a.replace(' ', '') for a in bare_amounts]
            from collections import Counter
            most_common, count = Counter(normalized_bare).most_common(1)[0]
            if count >= 2:
                total_matches = [most_common]

        # VAT rate - only counted when a literal % sign is actually present
        # next to the number, so we don't mistake a chunk of an amount for
        # a rate (e.g. "833,33" must never be read as a "%83" rate).
        vat_rate_pattern = (
            rf'(?:{KDV_KW})[^\d%]{{0,6}}%\s*(\d{{1,2}}(?:[.,]\d+)?)'
            rf'|(?:{KDV_KW})[^\d%]{{0,6}}(\d{{1,2}})\s*%'
        )
        vat_rate_matches = re.findall(vat_rate_pattern, text, re.IGNORECASE)

        # VAT amount - some receipts print the KDV amount directly instead
        # of (or in addition to) the rate, e.g. "KDV: 833,33".
        vat_amount_pattern = rf'(?:{KDV_KW})[^\d%]{{0,8}}({AMOUNT_TOKEN})'
        vat_amount_matches = re.findall(vat_amount_pattern, text, re.IGNORECASE)

        # Extract amount (before tax), if explicitly labeled
        if amount_matches:
            amt = self._normalize_amount(amount_matches[-1])
            if amt:
                receipt['TUTAR'] = amt
                field_count += 1
                found_fields.append('TUTAR')

        # Extract total amount
        if total_matches:
            total = self._normalize_amount(total_matches[-1])
            if total:
                receipt['TOPLAM TUTAR'] = total
                field_count += 1
                found_fields.append('TOPLAM TUTAR')

        # Extract VAT rate (only if a real % sign was found)
        if vat_rate_matches:
            # findall with multiple groups returns tuples; take whichever
            # side of the alternation matched (the non-empty one)
            last = vat_rate_matches[-1]
            vat_str = (last[0] or last[1]) if isinstance(last, tuple) else last
            vat_rate = re.sub(r'[^\d.,]', '', vat_str.strip())
            if vat_rate:
                receipt['KDV ORANI'] = f"%{vat_rate}"
                field_count += 1
                found_fields.append('KDV ORANI')

        # Extract VAT amount directly, if labeled (independent of rate)
        direct_kdv_amount = None
        if vat_amount_matches:
            direct_kdv_amount = self._normalize_amount(vat_amount_matches[-1])

        # --- Reconcile whichever combination of TUTAR / TOPLAM / KDV ORANI /
        # KDV TUTARI we actually found, filling in the rest by calculation.
        try:
            tutar_val = self._to_float(receipt['TUTAR']) if receipt.get('TUTAR') else None
            toplam_val = self._to_float(receipt['TOPLAM TUTAR']) if receipt.get('TOPLAM TUTAR') else None
            rate_val = None
            if receipt.get('KDV ORANI'):
                rate_val = float(receipt['KDV ORANI'].replace('%', '').replace(',', '.'))
            kdv_val = self._to_float(direct_kdv_amount) if direct_kdv_amount else None

            if kdv_val is None and tutar_val is not None and rate_val is not None:
                # Have base amount + rate -> derive KDV amount and total
                kdv_val = round(tutar_val * rate_val / 100, 2)
                if toplam_val is None:
                    toplam_val = round(tutar_val + kdv_val, 2)
            elif kdv_val is not None and toplam_val is not None and tutar_val is None:
                # Have total + KDV amount -> derive base amount and rate
                tutar_val = round(toplam_val - kdv_val, 2)
                if rate_val is None and tutar_val > 0:
                    rate_val = round(kdv_val / tutar_val * 100, 2)
            elif kdv_val is not None and tutar_val is not None and toplam_val is None:
                # Have base amount + KDV amount -> derive total and rate
                toplam_val = round(tutar_val + kdv_val, 2)
                if rate_val is None and tutar_val > 0:
                    rate_val = round(kdv_val / tutar_val * 100, 2)
            elif toplam_val is not None and tutar_val is None:
                # No VAT info at all - assume TUTAR equals TOPLAM (no split available)
                tutar_val = toplam_val
            elif tutar_val is not None and toplam_val is None:
                toplam_val = tutar_val

            def fmt(v):
                return f"{v:.2f}".replace('.', ',') if v is not None else ''

            if tutar_val is not None and 'TUTAR' not in found_fields:
                receipt['TUTAR'] = fmt(tutar_val)
                field_count += 1
                found_fields.append('TUTAR')
            if toplam_val is not None and 'TOPLAM TUTAR' not in found_fields:
                receipt['TOPLAM TUTAR'] = fmt(toplam_val)
                field_count += 1
                found_fields.append('TOPLAM TUTAR')
            if rate_val is not None and 'KDV ORANI' not in found_fields:
                receipt['KDV ORANI'] = f"%{rate_val:g}"
                field_count += 1
                found_fields.append('KDV ORANI')
            if kdv_val is not None and 'KDV TUTARI' not in found_fields:
                receipt['KDV TUTARI'] = fmt(kdv_val)
                field_count += 1
                found_fields.append('KDV TUTARI')
        except (ValueError, ZeroDivisionError, TypeError):
            pass

        # Consider a receipt valid if the core identifying fields are present:
        # date, receipt no, tax ID, vendor name, and at least one amount.
        # KDV rate/amount are bonus - OCR on crumpled/blurry thermal receipts
        # often can't recover them, but the accountant can fill those by hand
        # if the rest of the row is correct.
        core_fields = ['EVRAK TARİHİ', 'EVRAK NO', 'TCKN/VKN', 'SOYADI ÜNVAN']
        has_core = sum(1 for f in core_fields if f in found_fields) >= 3
        has_amount = 'TUTAR' in found_fields or 'TOPLAM TUTAR' in found_fields
        if has_core and has_amount:
            return receipt, field_count, found_fields

        return None, field_count, found_fields

    def _normalize_amount(self, amount_str: str) -> str:
        """
        Clean up an OCR-extracted Turkish-format amount and return it in
        Turkish format (comma decimal, dot thousands) for the output CSV,
        e.g. "650,00" or "1.234,56". Returns None if it isn't a valid number.
        """
        amount_str = amount_str.strip().replace(' ', '')

        comma_count = amount_str.count(',')
        dot_count = amount_str.count('.')

        if comma_count >= 1 and dot_count >= 1:
            # Mixed separators - whichever one appears LAST is the real
            # decimal separator; anything before it is a thousands group.
            if amount_str.rindex(',') > amount_str.rindex('.'):
                # Comma is last -> already Turkish format, keep as-is.
                pass
            else:
                # Dot is last -> English format (comma=thousands,
                # dot=decimal); convert to Turkish.
                amount_str = amount_str.replace(',', '').replace('.', ',')
        elif comma_count >= 2:
            # Multiple commas, no dots at all - OCR misread every "."
            # thousands separator as "," too (e.g. "5,650,00" meaning
            # 5.650,00). Every comma except the LAST one is a thousands
            # group; only the last is the real decimal separator.
            parts = amount_str.split(',')
            amount_str = '.'.join(parts[:-1]) + ',' + parts[-1]
        elif dot_count >= 2:
            # Symmetric case: multiple dots, no commas - last dot is decimal.
            parts = amount_str.split('.')
            amount_str = '.'.join(parts[:-1]) + ',' + parts[-1]
        elif dot_count == 1 and comma_count == 0:
            # Only a single dot - ambiguous, but for a receipt amount this
            # is far more likely a decimal point (English-style OCR read)
            # than a thousands separator, so treat it as the decimal.
            amount_str = amount_str.replace('.', ',')
        # else: exactly one comma and no dot - already correct Turkish format

        # Validate it's a real number by parsing it as a float
        try:
            self._parse_turkish_amount(amount_str)
            return amount_str
        except ValueError:
            return None

    def _parse_turkish_amount(self, amount_str: str) -> float:
        """Parse a Turkish-format amount string ("1.234,56" or "650,00") into a float."""
        return float(amount_str.replace('.', '').replace(',', '.'))

    def _to_float(self, amount_str: str) -> float:
        """Public helper: parse a Turkish-format amount string into a float."""
        return self._parse_turkish_amount(amount_str)

    def add_manual_receipt(self, data: Dict[str, str]) -> bool:
        """Add manually-entered receipt."""
        receipt = {field: data.get(field, '') for field in CSV_HEADERS}

        if not receipt.get('EVRAK TARİHİ'):
            return False

        self.receipts.append(receipt)
        return True

    def export_csv(self, output_path: str) -> None:
        """Export receipts to CSV with UTF-8 BOM and semicolon delimiter."""
        with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(
                f,
                fieldnames=CSV_HEADERS,
                delimiter=';',
                quoting=csv.QUOTE_MINIMAL
            )
            writer.writeheader()
            writer.writerows(self.receipts)

    def get_summary(self) -> Dict:
        """Get processing summary."""
        return {
            'extracted': len(self.receipts),
            'skipped': len(self.skipped),
            'skipped_details': [{'file': f, 'reason': r} for f, r in self.skipped],
            'extracted_details': self.extraction_notes
        }


# Routes

@app.get("/", response_class=HTMLResponse)
async def home():
    """Serve the main web interface."""
    return HTML_TEMPLATE


@app.post("/api/process")
async def process_receipts(files: List[UploadFile] = File(...)):
    """Process uploaded receipt files and return CSV."""

    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    # Validate file types
    for file in files:
        if Path(file.filename).suffix.lower() not in ALLOWED_EXTENSIONS:
            raise HTTPException(
                status_code=400,
                detail=f"File type not allowed: {file.filename}. Allowed: JPEG, PNG, PDF"
            )

    processor = ReceiptProcessor()

    # Save uploaded files and process
    with tempfile.TemporaryDirectory() as tmpdir:
        for file in files:
            file_path = Path(tmpdir) / file.filename

            # Save uploaded file
            with open(file_path, 'wb') as f:
                content = await file.read()
                f.write(content)

            # Process file
            processor.process_file(str(file_path), file.filename)

        # Generate output CSV
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"fis_aktarim_{timestamp}.csv"
        output_path = OUTPUT_DIR / output_filename

        processor.export_csv(str(output_path))

        # Return summary and CSV download link
        return JSONResponse({
            'success': True,
            'summary': processor.get_summary(),
            'csv_filename': output_filename,
            'csv_url': f"/download/{output_filename}"
        })


@app.get("/download/{filename}")
async def download_csv(filename: str):
    """Download generated CSV file."""
    file_path = OUTPUT_DIR / filename

    if not file_path.exists():
        raise HTTPException(status_code=404, detail="File not found")

    return FileResponse(
        path=file_path,
        filename=filename,
        media_type="text/csv; charset=utf-8"
    )


@app.get("/health")
async def health():
    """Health check endpoint."""
    return {
        'status': 'ok',
        'service': 'Fiş PDF Aktarım Aracı',
        'ocr_available': HAS_OCR,
        'pdf_support': HAS_PDF
    }


# HTML Template

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="tr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Fiş PDF Aktarım Aracı</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }

        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }

        .container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            max-width: 600px;
            width: 100%;
            padding: 40px;
        }

        header {
            text-align: center;
            margin-bottom: 30px;
        }

        h1 {
            color: #333;
            font-size: 28px;
            margin-bottom: 8px;
        }

        .subtitle {
            color: #666;
            font-size: 14px;
        }

        .upload-area {
            border: 2px dashed #667eea;
            border-radius: 8px;
            padding: 40px 20px;
            text-align: center;
            cursor: pointer;
            transition: all 0.3s;
            background: #f8f9ff;
            margin-bottom: 20px;
        }

        .upload-area:hover {
            border-color: #764ba2;
            background: #f0f1ff;
        }

        .upload-area.dragover {
            border-color: #764ba2;
            background: #f0f1ff;
            transform: scale(1.02);
        }

        .upload-icon {
            font-size: 48px;
            margin-bottom: 16px;
        }

        .upload-text {
            color: #333;
            font-weight: 500;
            margin-bottom: 4px;
        }

        .upload-hint {
            color: #999;
            font-size: 13px;
        }

        #fileInput {
            display: none;
        }

        .file-list {
            margin-bottom: 20px;
        }

        .file-item {
            background: #f5f5f5;
            padding: 12px;
            border-radius: 6px;
            margin-bottom: 8px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 14px;
        }

        .file-item .filename {
            color: #333;
            flex: 1;
        }

        .file-item .size {
            color: #999;
            font-size: 12px;
        }

        .file-item .remove {
            background: #ff4444;
            color: white;
            border: none;
            border-radius: 4px;
            padding: 4px 12px;
            cursor: pointer;
            font-size: 12px;
        }

        .button-group {
            display: flex;
            gap: 12px;
            margin-bottom: 20px;
        }

        button {
            flex: 1;
            padding: 12px 24px;
            border: none;
            border-radius: 6px;
            font-size: 16px;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s;
        }

        .btn-process {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }

        .btn-process:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.3);
        }

        .btn-process:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        .btn-clear {
            background: #f0f0f0;
            color: #333;
        }

        .btn-clear:hover {
            background: #e0e0e0;
        }

        .results {
            display: none;
            background: #f9f9f9;
            border-radius: 8px;
            padding: 24px;
            border-left: 4px solid #667eea;
        }

        .results.show {
            display: block;
        }

        .results h3 {
            color: #333;
            margin-bottom: 16px;
            font-size: 18px;
        }

        .stats {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
            margin-bottom: 16px;
        }

        .stat {
            background: white;
            padding: 16px;
            border-radius: 6px;
            border: 1px solid #eee;
        }

        .stat-number {
            font-size: 24px;
            font-weight: bold;
            color: #667eea;
        }

        .stat-label {
            font-size: 13px;
            color: #666;
            margin-top: 4px;
        }

        .skipped-list {
            max-height: 200px;
            overflow-y: auto;
            margin-top: 12px;
        }

        .skipped-item {
            font-size: 13px;
            color: #666;
            padding: 8px;
            background: white;
            border-radius: 4px;
            margin-bottom: 4px;
            border-left: 3px solid #ff9800;
        }

        .download-btn {
            background: #4CAF50;
            color: white;
            width: 100%;
            padding: 14px;
            text-decoration: none;
            text-align: center;
            border-radius: 6px;
            display: block;
            margin-top: 16px;
            font-weight: 500;
        }

        .download-btn:hover {
            background: #45a049;
        }

        .error {
            background: #ffebee;
            color: #c62828;
            padding: 16px;
            border-radius: 6px;
            margin-bottom: 20px;
            border-left: 4px solid #c62828;
        }

        .loading {
            display: none;
            text-align: center;
            padding: 20px;
        }

        .loading.show {
            display: block;
        }

        .spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
            margin: 0 auto 16px;
        }

        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }

        .info-box {
            background: #e3f2fd;
            border-left: 4px solid #2196F3;
            padding: 12px;
            border-radius: 4px;
            font-size: 13px;
            color: #1565c0;
            margin-bottom: 20px;
        }
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>📄 Fiş PDF Aktarım Aracı</h1>
            <p class="subtitle">Fatura/Fiş fotoğraflarını muhasebe CSV'sine dönüştür</p>
        </header>

        <div class="info-box">
            💡 JPEG, PNG veya PDF dosyalarını yükle. 9 alandan muhasebe CSV'si oluşturulur.
        </div>

        <div id="errorMessage" class="error" style="display: none;"></div>

        <div class="upload-area" id="uploadArea">
            <div class="upload-icon">📤</div>
            <div class="upload-text">Fiş fotoğraflarını buraya sürükle veya tıkla</div>
            <div class="upload-hint">JPEG, PNG, PDF — max 50 dosya</div>
            <input type="file" id="fileInput" multiple accept=".jpg,.jpeg,.png,.pdf">
        </div>

        <div class="file-list" id="fileList"></div>

        <div class="button-group">
            <button class="btn-process" id="processBtn" onclick="processFiles()">CSV'ye Çevir</button>
            <button class="btn-clear" onclick="clearFiles()">Temizle</button>
        </div>

        <div class="loading" id="loading">
            <div class="spinner"></div>
            <p>Fişler işleniyor...</p>
        </div>

        <div class="results" id="results">
            <h3>✓ İşlem Tamamlandı</h3>
            <div class="stats">
                <div class="stat">
                    <div class="stat-number" id="extractedCount">0</div>
                    <div class="stat-label">Çıkarılan Fiş</div>
                </div>
                <div class="stat">
                    <div class="stat-number" id="skippedCount">0</div>
                    <div class="stat-label">Atlanmış Fiş</div>
                </div>
            </div>
            <div id="skippedDetails"></div>
            <a id="downloadLink" class="download-btn" href="#">CSV'yi İndir</a>
        </div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileList = document.getElementById('fileList');
        const processBtn = document.getElementById('processBtn');
        const loading = document.getElementById('loading');
        const results = document.getElementById('results');
        const errorMessage = document.getElementById('errorMessage');

        let selectedFiles = [];

        // Drag and drop
        uploadArea.addEventListener('click', () => fileInput.click());
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            handleFiles(e.dataTransfer.files);
        });

        fileInput.addEventListener('change', (e) => {
            handleFiles(e.target.files);
        });

        function handleFiles(files) {
            selectedFiles = Array.from(files);
            updateFileList();
            errorMessage.style.display = 'none';
            results.classList.remove('show');
        }

        function updateFileList() {
            fileList.innerHTML = '';
            selectedFiles.forEach((file, index) => {
                const item = document.createElement('div');
                item.className = 'file-item';
                item.innerHTML = `
                    <span class="filename">${file.name}</span>
                    <span class="size">${(file.size / 1024).toFixed(1)} KB</span>
                    <button class="file-item remove" onclick="removeFile(${index})">Sil</button>
                `;
                fileList.appendChild(item);
            });

            processBtn.disabled = selectedFiles.length === 0;
        }

        function removeFile(index) {
            selectedFiles.splice(index, 1);
            updateFileList();
        }

        function clearFiles() {
            selectedFiles = [];
            fileInput.value = '';
            updateFileList();
            results.classList.remove('show');
            errorMessage.style.display = 'none';
        }

        async function processFiles() {
            if (selectedFiles.length === 0) {
                showError('Lütfen dosya seçin');
                return;
            }

            loading.classList.add('show');
            errorMessage.style.display = 'none';

            try {
                const formData = new FormData();
                selectedFiles.forEach(file => {
                    formData.append('files', file);
                });

                const response = await fetch('/api/process', {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const error = await response.json();
                    throw new Error(error.detail || 'İşlem başarısız');
                }

                const data = await response.json();
                showResults(data);

            } catch (error) {
                showError(error.message);
            } finally {
                loading.classList.remove('show');
            }
        }

        function showResults(data) {
            const summary = data.summary;
            document.getElementById('extractedCount').textContent = summary.extracted;
            document.getElementById('skippedCount').textContent = summary.skipped;

            const skippedDetails = document.getElementById('skippedDetails');
            let detailsHtml = '';

            if (summary.extracted_details && summary.extracted_details.length > 0) {
                detailsHtml += '<div class="skipped-list"><strong>Çıkarılan fişler:</strong>';
                summary.extracted_details.forEach(item => {
                    const vals = Object.entries(item.values).map(([k, v]) => `${k}: ${v}`).join(', ');
                    detailsHtml += `<div class="skipped-item" style="border-left-color:#4CAF50"><strong>${item.file}</strong> (${item.field_count}/9): ${vals}<br><span style="color:#999;font-size:11px">OCR: "${item.ocr_preview}"</span></div>`;
                });
                detailsHtml += '</div>';
            }

            if (summary.skipped_details.length > 0) {
                detailsHtml += '<div class="skipped-list"><strong>Atlanmış dosyalar:</strong>';
                summary.skipped_details.forEach(item => {
                    detailsHtml += `<div class="skipped-item"><strong>${item.file}</strong>: ${item.reason}</div>`;
                });
                detailsHtml += '</div>';
            }

            skippedDetails.innerHTML = detailsHtml;

            const downloadLink = document.getElementById('downloadLink');
            downloadLink.href = data.csv_url;
            downloadLink.download = data.csv_filename;

            results.classList.add('show');
        }

        function showError(message) {
            errorMessage.textContent = '❌ Hata: ' + message;
            errorMessage.style.display = 'block';
        }
    </script>
</body>
</html>
"""

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8093)
