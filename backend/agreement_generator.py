"""
backend/agreement_generator.py

Reads the clean Word template, substitutes all placeholders with
user-supplied values, and exports a PDF via LibreOffice.

Returns: (pdf_bytes: bytes, filename: str)
"""

import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

TEMPLATE = Path(__file__).parent.parent / "template" / "Subcontractor_Agreement_Template_clean.docx"
EXPORTS  = Path(__file__).parent.parent / "exports"
EXPORTS.mkdir(exist_ok=True)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _fmt_date(value: str) -> str:
    """Try to parse and return MM/DD/YYYY; fall back to the original string."""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime(value.strip(), fmt).strftime("%m/%d/%Y")
        except ValueError:
            pass
    return value


def _fmt_amount(value: str) -> str:
    """Format as $X,XXX.XX if the value looks numeric."""
    try:
        n = float(value.replace(",", "").replace("$", ""))
        return f"${n:,.2f}"
    except (ValueError, AttributeError):
        return value



def _replace_nth(text: str, old: str, new: str, n: int) -> str:
    """Replace the n-th (1-based) occurrence of *old* with *new*."""
    pos = -1
    for _ in range(n):
        pos = text.find(old, pos + 1)
        if pos == -1:
            return text
    return text[:pos] + new + text[pos + len(old):]


_RUN_20 = '<w:r w:rsidRPr="00BE323A"><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr><w:t>'


def _populate_xml(xml: str, f: dict) -> str:
    """Apply all placeholder substitutions to document.xml content."""

    # ── Simple one-to-one replacements ─────────────────────────────────────
    simple = {
        "[KHPXXX]":                  f["project_id"]        or "[KHPXXX]",
        "[Address]":                 f["project_address"]    or "[Address]",
        "[MM/DD/YYYY]":              _fmt_date(f["agreement_date"])   if f["agreement_date"]   else "[MM/DD/YYYY]",
        "[Project Start Date]":      _fmt_date(f["start_date"])       if f["start_date"]       else "[Project Start Date]",
        "[Project Completion Date]": _fmt_date(f["completion_date"])  if f["completion_date"]  else "[Project Completion Date]",
        "[Subcontractor Name]":      f["subcontractor_name"] or "[Subcontractor Name]",
        "[Company Name]":            f["company_name"]       or "[Company Name]",
        "[Amount]":                  _fmt_amount(f["total_amount"])   if f["total_amount"]     else "[Amount]",
        "[TITLE]":                   f["signatory_title"]    or "[TITLE]",
        "[License #]":               f["license_number"]     or "[License #]",
    }
    for old, new in simple.items():
        xml = xml.replace(old, new)

    # ── [NAME] appears twice: 1st = KAEDIX signatory, 2nd = subcontractor ──
    xml = _replace_nth(xml, "[NAME]", f["signatory_name"]    or "[NAME]", 1)
    xml = _replace_nth(xml, "[NAME]", f["subcontractor_name"] or "[NAME]", 1)

    # ── [Email] appears twice: 1st = KAEDIX signatory, 2nd = subcontractor ─
    xml = _replace_nth(xml, "[Email]", f["signatory_email"] or "[Email]", 1)
    xml = _replace_nth(xml, "[Email]", f["sub_email"]       or "[Email]", 1)

    # ── Notes → Appendix A ──────────────────────────────────────────────────
    if f["notes"]:
        xml = xml.replace(
            "<w:t>See attached pages.</w:t>",
            f'<w:t xml:space="preserve">See attached pages.\n\nScope Summary: {f["notes"]}</w:t>',
        )

    return xml


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def generate_agreement_pdf(
    project_id: str,
    project_address: str,
    agreement_date: str,
    start_date: str,
    completion_date: str,
    subcontractor_name: str,
    company_name: str,
    license_number: str,
    sub_email: str,
    total_amount: str,
    signatory_name: str,
    signatory_title: str,
    signatory_email: str,
    notes: str,
) -> tuple[bytes, str]:
    """
    Populate the Word template and convert to PDF.
    Returns (pdf_bytes, filename).
    """

    fields = {
        "project_id":        project_id.strip(),
        "project_address":   project_address.strip(),
        "agreement_date":    agreement_date.strip(),
        "start_date":        start_date.strip(),
        "completion_date":   completion_date.strip(),
        "subcontractor_name": subcontractor_name.strip(),
        "company_name":      company_name.strip(),
        "license_number":    license_number.strip(),
        "sub_email":         sub_email.strip(),
        "total_amount":      total_amount.strip(),
        "signatory_name":    signatory_name.strip(),
        "signatory_title":   signatory_title.strip(),
        "signatory_email":   signatory_email.strip(),
        "notes":             notes.strip(),
    }

    # Build output filename
    date_str = ""
    if fields["agreement_date"]:
        try:
            date_str = datetime.strptime(
                _fmt_date(fields["agreement_date"]), "%m/%d/%Y"
            ).strftime("%Y%m%d")
        except ValueError:
            date_str = datetime.today().strftime("%Y%m%d")
    else:
        date_str = datetime.today().strftime("%Y%m%d")

    pid      = fields["project_id"] or "KHPXXX"
    company  = fields["company_name"] or fields["subcontractor_name"] or "Subcontractor"
    safe_co  = re.sub(r"[^\w\s-]", "", company).strip().replace(" ", "_")
    stem     = f"{date_str}_{pid}_Subcontractor_Agreement_{safe_co}"
    docx_out = EXPORTS / f"{stem}.docx"
    pdf_out  = EXPORTS / f"{stem}.pdf"

    # --- 1. Populate XML inside a temp copy of the template ----------------
    with tempfile.TemporaryDirectory() as tmp:
        tmp_docx = Path(tmp) / "output.docx"

        # Copy template into tmp
        shutil.copy(TEMPLATE, tmp_docx)

        # Read, modify, and rewrite document.xml inside the ZIP
        with zipfile.ZipFile(tmp_docx, "r") as zin:
            names  = zin.namelist()
            files  = {name: zin.read(name) for name in names}

        doc_xml = files["word/document.xml"].decode("utf-8")
        doc_xml = _populate_xml(doc_xml, fields)
        files["word/document.xml"] = doc_xml.encode("utf-8")

        with zipfile.ZipFile(tmp_docx, "w", zipfile.ZIP_DEFLATED) as zout:
            for name, data in files.items():
                zout.writestr(name, data)

        # Copy filled docx to exports
        shutil.copy(tmp_docx, docx_out)

        # --- 2. Convert DOCX → PDF via LibreOffice -------------------------
        result = subprocess.run(
            [
                "soffice",
                "--headless",
                "--convert-to", "pdf",
                "--outdir", str(EXPORTS),
                str(docx_out),
            ],
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice conversion failed:\n{result.stderr}")

    # soffice names the output after the docx stem
    soffice_out = EXPORTS / f"{stem}.pdf"
    if not soffice_out.exists():
        raise FileNotFoundError(f"Expected PDF not found: {soffice_out}")

    pdf_bytes = soffice_out.read_bytes()
    return pdf_bytes, pdf_out.name
