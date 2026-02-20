"""
CSE Disclosure Tracker v9 — Headless (for automation)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Runs without GUI. Perfect for:
  - GitHub Actions (free, every hour)
  - PythonAnywhere scheduled tasks
  - Cron jobs on any server
  - Google Cloud Functions

Usage: python cse_tracker_headless.py
"""

import time
import os
import sys
import requests
import io
import re
import json
from datetime import datetime
from collections import defaultdict
import PyPDF2

try:
    from pdf2image import convert_from_bytes
    import pytesseract
    HAS_OCR = True
except ImportError:
    HAS_OCR = False

# ─── CONFIGURATION ────────────────────────────────────────────────────

# Can be overridden by environment variables
SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "service_account.json")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "CSE Disclosures Tracker")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AIzaSyB7tgMltxeeV-p4jmhn8s-tMWmCgF_tXJM")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "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")
MAX_DISCLOSURES = 100
SUMMARY_DELAY = 6

CSE_API = "https://www.cse.lk/api/"
CSE_CDN = "https://cdn.cse.lk/"
HTTP_HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
    'Referer': 'https://www.cse.lk/',
    'Origin': 'https://www.cse.lk',
}


def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")


# ─── STRUCTURED SUMMARY BUILDERS ────────────────────────────────────

def build_structured_summary(ann_data):
    base = ann_data.get("reqBaseAnnouncement", {})
    dtype = base.get("dType", "")
    company = base.get("companyName", "")
    remarks = base.get("remarks", "") or ""

    if dtype == "CashDividendWithDates":
        return _build_dividend_summary(base, company)
    elif dtype == "DealingsByDirectors":
        return _build_dealings_summary(base, company)
    elif dtype == "AppointmentOfDirectors":
        return _build_appointment_summary(base, company)
    elif dtype in ("ResignationOfDirectors", "ResignationOfChp"):
        return _build_resignation_summary(base, company, dtype)
    elif dtype in ("AppointOfChp",):
        return _build_chairperson_summary(base, company)
    elif dtype == "RightsIssue":
        return _build_rights_summary(base, company)
    elif dtype == "ExtraOrdinaryGeneralMeetingInitial":
        return _build_egm_summary(base, company)
    elif dtype in ("AgmInitial",):
        return _build_agm_summary(base, company)
    return None


def _build_dividend_summary(base, company):
    div_type = []
    if base.get("firstAndFinal"): div_type.append("first & final")
    elif base.get("finalDividend"): div_type.append("final")
    elif base.get("typeFirstInt"): div_type.append("first interim")
    elif base.get("typeSecondInt"): div_type.append("second interim")
    elif base.get("typeThirdInt"): div_type.append("third interim")
    elif base.get("typeFourthInt"): div_type.append("fourth interim")
    dtype_str = " ".join(div_type) if div_type else ""

    voting = base.get("votingDivPerShare")
    non_voting = base.get("nonVotingDivPerShare")
    fy = base.get("financialYear", "")
    xd = base.get("xd", "")
    payment = base.get("payment", "")
    agm = base.get("agm", "")
    approval = "subject to shareholder approval" if base.get("shrHolderApproval") == "R" else ""

    parts = [f"{company} declared a {dtype_str} cash dividend".strip()]
    if voting: parts.append(f"of Rs. {voting}/- per voting share")
    if non_voting and non_voting > 0: parts.append(f"and Rs. {non_voting}/- per non-voting share")
    if fy: parts.append(f"for FY {fy}")
    if approval: parts.append(f"({approval})")
    summary = " ".join(parts) + "."
    dates = []
    if agm: dates.append(f"AGM: {agm}")
    if xd: dates.append(f"XD: {xd}")
    if payment: dates.append(f"Payment: {payment}")
    if dates: summary += " " + ", ".join(dates) + "."
    return summary


def _build_dealings_summary(base, company):
    nature = base.get("natureOfDir", "")
    txns = base.get("directorTransactions", [])

    if "refer attachment" in nature.lower() or "refer attachment" in (base.get("relInterestAccountName", "") or "").lower():
        return None

    nature = re.sub(r'\s*Directors?\s*$', '', nature).strip()
    if not nature: nature = "Director"

    parts = [f"{company}: Dealings by {nature}."]
    by_type = defaultdict(lambda: {"qty": 0, "total_value": 0, "prices": [], "dates": []})
    for tx in txns:
        tx_type = tx.get("transType", "Transaction")
        qty = tx.get("quantity", 0) or 0
        price = tx.get("price", 0) or 0
        tx_date = tx.get("transactionDate", "")
        by_type[tx_type]["qty"] += qty
        by_type[tx_type]["total_value"] += qty * price
        if price and price not in by_type[tx_type]["prices"]: by_type[tx_type]["prices"].append(price)
        if tx_date and tx_date not in by_type[tx_type]["dates"]: by_type[tx_type]["dates"].append(tx_date)

    for tx_type, d in by_type.items():
        qty = d["qty"]
        prices = d["prices"]
        dates = d["dates"]
        qty_str = f"{int(qty):,}" if qty == int(qty) else f"{qty:,.2f}"
        if len(prices) == 1: price_str = f"at Rs. {prices[0]}"
        elif prices:
            avg = d["total_value"] / qty if qty else 0
            price_str = f"at avg Rs. {avg:,.2f}"
        else: price_str = ""
        date_str = dates[0] if len(dates) == 1 else f"{dates[0]}-{dates[-1]}" if dates else ""
        parts.append(f"{tx_type}: {qty_str} shares {price_str} on {date_str}.".strip())
    return " ".join(parts)


def _build_appointment_summary(base, company):
    dirs = base.get("dirList", [])
    if not dirs: return None
    parts = [f"{company}:"]
    for d in dirs:
        nature = d.get("natureOfDir", "Director")
        if nature and "director" not in nature.lower() and "chairperson" not in nature.lower():
            nature = nature + " Director"
        eff = d.get("effectiveDate", "")
        shares = d.get("numberOfShares", 0)
        parts.append(f"Appointed {nature}, effective {eff}.")
        if shares: parts.append(f"Holds {int(shares):,} shares.")
    return " ".join(parts)


def _build_resignation_summary(base, company, dtype):
    role = "Chairperson" if "Chp" in dtype else "Director"
    remarks = base.get("remarks", "") or ""
    date_match = re.search(r'w\.?e\.?f\.?\s*(\d{1,2}[./]\d{1,2}[./]\d{4})', remarks)
    eff_date = f", effective {date_match.group(1)}" if date_match else ""
    return f"{company}: Resignation of {role}{eff_date}."


def _build_chairperson_summary(base, company):
    remarks = base.get("remarks", "") or ""
    date_match = re.search(r'(?:w\.?e\.?f\.?|effective|from)\s*[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4}|\d{1,2}\s+\w+\s+\d{4})', remarks, re.IGNORECASE)
    eff_date = f", effective {date_match.group(1)}" if date_match else ""
    return f"{company}: Appointment of Chairperson{eff_date}."


def _build_rights_summary(base, company):
    remarks = base.get("remarks", "") or ""
    voting = base.get("numOfVotingShrsIssued")
    consideration = base.get("votingShareConsideration")
    xr = base.get("xr", "")
    parts = [f"{company} — Rights Issue."]
    if voting: parts.append(f"{int(voting):,} voting shares to be issued.")
    if consideration: parts.append(f"At Rs. {consideration}/- per share.")
    if xr: parts.append(f"XR date: {xr}.")
    if remarks: parts.append(remarks)
    return " ".join(parts)


def _build_egm_summary(base, company):
    date = base.get("dateOfEgm", "")
    venue = base.get("venue", "")
    time_str = base.get("time", "")
    res = base.get("resToBePassed", "")
    parts = [f"{company} — EGM scheduled for {date}"]
    if time_str: parts.append(f"at {time_str}")
    if venue: parts.append(f"at {venue}")
    summary = " ".join(parts).strip() + "."
    if res:
        res_clean = re.sub(r'\s+', ' ', res).strip()
        summary += f" Resolutions: {res_clean}."
    return summary


def _build_agm_summary(base, company):
    agm = base.get("agm", "") or base.get("dateOfAgm", "")
    remarks = base.get("remarks", "") or ""
    parts = [f"{company} — AGM"]
    if agm: parts.append(f"scheduled for {agm}")
    summary = " ".join(parts).strip() + "."
    if remarks: summary += f" {remarks}"
    return summary


# ─── AI SUMMARIZATION ────────────────────────────────────────────────

JUNK_PATTERNS = [
    'dear madam', 'dear sir', 'dear madan',
    'yours faithfully', 'yours sincerely',
    'chief regulatory officer', 'west block', 'echelon square',
    'world trade centre', 'p w corporate', 'heed oltrce', 'tel:', 'fax:',
]


def pre_clean_for_ai(raw_text):
    if not raw_text: return ""
    text = raw_text
    for p in [r'Yours\s+(faithfully|sincerely|truly).*',
              r'BY\s+ORDER\s+OF\s+THE\s+BOARD.*',
              r'For\s+and\s+on\s+behalf\s+of.*']:
        text = re.sub(p, '', text, flags=re.IGNORECASE | re.DOTALL)
    remove = [
        r'Dear\s+(Sir|Madam|Madan-?r?|Sirs?)[\s,]*',
        r'Ms\.?\s+Nilupa\s+Perar?a.{0,100}',
        r'Mrs\.?\s+Nilupa\s+Perar?a.{0,100}',
        r'Chief\s+Regulatory\s+Officer.{0,100}',
        r'Colombo\s+Stock\s+Exc[a-z]*.{0,100}',
        r'Echelon\s+Square.{0,60}', r'World\s*\'?[Tt]rade\s+Centr?e.{0,60}',
        r'West\s+Block.{0,60}',
        r'#?\d+[-/]?\d*,?\s*\w+\s+(Road|Street|Lane|Mawatha|Place).{0,80}',
        r'Colombo\s*\d{1,2}.{0,40}', r'Sri\s+Lanka\.?',
        r'Tel(?:ephone)?:?\s*[\+\d\s\-\(\)\']{5,30}',
        r'Fax:?\s*[\+\d\s\-\(\)\']{5,30}',
        r'E-?mail:?\s*\S+@\S+', r'P\.?O\.?\s*Box\s*\d+',
        r'P\s*W\s*(?:Corporate|Gorporate)\s*Secretarial.{0,80}',
        r'M&S\s*Managers\s*&\s*Secretaries.{0,80}',
        r'JACEY\s*&?\s*(?:COMPANY|GOMPANY).{0,80}',
        r'JULIUS\s*&?\s*CREASY.{0,80}',
    ]
    for p in remove:
        text = re.sub(p, ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'[{}\[\]|\\@#$^~`]', '', text)
    text = re.sub(r'[!.]{2,}', '.', text)
    text = re.sub(r'\s+', ' ', text).strip()
    return text[:3000]


def is_good_summary(summary):
    if not summary or len(summary) < 20: return False
    s = summary.lower()
    for p in JUNK_PATTERNS:
        if p in s: return False
    if len(summary) > 800: return False
    if len(summary.split()) < 5: return False
    lazy = ['key details were not provided', 'key details are not provided',
            'unfortunately', 'does not contain sufficient', 'not enough information',
            'cannot extract specific', 'the provided text does not', 'the given text',
            'here are the specific facts', 'nilupa perera']
    for l in lazy:
        if l in s: return False
    return True


def is_fallback_summary(summary):
    if not summary: return True
    s = summary.lower()
    bad = ['i don\'t see', 'unfortunately', 'does not contain', 'not enough information',
           'cannot extract', 'the provided text', 'key details were not provided',
           'here are the specific facts', 'nilupa perera', 'company registration number']
    for b in bad:
        if b in s: return True
    return False


def ai_summarize(raw_text, company_name="", subject=""):
    if not raw_text or len(raw_text) < 30: return None
    cleaned = pre_clean_for_ai(raw_text)
    if len(cleaned) < 30:
        cleaned = re.sub(r'\s+', ' ', raw_text).strip()[:2000]
        if len(cleaned) < 30: return None

    system_msg = (
        "You extract key facts from Colombo Stock Exchange corporate disclosures. "
        "Write 2-3 sentences with SPECIFIC details. "
        "Include: share quantities, rupee amounts, percentages, dates, positions/titles. "
        "NEVER include person names — use position/title instead. "
        "Focus ONLY on what the company is announcing. "
        "NEVER include: addresses, phone/fax, emails, signatures, person names. "
        "NEVER start with 'Here are the facts' — just write the summary directly."
    )
    user_msg = f"Company: {company_name}\nCategory: {subject}\n\nExtract specific facts:\n{cleaned[:2000]}"

    # Groq
    if GROQ_API_KEY:
        for attempt in range(3):
            try:
                resp = requests.post("https://api.groq.com/openai/v1/chat/completions",
                    headers={"Content-Type": "application/json", "Authorization": f"Bearer {GROQ_API_KEY}"},
                    json={"model": "llama-3.3-70b-versatile",
                          "messages": [{"role": "system", "content": system_msg},
                                       {"role": "user", "content": user_msg}],
                          "max_tokens": 200, "temperature": 0.1}, timeout=25)
                if resp.status_code == 200:
                    choices = resp.json().get("choices", [])
                    if choices:
                        summary = choices[0].get("message", {}).get("content", "").strip()
                        summary = re.sub(r'\*+', '', summary).strip()
                        summary = re.sub(r'\n+', ' ', summary).strip()
                        for px in [r'^here are the specific facts[^:]*:\s*', r'^summary:\s*']:
                            summary = re.sub(px, '', summary, flags=re.IGNORECASE).strip()
                        if is_good_summary(summary): return summary
                        elif attempt < 2:
                            time.sleep(3); continue
                        elif summary and len(summary) > 30:
                            return summary
                elif resp.status_code == 429:
                    log(f"  Groq rate limited, waiting..."); time.sleep(20)
                else: break
            except Exception as e:
                log(f"  Groq error: {e}"); break

    # Gemini fallback
    if GEMINI_API_KEY:
        try:
            resp = requests.post(
                f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key={GEMINI_API_KEY}",
                headers={"Content-Type": "application/json"},
                json={"contents": [{"parts": [{"text": f"{system_msg}\n\n{user_msg}"}]}],
                      "generationConfig": {"maxOutputTokens": 200, "temperature": 0.1}}, timeout=25)
            if resp.status_code == 200:
                cands = resp.json().get("candidates", [])
                if cands:
                    parts = cands[0].get("content", {}).get("parts", [])
                    if parts:
                        summary = re.sub(r'\*+', '', parts[0].get("text", "")).strip()
                        if is_good_summary(summary): return summary
        except Exception as e:
            log(f"  Gemini error: {e}")
    return None


# ─── GOOGLE SHEETS ────────────────────────────────────────────────────

class GoogleManager:
    def __init__(self):
        import gspread
        from google.oauth2.service_account import Credentials

        scopes = ["https://www.googleapis.com/auth/spreadsheets",
                   "https://www.googleapis.com/auth/drive.readonly"]

        # Support both file and env var (for GitHub Actions)
        sa_json = os.environ.get("SERVICE_ACCOUNT_KEY")
        if sa_json:
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                f.write(sa_json)
                sa_path = f.name
            creds = Credentials.from_service_account_file(sa_path, scopes=scopes)
            os.unlink(sa_path)
            log("✓ Authenticated via env var")
        else:
            creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_JSON, scopes=scopes)
            log("✓ Authenticated via file")

        self.gc = gspread.authorize(creds)
        self.spreadsheet = self.gc.open(SPREADSHEET_NAME)
        self.worksheet = self.spreadsheet.sheet1
        log(f"✓ Opened: {SPREADSHEET_NAME}")

        # Ensure headers
        headers = ["Date", "Time", "Company", "Subject", "Description",
                    "AI Summary", "PDF Link", "PDF Count", "Unique Key"]
        first_row = self.worksheet.row_values(1)
        if not first_row or not any(first_row):
            self.worksheet.update('A1:I1', [headers])

    def get_existing_keys(self):
        try:
            keys = self.worksheet.col_values(9)
            return set(keys[1:])
        except:
            return set()

    def add_rows(self, rows):
        if not rows: return 0
        existing = self.get_existing_keys()
        new_rows = []
        for r in rows:
            key = r.get("Unique Key", "")
            if key and key not in existing:
                pdf = r.get("PDF Link", "")
                pdf_cell = f'=HYPERLINK("{pdf}", "View PDF")' if pdf else ""
                new_rows.append([
                    r.get("Date", ""), r.get("Time", ""), r.get("Company", ""),
                    r.get("Subject", ""), r.get("Description", ""),
                    r.get("AI Summary", ""), pdf_cell,
                    r.get("PDF Count", 0), key,
                ])
        if new_rows:
            self.worksheet.append_rows(new_rows, value_input_option='USER_ENTERED')
        return len(new_rows)

    def get_all_rows(self):
        return self.worksheet.get_all_values()

    def update_cell(self, row, col, value):
        self.worksheet.update_cell(row, col, value)


# ─── CSE API ──────────────────────────────────────────────────────────

def fetch_announcements():
    resp = requests.post(CSE_API + "approvedAnnouncement",
        headers={**HTTP_HEADERS, 'Content-Type': 'application/x-www-form-urlencoded'}, timeout=20)
    if resp.status_code != 200:
        log(f"✗ API {resp.status_code}")
        return []
    items = resp.json().get("approvedAnnouncements", [])
    log(f"✓ {len(items)} announcements from API")
    return items[:MAX_DISCLOSURES]


def get_detail(ann_id):
    try:
        resp = requests.post(CSE_API + "getAnnouncementById",
            data={"announcementId": ann_id},
            headers={**HTTP_HEADERS, 'Content-Type': 'application/x-www-form-urlencoded'}, timeout=15)
        if resp.status_code == 200 and len(resp.content) > 5:
            return resp.json()
    except: pass
    try:
        resp = requests.post(CSE_API + "getGeneralAnnouncementById",
            data={"announcementId": ann_id},
            headers={**HTTP_HEADERS, 'Content-Type': 'application/x-www-form-urlencoded'}, timeout=15)
        if resp.status_code == 200 and resp.text != "{}":
            return resp.json()
    except: pass
    return None


def download_pdf_text(file_url, base_url=None):
    full_url = ((base_url or CSE_CDN) + file_url).replace(' ', '%20')
    try:
        resp = requests.get(full_url, headers={'User-Agent': HTTP_HEADERS['User-Agent'],
            'Referer': 'https://www.cse.lk/'}, timeout=20)
        if resp.status_code != 200 or resp.content[:5] != b'%PDF-':
            return None
        reader = PyPDF2.PdfReader(io.BytesIO(resp.content))
        text = ""
        for page in reader.pages:
            text += (page.extract_text() or "") + "\n"
        text = text.strip()
        if len(text) > 30:
            return text
        if HAS_OCR:
            try:
                images = convert_from_bytes(resp.content, dpi=200)
                ocr_text = ""
                for img in images[:3]:
                    ocr_text += pytesseract.image_to_string(img) + "\n"
                if len(ocr_text.strip()) > 30:
                    return ocr_text.strip()
            except: pass
    except: pass
    return None


# ─── MAIN RUN ─────────────────────────────────────────────────────────

def main():
    start = time.time()
    log("=" * 50)
    log("CSE DISCLOSURE TRACKER v9 — Headless")
    log("=" * 50)

    # Connect to Google
    gm = GoogleManager()
    existing_keys = gm.get_existing_keys()
    log(f"  Existing: {len(existing_keys)} entries")

    # Phase 1: Fetch from API
    log("─" * 40)
    log("PHASE 1: API SCRAPE")
    log("─" * 40)

    items = fetch_announcements()
    if not items:
        log("No announcements found, exiting")
        return

    # Cache all items for later ID lookup
    api_cache = items

    # Find new items
    new_items = []
    for item in items:
        company = item.get("company", "")
        date_str = item.get("dateOfAnnouncement", "")
        created = item.get("createdDate", 0)
        try:
            dt = datetime.fromtimestamp(created / 1000)
            time_str = dt.strftime("%I:%M:%S %p")
        except:
            time_str = ""
        ukey = f"{date_str}|{time_str}|{company}"
        if ukey not in existing_keys and not any(k.startswith(ukey) for k in existing_keys):
            new_items.append(item)

    log(f"  {len(new_items)} NEW items")

    if not new_items:
        log("  Nothing new to scrape")
    else:
        rows = []
        for item in new_items:
            ann_id = item.get("announcementId")
            company = item.get("company", "")
            category = item.get("announcementCategory", "")
            date_str = item.get("dateOfAnnouncement", "")
            remarks = item.get("remarks") or ""
            created = item.get("createdDate", 0)

            try:
                dt = datetime.fromtimestamp(created / 1000)
                time_str = dt.strftime("%I:%M:%S %p")
                if not date_str:
                    date_str = dt.strftime("%d %b %Y").upper()
            except:
                time_str = ""

            ukey = f"{date_str}|{time_str}|{company}"
            detail = get_detail(ann_id)
            pdf_links = []
            if detail:
                docs = detail.get("reqAnnouncementDocs", [])
                for doc in docs:
                    fu = doc.get("fileUrl", "")
                    bu = doc.get("baseUrl", CSE_CDN)
                    if fu:
                        pdf_links.append((bu or CSE_CDN) + fu)

            if len(pdf_links) > 1:
                for pi, link in enumerate(pdf_links):
                    rows.append({"Date": date_str, "Time": time_str, "Company": company,
                                 "Subject": category, "Description": f"PDF {pi+1} of {len(pdf_links)}",
                                 "AI Summary": "", "PDF Link": link, "PDF Count": len(pdf_links),
                                 "Unique Key": f"{ukey}|PDF{pi+1}", "_detail": detail, "_ann_id": ann_id})
            else:
                rows.append({"Date": date_str, "Time": time_str, "Company": company,
                             "Subject": category, "Description": remarks[:200],
                             "AI Summary": "", "PDF Link": pdf_links[0] if pdf_links else "",
                             "PDF Count": len(pdf_links), "Unique Key": ukey,
                             "_detail": detail, "_ann_id": ann_id})

            log(f"  [{len(rows)}] {company[:40]} — {category} ({len(pdf_links)} PDF)")
            time.sleep(0.3)

        # Save to sheet
        log("─" * 40)
        log("SAVING TO SHEET")
        log("─" * 40)
        clean_rows = [{k: v for k, v in r.items() if not k.startswith('_')} for r in rows]
        count = gm.add_rows(clean_rows)
        log(f"  {count} rows saved")

    # Phase 2: Fill missing summaries
    log("─" * 40)
    log("PHASE 2: AI SUMMARIES")
    log("─" * 40)

    all_rows = gm.get_all_rows()
    needs_summary = []
    for i, row in enumerate(all_rows[1:], start=2):
        summary = row[5] if len(row) > 5 else ""
        if not summary or is_fallback_summary(summary):
            needs_summary.append((i, row))

    log(f"  {len(needs_summary)} rows need summaries")

    for idx, (row_num, row) in enumerate(needs_summary):
        company = row[2] if len(row) > 2 else ""
        subject = row[3] if len(row) > 3 else ""
        pdf_link = row[6] if len(row) > 6 else ""

        log(f"  [{idx+1}/{len(needs_summary)}] {company[:40]} — {subject}")
        summary = None

        # Strategy 1: Structured summary from API
        for item in api_cache:
            if (item.get("company", "").strip() == company.strip() and
                item.get("announcementCategory", "").strip() == subject.strip()):
                detail = get_detail(item.get("announcementId"))
                if detail:
                    summary = build_structured_summary(detail)
                    if summary:
                        log(f"    ✓ Structured")
                break

        # Strategy 2: PDF + AI
        if not summary and pdf_link:
            url = None
            m = re.search(r'HYPERLINK\("([^"]+)"', pdf_link)
            if m: url = m.group(1)
            elif pdf_link.startswith("http"): url = pdf_link

            if url:
                file_url = url.replace(CSE_CDN, "").replace("https://cdn.cse.lk/", "")
                pdf_text = download_pdf_text(file_url)
                if pdf_text:
                    summary = ai_summarize(pdf_text, company, subject)
                    if summary:
                        log(f"    ✓ AI summary")

        # Strategy 3: Fallback
        if not summary:
            summary = f"{company} — {subject}."
            log(f"    ⚠ Fallback")

        try:
            gm.update_cell(row_num, 6, summary)
        except Exception as e:
            log(f"    ✗ Write error: {e}")

        if idx < len(needs_summary) - 1:
            time.sleep(SUMMARY_DELAY)

    elapsed = time.time() - start
    log("=" * 50)
    log(f"✓ DONE in {elapsed:.0f}s")
    log("=" * 50)


if __name__ == "__main__":
    main()
