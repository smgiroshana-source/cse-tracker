"""
CSE Disclosure Tracker v9 — Unified
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Pure API. No Selenium. One file.

Auto-detects:
  - Mac/PC with display → opens GUI
  - GitHub Actions / server → runs headless

Usage:
  python cse_tracker_v9.py            ← GUI on Mac
  python cse_tracker_v9.py --headless ← force headless
"""

import time, os, sys, threading, webbrowser, requests, io, re, json
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

SERVICE_ACCOUNT_JSON = os.environ.get("SERVICE_ACCOUNT_JSON", "service_account.json")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "CSE Disclosures Tracker")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "AIzaSyB7tgMltxeeV-p4jmhn8s-tMWmCgF_tXJM")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "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")
MAX_DISCLOSURES = 100
SUMMARY_DELAY = 6
CSE_API = "https://www.cse.lk/api/"
CSE_CDN = "https://cdn.cse.lk/"
HTTP_HEADERS = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36',
    'Referer': 'https://www.cse.lk/', 'Origin': 'https://www.cse.lk'}
SETUP_GUIDE = "\n  pip install gspread google-auth google-api-python-client PyPDF2 requests\n  No Chrome/Selenium needed!\n"

# ─── STRUCTURED SUMMARIES ────────────────────────────────────────────
def build_structured_summary(ann_data):
    base = ann_data.get("reqBaseAnnouncement", {}); dtype = base.get("dType", ""); company = base.get("companyName", "")
    if dtype == "CashDividendWithDates": return _build_dividend(base, company)
    elif dtype == "DealingsByDirectors": return _build_dealings(base, company)
    elif dtype == "AppointmentOfDirectors": return _build_appointment(base, company)
    elif dtype in ("ResignationOfDirectors", "ResignationOfChp"): return _build_resignation(base, company, dtype)
    elif dtype in ("AppointOfChp",): return _build_chairperson(base, company)
    elif dtype == "RightsIssue": return _build_rights(base, company)
    elif dtype == "ExtraOrdinaryGeneralMeetingInitial": return _build_egm(base, company)
    elif dtype in ("AgmInitial",): return _build_agm(base, company)
    return None

def _build_dividend(b, co):
    dt = []
    if b.get("firstAndFinal"): dt.append("first & final")
    elif b.get("finalDividend"): dt.append("final")
    elif b.get("typeFirstInt"): dt.append("first interim")
    elif b.get("typeSecondInt"): dt.append("second interim")
    elif b.get("typeThirdInt"): dt.append("third interim")
    elif b.get("typeFourthInt"): dt.append("fourth interim")
    ds = " ".join(dt) if dt else ""
    v = b.get("votingDivPerShare"); nv = b.get("nonVotingDivPerShare")
    fy=b.get("financialYear",""); xd=b.get("xd",""); pay=b.get("payment",""); agm=b.get("agm","")
    ap = "subject to shareholder approval" if b.get("shrHolderApproval")=="R" else ""
    p = [f"{co} declared a {ds} cash dividend".strip()]
    if v: p.append(f"of Rs. {v}/- per voting share")
    if nv and nv > 0: p.append(f"and Rs. {nv}/- per non-voting share")
    if fy: p.append(f"for FY {fy}")
    if ap: p.append(f"({ap})")
    s = " ".join(p) + "."
    d = []
    if agm: d.append(f"AGM: {agm}")
    if xd: d.append(f"XD: {xd}")
    if pay: d.append(f"Payment: {pay}")
    if d: s += " " + ", ".join(d) + "."
    return s

def _build_dealings(b, co):
    nature = b.get("natureOfDir",""); txns = b.get("directorTransactions",[])
    if "refer attachment" in nature.lower() or "refer attachment" in (b.get("relInterestAccountName","") or "").lower(): return None
    nature = re.sub(r'\s*Directors?\s*$','',nature).strip()
    if not nature: nature = "Director"
    parts = [f"{co}: Dealings by {nature}."]
    by_type = defaultdict(lambda:{"qty":0,"tv":0,"prices":[],"dates":[]})
    for tx in txns:
        tt=tx.get("transType","Transaction"); q=tx.get("quantity",0) or 0; pr=tx.get("price",0) or 0; td=tx.get("transactionDate","")
        by_type[tt]["qty"]+=q; by_type[tt]["tv"]+=q*pr
        if pr and pr not in by_type[tt]["prices"]: by_type[tt]["prices"].append(pr)
        if td and td not in by_type[tt]["dates"]: by_type[tt]["dates"].append(td)
    for tt,d in by_type.items():
        q=d["qty"]; ps=d["prices"]; ds=d["dates"]
        qs = f"{int(q):,}" if q==int(q) else f"{q:,.2f}"
        if len(ps)==1: prs=f"at Rs. {ps[0]}"
        elif ps: prs=f"at avg Rs. {d['tv']/q:,.2f}" if q else ""
        else: prs=""
        dts = ds[0] if len(ds)==1 else f"{ds[0]}-{ds[-1]}" if ds else ""
        parts.append(f"{tt}: {qs} shares {prs} on {dts}.".strip())
    return " ".join(parts)

def _build_appointment(b, co):
    dirs = b.get("dirList",[])
    if not dirs: return None
    parts = [f"{co}:"]
    for d in dirs:
        n = d.get("natureOfDir","Director")
        if n and "director" not in n.lower() and "chairperson" not in n.lower(): n += " Director"
        parts.append(f"Appointed {n}, effective {d.get('effectiveDate','')}.")
        if d.get("numberOfShares",0): parts.append(f"Holds {int(d['numberOfShares']):,} shares.")
    return " ".join(parts)

def _build_resignation(b, co, dtype):
    role = "Chairperson" if "Chp" in dtype else "Director"
    m = re.search(r'w\.?e\.?f\.?\s*(\d{1,2}[./]\d{1,2}[./]\d{4})', b.get("remarks","") or "")
    return f"{co}: Resignation of {role}{f', effective {m.group(1)}' if m else ''}."

def _build_chairperson(b, co):
    m = re.search(r'(?:w\.?e\.?f\.?|effective|from)\s*[:\s]*(\d{1,2}[./]\d{1,2}[./]\d{4}|\d{1,2}\s+\w+\s+\d{4})', b.get("remarks","") or "", re.IGNORECASE)
    return f"{co}: Appointment of Chairperson{f', effective {m.group(1)}' if m else ''}."

def _build_rights(b, co):
    parts = [f"{co} — Rights Issue."]
    if b.get("numOfVotingShrsIssued"): parts.append(f"{int(b['numOfVotingShrsIssued']):,} voting shares.")
    if b.get("votingShareConsideration"): parts.append(f"At Rs. {b['votingShareConsideration']}/- per share.")
    if b.get("xr"): parts.append(f"XR: {b['xr']}.")
    if b.get("remarks"): parts.append(b["remarks"])
    return " ".join(parts)

def _build_egm(b, co):
    parts = [f"{co} — EGM scheduled for {b.get('dateOfEgm','')}"]
    if b.get("time"): parts.append(f"at {b['time']}")
    if b.get("venue"): parts.append(f"at {b['venue']}")
    s = " ".join(parts).strip() + "."
    if b.get("resToBePassed"):
        res_clean = re.sub(r'\s+',' ',b['resToBePassed']).strip()
        s += f" Resolutions: {res_clean}."
    return s

def _build_agm(b, co):
    agm = b.get("agm","") or b.get("dateOfAgm","")
    s = f"{co} — AGM" + (f" scheduled for {agm}" if agm else "") + "."
    if b.get("remarks"): s += f" {b['remarks']}"
    return s

# ─── AI SUMMARIZATION ────────────────────────────────────────────────
JUNK = ['dear madam','dear sir','yours faithfully','yours sincerely','chief regulatory officer',
    'west block','echelon square','world trade centre','p w corporate','heed oltrce','tel:','fax:']

def pre_clean(t):
    if not t: return ""
    for p in [r'Yours\s+(faithfully|sincerely|truly).*',r'BY\s+ORDER\s+OF\s+THE\s+BOARD.*',r'For\s+and\s+on\s+behalf\s+of.*']:
        t = re.sub(p,'',t,flags=re.IGNORECASE|re.DOTALL)
    for p in [r'Dear\s+(Sir|Madam|Madan-?r?|Sirs?)[\s,]*',r'Ms\.?\s+Nilupa\s+Perar?a.{0,100}',r'Mrs\.?\s+Nilupa\s+Perar?a.{0,100}',
        r'Chief\s+Regulatory\s+Officer.{0,100}',r'Colombo\s+Stock\s+Exc[a-z]*.{0,100}',r'Echelon\s+Square.{0,60}',
        r'World\s*\'?[Tt]rade\s+Centr?e.{0,60}',r'West\s+Block.{0,60}',
        r'#?\d+[-/]?\d*,?\s*\w+\s+(Road|Street|Lane|Mawatha|Place).{0,80}',r'Colombo\s*\d{1,2}.{0,40}',r'Sri\s+Lanka\.?',
        r'Tel(?:ephone)?:?\s*[\+\d\s\-\(\)\']{5,30}',r'Fax:?\s*[\+\d\s\-\(\)\']{5,30}',r'E-?mail:?\s*\S+@\S+',
        r'P\s*W\s*(?:Corporate|Gorporate)\s*Secretarial.{0,80}',r'M&S\s*Managers\s*&\s*Secretaries.{0,80}',
        r'JACEY\s*&?\s*(?:COMPANY|GOMPANY).{0,80}',r'JULIUS\s*&?\s*CREASY.{0,80}']:
        t = re.sub(p,' ',t,flags=re.IGNORECASE)
    return re.sub(r'\s+',' ',re.sub(r'[!.]{2,}','.',re.sub(r'[{}\[\]|\\@#$^~`]','',t))).strip()[:3000]

def is_good(s):
    if not s or len(s)<20 or len(s)>800 or len(s.split())<5: return False
    sl=s.lower()
    for p in JUNK:
        if p in sl: return False
    for l in ['key details were not provided','unfortunately','does not contain sufficient','not enough information',
              'cannot extract specific','the provided text does not','the given text','here are the specific facts','nilupa perera']:
        if l in sl: return False
    return True

def is_fallback(s):
    if not s: return True
    sl=s.lower()
    for b in ['i don\'t see','unfortunately','does not contain','not enough information','cannot extract',
              'the provided text','key details were not provided','here are the specific facts','nilupa perera','company registration number']:
        if b in sl: return True
    return False

def ai_summarize(raw, company="", subject="", log=print):
    if not raw or len(raw)<30: return None
    cleaned = pre_clean(raw)
    if len(cleaned)<30: cleaned = re.sub(r'\s+',' ',raw).strip()[:2000]
    if len(cleaned)<30: return None
    sys_msg = ("You extract key facts from CSE corporate disclosures. Write 2-3 sentences with SPECIFIC details. "
        "Include: quantities, rupee amounts, percentages, dates, positions. NEVER include person names. "
        "Focus ONLY on what the company announces. NEVER start with 'Here are the facts'.")
    usr = f"Company: {company}\nCategory: {subject}\n\nExtract specific facts:\n{cleaned[:2000]}"
    if GROQ_API_KEY:
        for att in range(3):
            try:
                r = requests.post("https://api.groq.com/openai/v1/chat/completions",
                    headers={"Content-Type":"application/json","Authorization":f"Bearer {GROQ_API_KEY}"},
                    json={"model":"llama-3.3-70b-versatile","messages":[{"role":"system","content":sys_msg},{"role":"user","content":usr}],
                          "max_tokens":200,"temperature":0.1},timeout=25)
                if r.status_code==200:
                    ch=r.json().get("choices",[])
                    if ch:
                        s=re.sub(r'\n+',' ',re.sub(r'\*+','',ch[0].get("message",{}).get("content",""))).strip()
                        for px in [r'^here are the specific facts[^:]*:\s*',r'^summary:\s*']:
                            s=re.sub(px,'',s,flags=re.IGNORECASE).strip()
                        if is_good(s): return s
                        elif att<2: time.sleep(3); continue
                        elif s and len(s)>30: return s
                elif r.status_code==429: log("      Groq rate limited"); time.sleep(20)
                else: break
            except Exception as e: log(f"      Groq error: {e}"); break
    if GEMINI_API_KEY:
        try:
            r=requests.post(f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-lite:generateContent?key={GEMINI_API_KEY}",
                headers={"Content-Type":"application/json"},
                json={"contents":[{"parts":[{"text":f"{sys_msg}\n\n{usr}"}]}],"generationConfig":{"maxOutputTokens":200,"temperature":0.1}},timeout=25)
            if r.status_code==200:
                c=r.json().get("candidates",[])
                if c:
                    p=c[0].get("content",{}).get("parts",[])
                    if p:
                        s=re.sub(r'\*+','',p[0].get("text","")).strip()
                        if is_good(s): return s
        except Exception as e: log(f"      Gemini error: {e}")
    return None

# ─── GOOGLE SHEETS ────────────────────────────────────────────────────
class GoogleManager:
    def __init__(self, log_callback=None):
        self.log = log_callback or print
        import gspread; from google.oauth2.service_account import Credentials
        scopes=["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive.readonly"]
        sa_json = os.environ.get("SERVICE_ACCOUNT_KEY")
        if sa_json:
            import tempfile
            with tempfile.NamedTemporaryFile(mode='w',suffix='.json',delete=False) as f: f.write(sa_json); sa_path=f.name
            creds=Credentials.from_service_account_file(sa_path,scopes=scopes); os.unlink(sa_path)
            self.log("✓ Auth via env var")
        else:
            creds=Credentials.from_service_account_file(SERVICE_ACCOUNT_JSON,scopes=scopes)
            self.log("✓ Auth via file")
        self.gc=gspread.authorize(creds); self.service_account_email=creds.service_account_email
        self.spreadsheet=self.gc.open(SPREADSHEET_NAME); self.worksheet=self.spreadsheet.sheet1
        self.log(f"✓ Opened: {SPREADSHEET_NAME}")
        hdrs=["Date","Time","Company","Subject","Description","AI Summary","PDF Link","PDF Count","Unique Key"]
        fr=self.worksheet.row_values(1)
        if not fr or not any(fr): self.worksheet.update('A1:I1',[hdrs]); self.log("  ✓ Headers written")
        else: self.log(f"  {len(self.worksheet.get_all_values())-1} existing rows")
    def get_existing_keys(self):
        try: return set(self.worksheet.col_values(9)[1:])
        except: return set()

# ─── CSE API ──────────────────────────────────────────────────────────
def fetch_announcements(log=print):
    r=requests.post(CSE_API+"approvedAnnouncement",headers={**HTTP_HEADERS,'Content-Type':'application/x-www-form-urlencoded'},timeout=20)
    if r.status_code!=200: log(f"✗ API {r.status_code}"); return []
    items=r.json().get("approvedAnnouncements",[]); log(f"✓ {len(items)} announcements"); return items[:MAX_DISCLOSURES]

def get_detail(ann_id):
    try:
        r=requests.post(CSE_API+"getAnnouncementById",data={"announcementId":ann_id},headers={**HTTP_HEADERS,'Content-Type':'application/x-www-form-urlencoded'},timeout=15)
        if r.status_code==200 and len(r.content)>5: return r.json()
    except: pass
    try:
        r=requests.post(CSE_API+"getGeneralAnnouncementById",data={"announcementId":ann_id},headers={**HTTP_HEADERS,'Content-Type':'application/x-www-form-urlencoded'},timeout=15)
        if r.status_code==200 and r.text!="{}": return r.json()
    except: pass
    return None

def download_pdf_text(file_url, log=print):
    url=(CSE_CDN+file_url).replace(' ','%20')
    try:
        r=requests.get(url,headers={'User-Agent':HTTP_HEADERS['User-Agent'],'Referer':'https://www.cse.lk/'},timeout=20)
        if r.status_code!=200 or r.content[:5]!=b'%PDF-': return None
        log(f"      PDF ({len(r.content):,} bytes)")
        reader=PyPDF2.PdfReader(io.BytesIO(r.content))
        text="".join((p.extract_text() or "")+"\n" for p in reader.pages).strip()
        if len(text)>30: return text
        if HAS_OCR:
            try:
                imgs=convert_from_bytes(r.content,dpi=200)
                ocr="".join(pytesseract.image_to_string(i)+"\n" for i in imgs[:3]).strip()
                if len(ocr)>30: return ocr
            except: pass
    except: pass
    return None

# ─── CORE ENGINE ──────────────────────────────────────────────────────
def process_one_item(gm, item, existing_keys, log=print):
    ann_id=item.get("announcementId"); co=item.get("company",""); cat=item.get("announcementCategory","")
    ds=item.get("dateOfAnnouncement",""); rem=item.get("remarks") or ""; cr=item.get("createdDate",0)
    try: dt=datetime.fromtimestamp(cr/1000); ts=dt.strftime("%I:%M:%S %p"); ds=ds or dt.strftime("%d %b %Y").upper()
    except: ts=""
    ukey=f"{ds}|{ts}|{co}"
    if ukey in existing_keys or any(k.startswith(ukey) for k in existing_keys): return 0
    log(f"  {co[:40]} — {cat}")
    detail=get_detail(ann_id); pdfs=[]
    if detail:
        for doc in detail.get("reqAnnouncementDocs",[]):
            fu=doc.get("fileUrl",""); bu=doc.get("baseUrl",CSE_CDN)
            if fu: pdfs.append((bu or CSE_CDN)+fu)
    log(f"    {len(pdfs)} PDF(s)")
    # Summary
    summary=None
    if detail:
        summary=build_structured_summary(detail)
        if summary: log(f"    ✓ Structured")
    if not summary and pdfs:
        pt=download_pdf_text(pdfs[0].replace(CSE_CDN,"").replace("https://cdn.cse.lk/",""),log=log)
        if pt:
            summary=ai_summarize(pt,co,cat,log=log)
            if summary: log(f"    ✓ AI summary")
    if not summary and detail and not pdfs:
        # No PDF — try AI on description/remarks text
        base=detail.get("reqBaseAnnouncement",{})
        desc_text=(base.get("description","") or "") + " " + (base.get("remarks","") or "")
        desc_text=desc_text.strip()
        if desc_text and len(desc_text)>30:
            summary=ai_summarize(desc_text,co,cat,log=log)
            if summary: log(f"    ✓ AI summary (from description)")
    if not summary and detail:
        # Last resort: use raw description/remarks from API
        base=detail.get("reqBaseAnnouncement",{})
        desc_text=(base.get("description","") or "") + " " + (base.get("remarks","") or "")
        desc_text=re.sub(r'\s+',' ',desc_text).strip()
        if desc_text and len(desc_text)>20:
            if len(desc_text)>500: desc_text=desc_text[:497]+"..."
            summary=f"{co}: {desc_text}"
            log(f"    ✓ Description fallback")
    if not summary: summary=f"{co} — {cat}."; log(f"    ⚠ Fallback")
    # Write
    w=0
    try:
        if len(pdfs)>1:
            for pi,link in enumerate(pdfs):
                key=f"{ukey}|PDF{pi+1}"
                gm.worksheet.append_row([ds,ts,co,cat,f"PDF {pi+1} of {len(pdfs)}",summary,link,len(pdfs),key],value_input_option='RAW')
                existing_keys.add(key); w+=1
        else:
            gm.worksheet.append_row([ds,ts,co,cat,rem[:200],summary,pdfs[0] if pdfs else "",len(pdfs),ukey],value_input_option='RAW')
            existing_keys.add(ukey); w=1
        log(f"    ✓ Written")
    except Exception as e: log(f"    ✗ Write error: {e}")
    return w

def fix_old_summaries(gm, items, log=print, running_check=None):
    log("━"*50); log("CHECKING OLD SUMMARIES..."); log("━"*50)
    all_rows=gm.worksheet.get_all_values()
    needs=[(i,r) for i,r in enumerate(all_rows[1:],start=2) if not(r[5] if len(r)>5 else "") or is_fallback(r[5] if len(r)>5 else "")]
    if not needs: log("  All OK"); return
    log(f"  {len(needs)} need fixing")
    for idx,(rn,row) in enumerate(needs):
        if running_check and not running_check(): break
        co=row[2] if len(row)>2 else ""; subj=row[3] if len(row)>3 else ""; pl=row[6] if len(row)>6 else ""
        log(f"  [{idx+1}/{len(needs)}] {co[:40]} — {subj}")
        s=None
        for it in items:
            if it.get("company","").strip()==co.strip() and it.get("announcementCategory","").strip()==subj.strip():
                d=get_detail(it.get("announcementId"))
                if d: s=build_structured_summary(d)
                if s: log(f"    ✓ Structured")
                break
        if not s and pl:
            url=None; m=re.search(r'HYPERLINK\("([^"]+)"',pl)
            if m: url=m.group(1)
            elif pl.startswith("http"): url=pl
            if url:
                pt=download_pdf_text(url.replace(CSE_CDN,"").replace("https://cdn.cse.lk/",""),log=log)
                if pt: s=ai_summarize(pt,co,subj,log=log)
                if s: log(f"    ✓ AI")
        if not s and not pl:
            # No PDF — try AI on description from API
            for it in items:
                if it.get("company","").strip()==co.strip() and it.get("announcementCategory","").strip()==subj.strip():
                    d=get_detail(it.get("announcementId"))
                    if d:
                        base=d.get("reqBaseAnnouncement",{})
                        desc_text=(base.get("description","") or "") + " " + (base.get("remarks","") or "")
                        desc_text=desc_text.strip()
                        if desc_text and len(desc_text)>30:
                            s=ai_summarize(desc_text,co,subj,log=log)
                            if s: log(f"    ✓ AI (from description)")
                    break
        if s:
            try: gm.worksheet.update_cell(rn,6,s)
            except Exception as e: log(f"    ✗ {e}")
        else: log(f"    ~ Skip")
        if idx<len(needs)-1: time.sleep(SUMMARY_DELAY)

# ═══════════════════════════════════════════════════════════════════════
#  HEADLESS
# ═══════════════════════════════════════════════════════════════════════
def run_headless():
    start=time.time(); hl=lambda m:print(f"[{datetime.now().strftime('%H:%M:%S')}] {m}")
    hl("="*50); hl("CSE TRACKER v9 — Headless"); hl("="*50)
    gm=GoogleManager(log_callback=hl); ek=gm.get_existing_keys(); hl(f"  Existing: {len(ek)}")
    items=fetch_announcements(log=hl)
    if not items: hl("No announcements"); return
    new=[]
    for it in items:
        co=it.get("company",""); ds=it.get("dateOfAnnouncement",""); cr=it.get("createdDate",0)
        try: dt=datetime.fromtimestamp(cr/1000); ts=dt.strftime("%I:%M:%S %p")
        except: ts=""
        uk=f"{ds}|{ts}|{co}"
        if uk not in ek and not any(k.startswith(uk) for k in ek): new.append(it)
    hl(f"  {len(new)} NEW")
    if new:
        hl("─"*40); hl("PROCESSING ROW BY ROW"); hl("─"*40)
        tot=0
        for i,it in enumerate(new):
            hl(f"\n[{i+1}/{len(new)}]")
            try: tot+=process_one_item(gm,it,ek,log=hl)
            except Exception as e: hl(f"    ✗ {e}")
            if i<len(new)-1: time.sleep(SUMMARY_DELAY)
        hl(f"\n  Total: {tot} rows")
    else: hl("  Up to date!")
    fix_old_summaries(gm,items,log=hl)
    hl("="*50); hl(f"✓ DONE in {time.time()-start:.0f}s"); hl("="*50)

# ═══════════════════════════════════════════════════════════════════════
#  GUI
# ═══════════════════════════════════════════════════════════════════════
def run_gui():
    import tkinter as tk
    from tkinter import messagebox, scrolledtext
    class App:
        def __init__(self, root):
            self.root=root; root.title("CSE Disclosure Tracker v9"); root.geometry("820x700"); root.configure(bg="#f0f0f0")
            self.running=False; self.gm=None; self._ui(); self._check()
        def _ui(self):
            tk.Label(self.root,text="📊 CSE Disclosure Tracker v9",font=("Helvetica",18,"bold"),bg="#f0f0f0").pack(pady=(10,0))
            tk.Label(self.root,text="Pure API — No Chrome needed!",font=("Helvetica",10),bg="#f0f0f0",fg="#666").pack()
            bf=tk.Frame(self.root,bg="#f0f0f0"); bf.pack(pady=10,fill=tk.X,padx=20)
            self.b1=tk.Button(bf,text="▶  Run Full",command=lambda:self._go(self._full),font=("Helvetica",13,"bold"),bg="#2196F3",fg="white",width=15,height=2); self.b1.pack(side=tk.LEFT,padx=5)
            self.b2=tk.Button(bf,text="🤖  Fix Summaries",command=lambda:self._go(self._fix),font=("Helvetica",13),bg="#4CAF50",fg="white",width=15,height=2); self.b2.pack(side=tk.LEFT,padx=5)
            self.b3=tk.Button(bf,text="📋  Open Sheet",command=self._sheet,font=("Helvetica",13),bg="#FF9800",fg="white",width=12,height=2); self.b3.pack(side=tk.LEFT,padx=5)
            self.b4=tk.Button(bf,text="⏹  Stop",command=self._stop,font=("Helvetica",13),bg="#f44336",fg="white",width=8,height=2,state=tk.DISABLED); self.b4.pack(side=tk.LEFT,padx=5)
            self.con=scrolledtext.ScrolledText(self.root,wrap=tk.WORD,font=("Courier",11),bg="#1e1e1e",fg="#00ff00",insertbackground="#00ff00",height=30)
            self.con.pack(padx=10,pady=5,fill=tk.BOTH,expand=True)
            self.st=tk.Label(self.root,text="Ready",font=("Helvetica",10),bg="#f0f0f0",fg="#666"); self.st.pack(pady=(0,5))
        def log(self,m): self.root.after(0,lambda:(self.con.insert(tk.END,m+"\n"),self.con.see(tk.END)))
        def _ss(self,t): self.root.after(0,lambda:self.st.config(text=t))
        def _sb(self,r):
            s=tk.DISABLED if r else tk.NORMAL
            self.root.after(0,lambda:(self.b1.config(state=s),self.b2.config(state=s),self.b4.config(state=tk.NORMAL if r else tk.DISABLED)))
        def _go(self,fn):
            if self.running: return
            self.running=True; self._sb(True); self.con.delete("1.0",tk.END); threading.Thread(target=fn,daemon=True).start()
        def _stop(self): self.running=False; self.log("\n⏹ Stopping..."); self._ss("Stopped"); self._sb(False)
        def _sheet(self):
            if self.gm: webbrowser.open(self.gm.spreadsheet.url)
            else: messagebox.showinfo("Sheet","Run tracker first.")
        def _conn(self):
            if self.gm: return True
            try:
                self.log("━"*50); self.log("CONNECTING..."); self.log("━"*50)
                self.gm=GoogleManager(log_callback=self.log); self.log(f"  URL: {self.gm.spreadsheet.url}"); return True
            except FileNotFoundError: self.log(f"\n✗ {SERVICE_ACCOUNT_JSON} not found!"); self.log(SETUP_GUIDE); return False
            except Exception as e: self.log(f"\n✗ {e}"); return False
        def _check(self):
            self.log("✓ Ready!")
            self.log(f"  Google: {'✓' if os.path.exists(SERVICE_ACCOUNT_JSON) else '✗'} service_account.json")
            self.log(f"  AI: Groq {'✓' if GROQ_API_KEY else '✗'} | Gemini {'✓' if GEMINI_API_KEY else '✗'}")
            self.log(f"  Chrome: NOT NEEDED ✓")
        def _full(self):
            try:
                if not self._conn(): self._sb(False); self.running=False; return
                ek=self.gm.get_existing_keys(); self.log(f"  Existing: {len(ek)}")
                self.log("━"*50); self.log("FETCHING..."); self.log("━"*50); self._ss("Fetching...")
                items=fetch_announcements(log=self.log)
                if not items: self.log("✗ None"); self._sb(False); self.running=False; return
                new=[]
                for it in items:
                    co=it.get("company",""); ds=it.get("dateOfAnnouncement",""); cr=it.get("createdDate",0)
                    try: dt=datetime.fromtimestamp(cr/1000); ts=dt.strftime("%I:%M:%S %p")
                    except: ts=""
                    uk=f"{ds}|{ts}|{co}"
                    if uk not in ek and not any(k.startswith(uk) for k in ek): new.append(it)
                self.log(f"  {len(new)} NEW")
                if not new: self.log("  Up to date!")
                else:
                    self.log("━"*50); self.log("PROCESSING ROW BY ROW..."); self.log("━"*50)
                    tot=0
                    for i,it in enumerate(new):
                        if not self.running: break
                        self._ss(f"Processing {i+1}/{len(new)}..."); self.log(f"\n[{i+1}/{len(new)}]")
                        tot+=process_one_item(self.gm,it,ek,log=self.log)
                        if i<len(new)-1: time.sleep(SUMMARY_DELAY)
                    self.log(f"\n  Total: {tot} rows")
                if self.running: fix_old_summaries(self.gm,items,log=self.log,running_check=lambda:self.running)
                self.log("\n"+"═"*50); self.log("✓ ALL DONE!"); self.log("═"*50); self._ss("Complete!")
            except Exception as e: self.log(f"\n✗ {e}"); import traceback; self.log(traceback.format_exc())
            finally: self.running=False; self._sb(False)
        def _fix(self):
            try:
                if not self._conn(): self._sb(False); self.running=False; return
                items=fetch_announcements(log=self.log)
                fix_old_summaries(self.gm,items,log=self.log,running_check=lambda:self.running)
                self.log("\n✓ Done!"); self._ss("Complete!")
            except Exception as e: self.log(f"\n✗ {e}")
            finally: self.running=False; self._sb(False)
    root=tk.Tk(); App(root); root.mainloop()

# ═══════════════════════════════════════════════════════════════════════
if __name__=="__main__":
    headless = "--headless" in sys.argv or os.environ.get("SERVICE_ACCOUNT_KEY")
    if not headless:
        try: import tkinter; tkinter.Tk().destroy()
        except: headless=True
    if headless: run_headless()
    else: run_gui()
