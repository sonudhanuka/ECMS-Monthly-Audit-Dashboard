"""
build_dashboard.py
══════════════════════════════════════════════════════════════════════
Takes your original HTML dashboard file and injects fresh data from
SharePoint into it. Run by GitHub Actions on a schedule.

MONTHLY WORKFLOW:
  1. Drop new Excel audit file into your SharePoint folder
  2. GitHub Actions runs nightly (or click "Run workflow" manually)
  3. Dashboard updates automatically at your GitHub Pages URL

SETUP:
  1. Rename your HTML file to:  ECMS_Monthly_Audit_Dashboard.html
  2. Add SHAREPOINT_FOLDER_URL as a GitHub Actions secret
  3. Enable GitHub Pages from /docs folder
══════════════════════════════════════════════════════════════════════
"""

import os, io, re, json, base64, requests, warnings
from datetime import datetime
import pandas as pd

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════════════
SHAREPOINT_FOLDER_URL = os.environ.get("SHAREPOINT_FOLDER_URL", "")
TEMPLATE_FILE         = "ECMS_Monthly_Audit_Dashboard.html"
OUTPUT_FILE           = "docs/index.html"

# Column headers in your Excel file (change only if different)
COL_WEEK      = "Week"
COL_TICKET    = "Ticket Number"
COL_PARAMETER = "Audit Parameter"
COL_EVAL      = "Evaluation"
COL_TOWER     = "Tower"
COL_TEAM      = "Teams"
COL_AGENT     = "Agent"
COL_COMMENTS  = "Comments"
# ══════════════════════════════════════════════════════════════════════


def log(msg): print(f"  {msg}")


# ── SharePoint downloader ──────────────────────────────────────────────

def _fetch(url, timeout=25):
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=timeout)
        r.raise_for_status()
        return r.content
    except Exception as e:
        log(f"Fetch failed: {e}"); return None

def _is_excel(b): return b and len(b) > 4 and b[:2] == b'PK'

def _read_excel(name, content):
    try:
        xl = pd.ExcelFile(io.BytesIO(content))
        for sheet in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name=sheet)
            df.columns = [str(c).strip() for c in df.columns]
            df = df.dropna(how="all").reset_index(drop=True)
            if len(df) > 0:
                df["_src"] = name
                log(f"Read '{name}' sheet '{sheet}' — {len(df)} rows")
                return df
    except Exception as e:
        log(f"Cannot read {name}: {e}")
    return None

def fetch_from_sharepoint(url):
    if not url.strip():
        log("No SharePoint URL — using sample data"); return []
    log("Connecting to SharePoint...")
    results = []

    # Graph API anonymous share resolution
    try:
        enc = base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
        r = requests.get(
            f"https://graph.microsoft.com/v1.0/shares/u!{enc}/driveItem/children",
            headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if r.status_code == 200:
            for item in r.json().get("value", []):
                name = item.get("name", "")
                if not name.lower().endswith((".xlsx", ".xls")): continue
                dl = item.get("@microsoft.graph.downloadUrl") or item.get("downloadUrl", "")
                if dl:
                    content = _fetch(dl)
                    if _is_excel(content):
                        results.append((name, content))
                        log(f"Downloaded: {name} ({len(content)//1024}KB)")
        else:
            log(f"Graph API: {r.status_code} — trying direct link")
    except Exception as e:
        log(f"Graph API error: {e}")

    # Fallback: treat URL as direct file download
    if not results:
        sep = "&" if "?" in url else "?"
        dl = url + sep + "download=1"
        content = _fetch(dl)
        if _is_excel(content):
            fname = url.split("/")[-1].split("?")[0] or "audit.xlsx"
            results.append((fname, content))
            log(f"Direct download: {fname}")

    return results


# ── Data loader ────────────────────────────────────────────────────────

def load_data(url):
    files = fetch_from_sharepoint(url)
    dfs, names = [], []
    for name, content in files:
        df = _read_excel(name, content)
        if df is not None:
            dfs.append(df); names.append(name)
    if dfs:
        merged = pd.concat(dfs, ignore_index=True)
        log(f"Merged: {len(merged)} rows from {len(names)} file(s)")
        return merged, names, True
    log("Using built-in sample data")
    return _sample_data(), ["sample_data_Apr2026"], False


# ── Sample data ────────────────────────────────────────────────────────

def _sample_data():
    agents = {
        "Aarti Lasure":("Workplace",117,0,12),"Abdul Kadar Aiyaz Mohiuddin":("Workplace",45,0,5),
        "Abhinav Pandey":("Network",27,0,3),"Abhishek B Biju Raj":("Cloud",17,1,2),
        "Abigyat Snehi":("Security",8,1,1),"Afnan Khalid":("Service Desk",72,0,8),
        "Akanksha Kharche":("Cloud",9,0,1),"Akshay Ramteke":("Service Desk",108,0,12),
        "Alok Kumar":("Workplace",144,0,16),"Aman Tiwari":("Service Desk",9,0,1),
        "Anchal Tiwari":("Cloud",18,0,2),"Ankita Uttekar":("Service Desk",162,0,18),
        "Apeksha Jadhav":("Cloud",8,1,1),"Archana Rajeshram Kahar":("Workplace",81,0,8),
        "Ashit Gangwar":("Workplace",90,0,10),"Bhagyesh Mandole":("Workplace",18,0,2),
        "Chandralekha Shesanala":("Service Desk",17,1,2),"Gaurav Shinde":("Network",27,0,3),
        "Habib Shaikh":("Service Desk",63,0,7),"Haripriya Dariboyina":("Workplace",117,0,13),
        "Junaid Ahmed":("Service Desk",9,0,1),"Komal Rao":("Workplace",189,0,20),
        "Lavanya Venkatarathnaiah":("Network",9,0,1),"Manik Borvadkar":("Workplace",54,0,5),
        "Manish Sonekar":("Workplace",108,0,11),"Niki Singh":("Workplace",9,0,1),
        "Nikita A. Kapgate":("Workplace",36,0,4),"Nikita Chari":("Workplace",18,0,2),
        "Paramraj Singh":("Network",18,0,2),"Pavan Kalyan s":("Service Desk",36,0,4),
        "Prakarsh Pandey":("Security",18,0,2),"Prakhar Saini":("Security",16,2,2),
        "Priyanka Anand":("Security",27,0,3),"Rahul Francis":("Workplace",9,0,1),
        "Rishabh Shrivastava":("Workplace",72,0,8),"Ruchi Vinod Bopche":("Service Desk",137,2,16),
        "SURAJ MOHANDAS":("Workplace",71,1,8),"Sabhyatha Suvarna":("Service Desk",144,0,15),
        "Sakshi Jagtap":("Service Desk",45,0,5),"Sarthak Shah":("Workplace",198,0,19),
        "Shahabaz Zainuddin Bagwan":("Cloud",18,0,2),"Shajid Jamil":("Cloud",17,1,2),
        "Sharukh Razeen A":("Network",47,3,5),"Shivani Nimse":("Cloud",18,0,2),
        "Shubham Patil":("Network",18,0,2),"Shweta R. Hebbal":("Network",8,1,1),
        "Sneha Zunake":("Cloud",36,0,4),"Sreerag Narayanan":("Workplace",35,1,4),
        "Sudarshan Jadhav":("Network",18,0,2),"Suman Aradhya A R":("Service Desk",45,0,5),
        "Surendra Paruchuri":("Workplace",9,0,1),"Vaishnavi Agawane":("Workplace",81,0,9),
        "Vaishnavi Chaudhari":("Workplace",108,0,12),"Vaishnavi Pawar":("Network",22,0,3),
        "Writojaya Dey":("Service Desk",8,1,1),"mithraa Supramanian":("Service Desk",149,0,17),
    }
    nm = [
        ("CW14","INC56052293","Closure Information","Not Met","IMS","Security","Prakhar Saini","New closure template not followed"),
        ("CW14","INC56166531","Closure Information","Not Met","IMS","Security","Abigyat Snehi","New closure template not followed"),
        ("CW14","INC56204600","Closure Information","Not Met","IMS","Cloud","Apeksha Jadhav","The ticket was closed using the 3-strike process, but the close code does not reflect this."),
        ("CW14","INC56206932","Closure Information","Not Met","IMS","Service Desk","Writojaya Dey","New closure template not followed"),
        ("CW14","INC56254753","Closure Information","Not Met","IMS","Service Desk","Chandralekha Shesanala","Close code is null"),
        ("CW15","INC56373053","Closure Information","Not Met","IMS","Cloud","Abhishek B Biju Raj","One point is missing in the closure template"),
        ("CW15","INC56149475","Closure Information","Not Met","IMS","Network","Sharukh Razeen A","New closure template not followed"),
        ("CW15","INC56270010","Closure Information","Not Met","IMS","Network","Sharukh Razeen A","Root Cause Analysis and Preventive Measures: template not followed"),
        ("CW15","INC56149475","Closure Information","Not Met","IMS","Network","Sharukh Razeen A","Root Cause Analysis and Preventive Measures step is not mentioned in the closure template"),
        ("CW15","INC56253551","Closure Information","Not Met","IMS","Security","Prakhar Saini","Closure template is not followed"),
        ("CW15","INC56218121","Closure Information","Not Met","IMS","Workplace","Sreerag Narayanan","New closure template not followed"),
        ("CW16","INC56405380","Closure Information","Not Met","IMS","Network","Shweta R. Hebbal","Closure template is not attached in closure information"),
        ("CW16","INC56407764","Closure Information","Not Met","IMS","Workplace","SURAJ MOHANDAS","In closure template - Root Cause Analysis and Preventive Measures cannot be NA"),
        ("CW16","INC56522362","Closure Information","Not Met","IMS","Service Desk","Ruchi Vinod Bopche","New closure template not followed"),
        ("CW17","INC56572172","Closure Information","Not Met","IMS","Cloud","Shajid Jamil","Closure template not followed."),
        ("CW17","INC56497608","Closure Information","Not Met","IMS","Service Desk","Ruchi Vinod Bopche","Closure template not used"),
    ]
    weeks = ["CW14","CW15","CW16","CW17"]
    rows = []
    for name,(team,met,nmc,tickets) in agents.items():
        for t in range(tickets):
            for _ in range(9):
                rows.append({COL_WEEK:weeks[t%4],COL_TICKET:f"TKT_{name[:4]}_{t}",
                    COL_PARAMETER:"Closure Information",COL_EVAL:"Met",
                    COL_TOWER:"IMS",COL_TEAM:team,COL_AGENT:name,COL_COMMENTS:"","_src":"sample"})
    nm_rows = [{COL_WEEK:w,COL_TICKET:t,COL_PARAMETER:p,COL_EVAL:e,
                COL_TOWER:tw,COL_TEAM:tm,COL_AGENT:a,COL_COMMENTS:c,"_src":"sample"}
               for w,t,p,e,tw,tm,a,c in nm]
    return pd.concat([pd.DataFrame(rows),pd.DataFrame(nm_rows)],ignore_index=True)


# ── Stats computer ─────────────────────────────────────────────────────

def _fc(df, *keys):
    for k in keys:
        m = next((c for c in df.columns if k.lower() in c.lower()), None)
        if m: return m
    return None

def compute_stats(df):
    wc  = _fc(df,COL_WEEK,"week","cw")        or COL_WEEK
    tkc = _fc(df,COL_TICKET,"ticket","inc")    or COL_TICKET
    pc  = _fc(df,COL_PARAMETER,"param","audit")or COL_PARAMETER
    ec  = _fc(df,COL_EVAL,"evaluation","eval") or COL_EVAL
    tc  = _fc(df,COL_TEAM,"team","teams")      or COL_TEAM
    ac  = _fc(df,COL_AGENT,"agent","engineer") or COL_AGENT
    cc  = _fc(df,COL_COMMENTS,"comment")       or COL_COMMENTS

    df = df.copy()
    df["_met"] = (df[ec].astype(str).str.strip().str.lower()
                  .isin(["met","yes","pass","1","true"])
                  if ec in df.columns else True)
    df["_nm"]  = ~df["_met"]

    wnum = lambda w: int(re.search(r'(\d+)',str(w)).group(1)) if re.search(r'(\d+)',str(w)) else 999

    # AGENT_STATS: {"Name": {"met":N,"notMet":N,"team":"X","tickets":N}}
    agent_stats = {}
    if ac in df.columns:
        for agent, grp in df.groupby(ac):
            agent_stats[str(agent)] = {
                "met":    int(grp["_met"].sum()),
                "notMet": int(grp["_nm"].sum()),
                "team":   str(grp[tc].iloc[0]) if tc in grp.columns else "Unknown",
                "tickets":int(grp[tkc].nunique()) if tkc in grp.columns else 0,
            }

    # WEEKLY: {"CW14": {"met":N,"notMet":N,"total":N,"tickets":N}}
    weekly = {}
    if wc in df.columns:
        for week, grp in df.groupby(wc):
            m = int(grp["_met"].sum()); n = int(grp["_nm"].sum())
            weekly[str(week)] = {"met":m,"notMet":n,"total":m+n,
                "tickets":int(grp[tkc].nunique()) if tkc in grp.columns else 0}
        weekly = dict(sorted(weekly.items(), key=lambda x: wnum(x[0])))

    # NM_RECORDS: [{"Week":..., "Ticket Number":..., ...}]
    nm_records = []
    for _, row in df[df["_nm"]].iterrows():
        nm_records.append({
            "Week":            str(row.get(wc,"")  if wc  in row.index else ""),
            "Ticket Number":   str(row.get(tkc,"") if tkc in row.index else ""),
            "Audit Parameter": str(row.get(pc,"")  if pc  in row.index else ""),
            "Evaluation":      "Not Met",
            "Tower":           str(row.get("Tower","IMS")),
            "Teams":           str(row.get(tc,"")  if tc  in row.index else ""),
            "Agent":           str(row.get(ac,"")  if ac  in row.index else ""),
            "Comments":        str(row.get(cc,"")  if cc  in row.index else ""),
        })

    t_met    = int(df["_met"].sum())
    t_nm     = int(df["_nm"].sum())
    t_params = t_met + t_nm
    t_tix    = int(df[tkc].nunique()) if tkc in df.columns else 0
    weeks    = list(weekly.keys())
    w_range  = f"{weeks[0]} – {weeks[-1]}" if len(weeks)>1 else (weeks[0] if weeks else "—")

    return dict(agent_stats=agent_stats, weekly=weekly, nm_records=nm_records,
                t_met=t_met, t_nm=t_nm, t_params=t_params, t_tix=t_tix,
                w_range=w_range, weeks=weeks)


# ── Trend chips ────────────────────────────────────────────────────────

def build_trend_chips(weekly):
    chips = []; prev_score = None
    for week, d in weekly.items():
        tot   = d["met"] + d["notMet"]
        score = d["met"] / tot * 100 if tot else 0
        if prev_score is None:
            cls = ""; arrow = ""
        elif score > prev_score:
            cls = "up"; arrow = " ↑"
        elif score < prev_score:
            cls = "down"; arrow = " ↓"
        else:
            cls = ""; arrow = ""
        chips.append(f'<span class="trend-chip {cls}">{week} → {score:.2f}%{arrow}</span>')
        prev_score = score
    return "\n      ".join(chips)


# ── HTML injector ──────────────────────────────────────────────────────

def inject(template, stats, file_names, is_live):
    """Replace only the data sections in the original HTML. CSS/layout unchanged."""
    h = template
    s = stats
    now = datetime.utcnow().strftime("%d %b %Y %H:%M UTC")
    month = datetime.utcnow().strftime("%B %Y")

    # 1. Title
    h = re.sub(r'<title>.*?</title>',
        f'<title>ECMS Monthly Audit Dashboard — {s["w_range"]}</title>', h)

    # 2. Header subtitle
    h = re.sub(r'<p>IMS Tower · Quality Audit · .*?</p>',
        f'<p>IMS Tower · Quality Audit · {s["w_range"]} · {month}</p>', h)

    # 3. Four totals constant
    h = re.sub(
        r'const TOTAL_MET_ALL = \d+, TOTAL_NM_ALL = \d+, TOTAL_PARAMS_ALL = \d+, TOTAL_TICKETS_ALL = \d+;',
        f'const TOTAL_MET_ALL = {s["t_met"]}, TOTAL_NM_ALL = {s["t_nm"]}, '
        f'TOTAL_PARAMS_ALL = {s["t_params"]}, TOTAL_TICKETS_ALL = {s["t_tix"]};', h)

    # 4. AGENT_STATS
    h = re.sub(r'const AGENT_STATS = \{.*?\};',
        'const AGENT_STATS = ' + json.dumps(s["agent_stats"], ensure_ascii=False) + ';',
        h, flags=re.DOTALL)

    # 5. WEEKLY
    h = re.sub(r'const WEEKLY = \{.*?\};',
        'const WEEKLY = ' + json.dumps(s["weekly"], ensure_ascii=False) + ';',
        h, flags=re.DOTALL)

    # 6. NM_RECORDS
    h = re.sub(r'const NM_RECORDS = \[.*?\];',
        'const NM_RECORDS = ' + json.dumps(s["nm_records"], ensure_ascii=False) + ';',
        h, flags=re.DOTALL)

    # 7. Trend chips
    h = re.sub(r'<div class="trend-chips">.*?</div>',
        '<div class="trend-chips">\n      ' + build_trend_chips(s["weekly"]) + '\n    </div>',
        h, flags=re.DOTALL)

    # 8. Data source banner (inserted before KPI cards)
    src_label = f"Live · {len(file_names)} file(s)" if is_live else "Sample data"
    dot_color = "#00c49a" if is_live else "#4e607e"
    banner = (
        f'<!-- Built: {now} -->\n'
        f'<div style="background:#0c1118;border:1px solid rgba(80,110,180,.15);'
        f'border-radius:8px;padding:8px 16px;margin-bottom:20px;font-family:\'JetBrains Mono\',monospace;'
        f'font-size:11px;color:#6a7e9e;display:flex;align-items:center;gap:10px;flex-wrap:wrap">'
        f'<span style="width:7px;height:7px;border-radius:50%;background:{dot_color};flex-shrink:0"></span>'
        f'{"🟢 Live SharePoint data" if is_live else "📦 Sample data"}'
        f' &nbsp;·&nbsp; {src_label}'
        f' &nbsp;·&nbsp; Built: {now}'
        f'</div>\n'
    )
    h = h.replace('<!-- KPI CARDS -->', banner + '<!-- KPI CARDS -->')

    return h


# ── Main ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    print("\n" + "="*58)
    print("  ECMS Audit Dashboard Builder")
    print("="*58)

    if not os.path.exists(TEMPLATE_FILE):
        print(f"\n❌  Template not found: {TEMPLATE_FILE}")
        print(f"    Rename your HTML file to: {TEMPLATE_FILE}")
        raise SystemExit(1)

    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        template = f.read()
    print(f"\n✅  Template: {TEMPLATE_FILE} ({len(template)//1024}KB)")

    df, file_names, is_live = load_data(SHAREPOINT_FOLDER_URL)
    print(f"{'🟢' if is_live else '📦'}  Source: {file_names}")

    stats = compute_stats(df)
    score = round(stats["t_met"]/stats["t_params"]*100,2) if stats["t_params"] else 0
    print(f"\n📊  {stats['w_range']} · {len(stats['agent_stats'])} engineers "
          f"· {stats['t_tix']} tickets · {stats['t_params']:,} params · {score}% score")

    output = inject(template, stats, file_names, is_live)

    os.makedirs("docs", exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(output)

    print(f"✅  Saved: {OUTPUT_FILE} ({os.path.getsize(OUTPUT_FILE)//1024}KB)")
    print("="*58 + "\n")
