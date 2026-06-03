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


# ══════════════════════════════════════════════════════════════════════
#  DYNAMIC EXECUTIVE SUMMARY GENERATOR
#  Called by inject() to replace all hardcoded text with live stats
# ══════════════════════════════════════════════════════════════════════

def generate_exec_summary(stats):
    """
    Generate all executive summary text dynamically from actual data.
    Returns a dict of text blocks that replace the hardcoded April 2026 text.
    """
    s   = stats
    ag  = s["agent_stats"]
    wk  = s["weekly"]
    nm  = s["nm_records"]

    total_params  = s["t_params"]
    total_met     = s["t_met"]
    total_nm      = s["t_nm"]
    total_tickets = s["t_tix"]
    w_range       = s["w_range"]
    weeks         = s["weeks"]

    overall_score = round(total_met / total_params * 100, 2) if total_params else 0

    # ── Team stats ─────────────────────────────────────────────────
    team_map = {}
    for name, a in ag.items():
        t = a["team"]
        if t not in team_map:
            team_map[t] = {"met": 0, "notMet": 0, "tickets": 0}
        team_map[t]["met"]     += a["met"]
        team_map[t]["notMet"]  += a["notMet"]
        team_map[t]["tickets"] += a["tickets"]

    for t, d in team_map.items():
        tot = d["met"] + d["notMet"]
        d["score"] = round(d["met"] / tot * 100, 2) if tot else 0
        d["total"] = tot

    teams_sorted = sorted(team_map.items(), key=lambda x: x[1]["score"], reverse=True)
    best_team    = teams_sorted[0]   if teams_sorted else ("—", {"score":0,"notMet":0,"tickets":0})
    worst_team   = teams_sorted[-1]  if teams_sorted else ("—", {"score":0,"notMet":0,"tickets":0})

    # ── Weekly stats ────────────────────────────────────────────────
    wk_scores = []
    for w, d in wk.items():
        tot = d["met"] + d["notMet"]
        wk_scores.append((w, round(d["met"]/tot*100, 2) if tot else 0, d["notMet"]))

    best_week  = max(wk_scores, key=lambda x: x[1])  if wk_scores else ("—", 0, 0)
    worst_week = min(wk_scores, key=lambda x: x[1])  if wk_scores else ("—", 0, 0)
    first_week_score = wk_scores[0][1]  if wk_scores else 0
    last_week_score  = wk_scores[-1][1] if wk_scores else 0
    trend_direction  = "improved" if last_week_score >= first_week_score else "decreased"

    # ── Engineer stats ──────────────────────────────────────────────
    eng_scores = []
    for name, a in ag.items():
        tot = a["met"] + a["notMet"]
        eng_scores.append({
            "name":   name,
            "team":   a["team"],
            "met":    a["met"],
            "notMet": a["notMet"],
            "total":  tot,
            "tickets":a["tickets"],
            "score":  round(a["met"]/tot*100, 2) if tot else 0,
        })

    total_engineers = len(eng_scores)
    perfect_count   = sum(1 for e in eng_scores if e["score"] == 100)
    imperfect       = sorted([e for e in eng_scores if e["score"] < 100],
                             key=lambda x: x["score"])

    # ── NM analysis ─────────────────────────────────────────────────
    # Top failing parameter
    param_count = {}
    for r in nm:
        p = r.get("Audit Parameter", "Unknown")
        param_count[p] = param_count.get(p, 0) + 1
    top_param      = max(param_count, key=param_count.get) if param_count else "Closure Information"
    top_param_count = param_count.get(top_param, 0)
    top_param_pct  = round(top_param_count / total_nm * 100) if total_nm else 0

    # Recurring offenders (agents with most NM)
    agent_nm = {}
    for r in nm:
        a = r.get("Agent", "")
        t = r.get("Teams", "")
        if a:
            agent_nm[a] = {"count": agent_nm.get(a, {}).get("count", 0) + 1, "team": t}
    top_offenders = sorted(agent_nm.items(), key=lambda x: x[1]["count"], reverse=True)[:3]

    # NM comment themes
    all_comments = " ".join(r.get("Comments","").lower() for r in nm)
    themes = []
    if "rca" in all_comments or "root cause" in all_comments:
        themes.append("missing RCA sections")
    if "null" in all_comments or "close code" in all_comments:
        themes.append("null close codes")
    if "3-strike" in all_comments or "strike" in all_comments:
        themes.append("3-strike process mismatches")
    if "template" in all_comments:
        themes.append("closure template gaps")
    if not themes:
        themes = ["process adherence gaps"]
    theme_text = ", ".join(themes)

    # ── Top performers (for reco section) ───────────────────────────
    top_performers = sorted([e for e in eng_scores if e["score"] == 100],
                            key=lambda x: x["tickets"], reverse=True)[:3]
    top_names = ", ".join(e["name"] for e in top_performers) if top_performers else "top-performing engineers"

    # ── Offender summary text ────────────────────────────────────────
    if top_offenders:
        off_parts = []
        for name, d in top_offenders:
            off_parts.append(f"{name} ({d['team']}) had {d['count']} Not-Met instance{'s' if d['count']>1 else ''}")
        offender_text = "; ".join(off_parts) + ". These require immediate coaching."
        offender_names = ", ".join(e[0] for e in top_offenders)
    else:
        offender_text  = "No recurring offenders identified this period."
        offender_names = "identified engineers"

    # ── Build all text blocks ────────────────────────────────────────
    month_year = datetime.utcnow().strftime("%B %Y")

    return dict(
        # Exec header
        exec_subtitle = f"Leadership review — ECMS IMS Tower · {month_year} · {w_range}",

        # Strength 1: overall score
        str1_title = f"{overall_score:.2f}% Overall Quality Score",
        str1_body  = (f"The IMS tower delivered {overall_score:.2f}% compliance rate across "
                      f"{total_params:,} audit parameters. "
                      f"Only {total_nm} non-compliance{'s were' if total_nm!=1 else ' was'} recorded, "
                      f"reflecting disciplined process adherence across {total_tickets} tickets."),

        # Strength 2: weekly trend
        str2_title = f"{'Improving' if trend_direction=='improved' else 'Notable'} Weekly Trend",
        str2_body  = (f"{best_week[0]} achieved the lowest Not-Met count ({best_week[2]}) "
                      f"of the entire review period, marking a clear {'upward' if trend_direction=='improved' else 'notable'} trajectory. "
                      f"The quality score {'improved' if trend_direction=='improved' else 'moved'} from "
                      f"{first_week_score:.2f}% in {weeks[0]} to {last_week_score:.2f}% in {weeks[-1]}."),

        # Strength 3: best team
        str3_title = f"{best_team[0]} Team Leading at {best_team[1]['score']:.2f}%",
        str3_body  = (f"{best_team[0]} team achieved {best_team[1]['score']:.2f}% score "
                      f"across {best_team[1]['tickets']} tickets with only "
                      f"{best_team[1]['notMet']} non-compliance{'s' if best_team[1]['notMet']!=1 else ''}. "
                      f"Multiple teams demonstrate consistent adherence to audit parameters."),

        # Strength 4: perfect engineers
        str4_title = f"High Perfect-Score Engineer Count",
        str4_body  = (f"The majority of engineers ({perfect_count} out of {total_engineers}) "
                      f"achieved a perfect 100% compliance score during {w_range}. "
                      f"This demonstrates strong overall team capability and process awareness."),

        # Issue 1: top failing parameter
        iss1_title = f"{top_param} — {top_param_pct}% of Failures",
        iss1_body  = (f"All {total_nm} Not-Met records are attributed to: {top_param}. "
                      f"This indicates a systemic gap in process adoption "
                      f"rather than broad non-compliance across parameters."),

        # Issue 2: worst team
        iss2_title = f"{worst_team[0]} Team — {worst_team[1]['score']:.2f}% Score",
        iss2_body  = (f"{worst_team[0]} team recorded {worst_team[1]['notMet']} Not-Met instance"
                      f"{'s' if worst_team[1]['notMet']!=1 else ''} "
                      f"across {worst_team[1]['total']:,} parameters, the lowest score among all teams. "
                      f"Targeted coaching around {top_param.lower()} is recommended urgently."),

        # Issue 3: recurring offenders
        iss3_title = "Recurring Offenders Identified",
        iss3_body  = offender_text,

        # Issue 4: template inconsistency
        iss4_title = "Template Adoption Inconsistency",
        iss4_body  = (f"Comments reveal varied failure modes: {theme_text}. "
                      f"Standardisation of the {top_param.lower()} workflow is critical "
                      f"to achieving consistent 100% compliance across all teams."),

        # Recommendations
        rec1_body  = (f"Issue a mandatory {top_param} SOP refresher to all IMS engineers. "
                      f"Ensure the updated template includes all required sections. "
                      f"Schedule within the next calendar week as an urgent action item."),

        rec2_title = f"1:1 Coaching — {offender_names}",
        rec2_body  = (f"Schedule immediate 1:1 coaching sessions for {offender_names}. "
                      f"Focus sessions on {top_param.lower()}, close code accuracy, "
                      f"and RCA documentation requirements."),

        rec3_title = f"Recognise {best_week[0]} Performance",
        rec3_body  = (f"Acknowledge {best_week[0]}'s achievement of {best_week[1]:.2f}% formally in team communications. "
                      f"Recognising this performance reinforces desired behaviour across the team."),

        rec4_title = f"{worst_team[0]} Team Structured Review",
        rec4_body  = (f"Conduct a structured quality review specifically for the {worst_team[0]} team. "
                      f"The {worst_team[1]['score']:.2f}% score is the lowest across all teams "
                      f"and requires dedicated process alignment. Target 100% next period."),

        rec5_body  = (f"Establish a clear close code selection guide integrated into ticket management tooling. "
                      f"Null close codes and process mismatches indicate a gap between policy and practice."),

        rec6_body  = (f"Document and share the workflows of top-performing engineers "
                      f"({top_names}) as internal best-practice guides. "
                      f"Peer learning from exemplars accelerates team-wide quality improvement."),
    )


def inject_exec_summary(html, stats):
    """Replace all hardcoded executive summary text with dynamically generated text."""
    t = generate_exec_summary(stats)

    # ── Header subtitle ─────────────────────────────────────────────
    html = re.sub(
        r'<p>Leadership review — ECMS IMS Tower.*?</p>',
        f'<p>{t["exec_subtitle"]}</p>',
        html)

    # ── Strengths section — replace each item ───────────────────────
    # Strength 1: overall score
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)99\.5% Overall Quality Score(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["str1_title"] + m.group(2) + t["str1_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Strength 2: weekly trend
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)Improving Weekly Trend(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["str2_title"] + m.group(2) + t["str2_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Strength 3: best team
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)Workplace &amp; Network Teams Leading(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["str3_title"] + m.group(2) + t["str3_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Strength 4: perfect engineers
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)High Perfect-Score Engineer Count(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["str4_title"] + m.group(2) + t["str4_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # ── Issues section ───────────────────────────────────────────────
    # Issue 1: top failing param
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)Closure Information — 100% of Failures(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["iss1_title"] + m.group(2) + t["iss1_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Issue 2: worst team
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)Security Team — 95\.8% Score(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["iss2_title"] + m.group(2) + t["iss2_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Issue 3: recurring offenders
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)Recurring Offenders Identified(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["iss3_title"] + m.group(2) + t["iss3_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Issue 4: template inconsistency
    html = re.sub(
        r'(<div class="exec-item">.*?<strong>)Template Adoption Inconsistency(</strong>.*?<span>).*?(</span>.*?</div>)',
        lambda m: m.group(1) + t["iss4_title"] + m.group(2) + t["iss4_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # ── Recommendations ──────────────────────────────────────────────
    # Rec 1: mandatory refresher
    html = re.sub(
        r'(<div class="reco-card">.*?<div class="reco-num"[^>]*>1</div>.*?<div class="reco-body">).*?(</div>\s*</div>)',
        lambda m: m.group(1) + t["rec1_body"] + m.group(2),
        html, flags=re.DOTALL, count=1)

    # Rec 2: 1:1 coaching
    html = re.sub(
        r'(<div class="reco-card">.*?<div class="reco-num"[^>]*>2</div>.*?<div class="reco-title">).*?(</div>.*?<div class="reco-body">).*?(</div>\s*</div>)',
        lambda m: m.group(1) + t["rec2_title"] + m.group(2) + t["rec2_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Rec 3: recognise best week
    html = re.sub(
        r'(<div class="reco-card">.*?<div class="reco-num"[^>]*>3</div>.*?<div class="reco-title">).*?(</div>.*?<div class="reco-body">).*?(</div>\s*</div>)',
        lambda m: m.group(1) + t["rec3_title"] + m.group(2) + t["rec3_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Rec 4: worst team review
    html = re.sub(
        r'(<div class="reco-card">.*?<div class="reco-num"[^>]*>4</div>.*?<div class="reco-title">).*?(</div>.*?<div class="reco-body">).*?(</div>\s*</div>)',
        lambda m: m.group(1) + t["rec4_title"] + m.group(2) + t["rec4_body"] + m.group(3),
        html, flags=re.DOTALL, count=1)

    # Rec 5: close code policy
    html = re.sub(
        r'(<div class="reco-card">.*?<div class="reco-num"[^>]*>5</div>.*?<div class="reco-body">).*?(</div>\s*</div>)',
        lambda m: m.group(1) + t["rec5_body"] + m.group(2),
        html, flags=re.DOTALL, count=1)

    # Rec 6: best practices
    html = re.sub(
        r'(<div class="reco-card">.*?<div class="reco-num"[^>]*>6</div>.*?<div class="reco-body">).*?(</div>\s*</div>)',
        lambda m: m.group(1) + t["rec6_body"] + m.group(2),
        html, flags=re.DOTALL, count=1)

    return html


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

    # 9. Dynamic executive summary text
    h = inject_exec_summary(h, stats)

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
