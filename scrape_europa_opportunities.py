"""
European Solidarity Corps - Volunteer Opportunities Scraper for Morocco
=========================================================================
This script fetches all open volunteer opportunities from the European Youth Portal
(youth.europa.eu) that are available for Moroccan residents, and exports them to an Excel file.

It uses the internal REST API of the portal to fetch structured JSON data.
"""

import requests
import pandas as pd
from datetime import datetime
import time
import sys

# Force UTF-8 encoding on standard output to prevent crash when printing emojis on Windows
if sys.stdout.encoding != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

# ─── Configuration ──────────────────────────────────────────────────────────────
API_URL = "https://youth.europa.eu/api/rest/eyp/v1/search_en"
COUNTRY_CODE = "MA"  # Morocco
PAGE_SIZE = 50  # Number of results per API call
PROJECT_BASE_URL = "https://youth.europa.eu/solidarity/opportunity"
OUTPUT_FILE_EXCEL = "Morocco_Volunteer_Opportunities.xlsx"
OUTPUT_FILE_HTML = "Morocco_Volunteer_Opportunities.html"
OUTPUT_FILE_JSON = "Morocco_Volunteer_Opportunities.json"

# Country code to name mapping (for display purposes)
COUNTRY_NAMES = {
    "AT": "Austria", "BE": "Belgium", "BG": "Bulgaria", "CY": "Cyprus",
    "CZ": "Czech Republic", "DE": "Germany", "DK": "Denmark", "EE": "Estonia",
    "EL": "Greece", "ES": "Spain", "FI": "Finland", "FR": "France",
    "HR": "Croatia", "HU": "Hungary", "IE": "Ireland", "IT": "Italy",
    "LT": "Lithuania", "LU": "Luxembourg", "LV": "Latvia", "MT": "Malta",
    "NL": "Netherlands", "PL": "Poland", "PT": "Portugal", "RO": "Romania",
    "SE": "Sweden", "SI": "Slovenia", "SK": "Slovakia", "IS": "Iceland",
    "LI": "Liechtenstein", "NO": "Norway", "MK": "North Macedonia",
    "TR": "Turkey", "AL": "Albania", "AM": "Armenia", "AZ": "Azerbaijan",
    "BA": "Bosnia and Herzegovina", "BY": "Belarus", "DZ": "Algeria",
    "EG": "Egypt", "GE": "Georgia", "IL": "Israel", "JO": "Jordan",
    "LB": "Lebanon", "LY": "Libya", "MA": "Morocco", "MD": "Moldova",
    "ME": "Montenegro", "PS": "Palestine", "RS": "Serbia", "RU": "Russia",
    "SY": "Syria", "TN": "Tunisia", "XK": "Kosovo", "UA": "Ukraine",
}

# ESC Topic code to name mapping
TOPIC_NAMES = {
    "socl": "Social Inclusion",
    "citzn": "Citizenship & Participation",
    "cult": "Culture & Creativity",
    "natr": "Environment & Nature",
    "hlth": "Health & Wellbeing",
    "educ": "Education & Training",
    "emply": "Employment & Entrepreneurship",
    "digi": "Digital Transformation",
    "sport": "Sport",
    "migr": "Migration",
    "peace": "Peace & Conflict",
    "disas": "Disaster Prevention",
}


def fetch_all_opportunities():
    """Fetch all open volunteering opportunities from the ESC API."""
    all_hits = []
    offset = 0
    total = None

    print("=" * 60)
    print("  🇲🇦 European Solidarity Corps - Morocco Opportunities")
    print("=" * 60)
    print()

    while True:
        params = {
            "type": "Opportunity",
            "filters[status]": "open",
            "sort[created]": "desc",
            "size": PAGE_SIZE,
            "from": offset,
        }

        # Retry logic with exponential backoff for 429 errors
        max_retries = 3
        for attempt in range(max_retries + 1):
            try:
                print(f"  📡 Fetching opportunities (offset: {offset})...")
                response = requests.get(API_URL, params=params, timeout=30)
                response.raise_for_status()
                data = response.json()
                break  # Success, exit retry loop
            except requests.exceptions.RequestException as e:
                if hasattr(response, 'status_code') and response.status_code == 429 and attempt < max_retries:
                    wait_time = (attempt + 1) * 5  # 5s, 10s, 15s
                    print(f"  ⚠️ Rate limited (429). Waiting {wait_time}s before retry ({attempt + 1}/{max_retries})...")
                    time.sleep(wait_time)
                    continue
                print(f"  ❌ Error fetching data: {e}")
                return all_hits  # Return what we have so far
        else:
            print(f"  ❌ Max retries reached. Stopping.")
            return all_hits

        hits = data.get("hits", {}).get("hits", [])
        if total is None:
            total = data.get("hits", {}).get("total", {}).get("value", 0)
            print(f"  📊 Total opportunities in database: {total}")
            print()

        if not hits:
            break

        all_hits.extend(hits)
        offset += PAGE_SIZE

        # Show progress
        progress = min(offset, total)
        pct = (progress / total * 100) if total > 0 else 100
        print(f"  ⏳ Progress: {progress}/{total} ({pct:.0f}%)")

        if offset >= total:
            break

        time.sleep(1.5)  # Be nice to the server (increased to avoid 429)

    return all_hits


def filter_morocco_opportunities(all_hits):
    """Filter opportunities where Morocco (MA) is in the residence countries list."""
    morocco_opps = []

    for hit in all_hits:
        source = hit.get("_source", {})
        funding = source.get("funding_programme", {})
        residence_countries = funding.get("residence_countries", [])

        # Check if Morocco (MA) is in the allowed residence countries
        if COUNTRY_CODE in residence_countries:
            morocco_opps.append(source)

    return morocco_opps


def format_date(date_str):
    """Format ISO date string to a readable format."""
    if not date_str:
        return "N/A"
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return date_str


def format_topics(topic_codes):
    """Convert topic codes to readable names."""
    if not topic_codes:
        return "N/A"
    names = [TOPIC_NAMES.get(code, code) for code in topic_codes]
    return " | ".join(names)


def build_dataframe(opportunities):
    """Build a pandas DataFrame from the list of opportunities."""
    rows = []

    for opp in opportunities:
        country_code = opp.get("country", "")
        country_name = COUNTRY_NAMES.get(country_code, country_code)

        # Build the project URL
        opp_id = opp.get("opid", "")
        project_url = f"{PROJECT_BASE_URL}/{opp_id}_en" if opp_id else "N/A"

        # Format volunteer countries
        vol_countries = opp.get("volunteer_countries", [])
        if vol_countries == ["all"]:
            vol_countries_str = "All countries"
        else:
            vol_countries_str = ", ".join(
                [COUNTRY_NAMES.get(c, c) for c in vol_countries]
            )

        # Determine if application deadline exists
        deadline = opp.get("date_application_end", "")
        has_no_deadline = opp.get("has_no_deadline", False)
        if has_no_deadline:
            deadline_str = "No deadline (open until filled)"
        else:
            deadline_str = format_date(deadline)

        row = {
            "ID": opp.get("opid", ""),
            "Title": opp.get("title", "N/A"),
            "Organisation": opp.get("organisation_name", "N/A"),
            "Country": country_name,
            "City": opp.get("town", "N/A"),
            "Description": opp.get("description", "N/A"),
            "Participant Profile / Requirements": opp.get("participant_profile", "N/A"),
            "Accommodation & Board": opp.get("boarding_arrangements", "N/A"),
            "Training Provided": opp.get("training", "N/A"),
            "Start Date": format_date(opp.get("date_start", "")),
            "End Date": format_date(opp.get("date_end", "")),
            "Application Deadline": deadline_str,
            "Topics": format_topics(opp.get("esc_topics", [])),
            "Activity Type": opp.get("volunteer_activity_type", "N/A"),
            "Strand": " | ".join(opp.get("strand", [])),
            "CV Required": "Yes" if opp.get("requires_cv") else "No",
            "Motivation Letter Required": "Yes" if opp.get("requires_motivation_statement") else "No",
            "Contact Person": opp.get("contact_person_name", "N/A"),
            "Contact Email": opp.get("contact_person_email", "N/A"),
            "Hosting Country for Volunteers": vol_countries_str,
            "Project URL": project_url,
        }

        rows.append(row)

    return pd.DataFrame(rows)


def export_to_excel(df, filename):
    """Export DataFrame to a nicely formatted Excel file."""
    import re
    # XML control character regex to remove illegal characters for Excel (openpyxl)
    illegal_chars_re = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
    
    def clean_illegal_chars(val):
        if isinstance(val, str):
            return illegal_chars_re.sub("", val)
        return val

    if hasattr(df, 'map'):
        df_clean = df.map(clean_illegal_chars)
    else:
        df_clean = df.applymap(clean_illegal_chars)

    with pd.ExcelWriter(filename, engine="openpyxl") as writer:
        df_clean.to_excel(writer, index=False, sheet_name="Opportunities")

        # Auto-adjust column widths
        worksheet = writer.sheets["Opportunities"]
        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter
            for cell in column_cells:
                try:
                    if cell.value:
                        max_length = max(max_length, min(len(str(cell.value)), 50))
                except:
                    pass
            adjusted_width = max_length + 2
            worksheet.column_dimensions[column_letter].width = adjusted_width

    print(f"  ✅ Saved Excel to: {filename}")


def export_to_html(df, filename):
    """Export DataFrame to an interactive premium HTML dashboard."""
    import json
    opportunities_list = df.to_dict(orient="records")
    json_data = json.dumps(opportunities_list, ensure_ascii=False)
    
    html_template = """<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>بوابة فرص التطوع الأوروبية للمغاربة 🇲🇦</title>
    <!-- Bootstrap 5 CSS (RTL) -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.rtl.min.css">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts: Cairo (Arabic) & Outfit (Numbers/English) -->
    <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700;800&family=Outfit:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    
    <style>
        :root {
            --primary: #0f62fe;
            --primary-hover: #0043ce;
            --primary-light: #edf5ff;
            --bg: #f4f6fa;
            --card-bg: #ffffff;
            --text: #161616;
            --text-muted: #525252;
            --border: #e0e0e0;
            --success: #24a148;
            --warning: #f1c21b;
            --danger: #da1e28;
            --shadow: 0 4px 20px rgba(0, 0, 0, 0.05);
            --shadow-hover: 0 10px 30px rgba(0, 0, 0, 0.1);
            --radius: 12px;
            --radius-lg: 16px;
            --transition: all 0.3s cubic-bezier(0.16, 1, 0.3, 1);
        }

        body.dark-mode {
            --bg: #121212;
            --card-bg: #1e1e1e;
            --text: #f4f4f4;
            --text-muted: #a8a8a8;
            --border: #353535;
            --shadow: 0 4px 25px rgba(0, 0, 0, 0.25);
            --shadow-hover: 0 10px 35px rgba(0, 0, 0, 0.35);
            --primary-light: #262626;
        }

        body {
            font-family: 'Cairo', system-ui, -apple-system, sans-serif;
            background-color: var(--bg);
            color: var(--text);
            transition: var(--transition);
            direction: rtl;
            text-align: right;
            padding-bottom: 60px;
        }

        .font-en {
            font-family: 'Outfit', sans-serif;
        }

        /* Scrollbar styling */
        ::-webkit-scrollbar {
            width: 8px;
            height: 8px;
        }
        ::-webkit-scrollbar-track {
            background: var(--bg);
        }
        ::-webkit-scrollbar-thumb {
            background: var(--border);
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: var(--text-muted);
        }

        /* Banner Header */
        .dashboard-header {
            background: linear-gradient(135deg, #0b1c3d 0%, #0f62fe 100%);
            color: white;
            padding: 40px 0;
            border-radius: 0 0 24px 24px;
            box-shadow: 0 10px 35px rgba(15, 98, 254, 0.15);
            margin-bottom: 40px;
            position: relative;
        }
        
        /* Stat Cards */
        .stat-card {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 20px;
            box-shadow: var(--shadow);
            transition: var(--transition);
            height: 100%;
            display: flex;
            align-items: center;
            gap: 16px;
        }
        .stat-card:hover {
            transform: translateY(-4px);
            box-shadow: var(--shadow-hover);
        }
        .stat-icon {
            width: 56px;
            height: 56px;
            border-radius: var(--radius);
            background: var(--primary-light);
            color: var(--primary);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.5rem;
            flex-shrink: 0;
        }
        .stat-value {
            font-size: 1.6rem;
            font-weight: 800;
            line-height: 1.2;
            color: var(--text);
        }
        .stat-label {
            font-size: 0.85rem;
            color: var(--text-muted);
            font-weight: 600;
        }

        /* Search & Filters Sidebar */
        .filter-sidebar {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 24px;
            box-shadow: var(--shadow);
            height: fit-content;
            position: sticky;
            top: 24px;
            transition: var(--transition);
        }
        .filter-group {
            margin-bottom: 18px;
        }
        .filter-label {
            font-weight: 700;
            font-size: 0.88rem;
            margin-bottom: 8px;
            display: block;
            color: var(--text);
        }
        .custom-input {
            background-color: var(--bg);
            color: var(--text);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 10px 14px;
            width: 100%;
            transition: var(--transition);
            font-size: 0.9rem;
        }
        .custom-input:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px var(--primary-light);
        }
        
        .checkbox-container {
            display: flex;
            align-items: center;
            gap: 8px;
            margin-bottom: 8px;
            cursor: pointer;
            font-size: 0.85rem;
            font-weight: 600;
        }
        .checkbox-container input {
            cursor: pointer;
            width: 16px;
            height: 16px;
        }

        /* Opportunity Cards */
        .opp-card {
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: var(--radius-lg);
            padding: 24px;
            box-shadow: var(--shadow);
            transition: var(--transition);
            height: 100%;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
            position: relative;
            overflow: hidden;
        }
        .opp-card::before {
            content: '';
            position: absolute;
            top: 0;
            right: 0;
            width: 4px;
            height: 100%;
            background: var(--primary);
            opacity: 0;
            transition: var(--transition);
        }
        .opp-card:hover {
            transform: translateY(-5px);
            box-shadow: var(--shadow-hover);
        }
        .opp-card:hover::before {
            opacity: 1;
        }
        .opp-card-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 12px;
            gap: 10px;
        }
        .opp-location {
            font-size: 0.85rem;
            font-weight: 700;
            color: var(--text-muted);
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .opp-title {
            font-size: 1.12rem;
            font-weight: 800;
            line-height: 1.4;
            margin-bottom: 8px;
            color: var(--text);
            display: -webkit-box;
            -webkit-line-clamp: 2;
            -webkit-box-orient: vertical;
            overflow: hidden;
            height: 2.8em;
            cursor: pointer;
            transition: var(--transition);
        }
        .opp-title:hover {
            color: var(--primary);
        }
        .opp-org {
            font-size: 0.85rem;
            color: var(--text-muted);
            font-weight: 600;
            margin-bottom: 12px;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .opp-badges {
            display: flex;
            flex-wrap: wrap;
            gap: 6px;
            margin-bottom: 16px;
        }
        .opp-badge {
            font-size: 0.72rem;
            padding: 5px 12px;
            border-radius: 50px;
            font-weight: 700;
        }
        .badge-topic {
            background: var(--primary-light);
            color: var(--primary);
        }
        .badge-deadline {
            background: rgba(218, 30, 40, 0.08);
            color: var(--danger);
        }
        .badge-deadline.safe {
            background: rgba(36, 161, 72, 0.08);
            color: var(--success);
        }
        .badge-ongoing {
            background: rgba(241, 194, 27, 0.08);
            color: #a67c00;
        }
        .opp-footer {
            border-top: 1px solid var(--border);
            padding-top: 14px;
            margin-top: auto;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .opp-req-badge {
            font-size: 0.75rem;
            font-weight: 600;
            color: var(--text-muted);
            display: flex;
            align-items: center;
            gap: 4px;
        }

        /* Theme button & Action Buttons */
        .btn-theme-switch {
            width: 42px;
            height: 42px;
            border-radius: 50%;
            border: 1px solid rgba(255, 255, 255, 0.25);
            background: rgba(255, 255, 255, 0.15);
            color: white;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 1.1rem;
            cursor: pointer;
            transition: var(--transition);
        }
        .btn-theme-switch:hover {
            background: rgba(255, 255, 255, 0.3);
            transform: rotate(15deg);
        }
        
        .btn-primary-custom {
            background-color: var(--primary);
            color: white;
            border: none;
            border-radius: var(--radius);
            padding: 10px 20px;
            font-weight: 700;
            transition: var(--transition);
        }
        .btn-primary-custom:hover {
            background-color: var(--primary-hover);
            color: white;
        }
        .btn-outline-custom {
            background-color: transparent;
            color: var(--text);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 10px 20px;
            font-weight: 700;
            transition: var(--transition);
        }
        .btn-outline-custom:hover {
            background-color: var(--bg);
            border-color: var(--text-muted);
        }

        /* Modal & Tabs */
        .modal-dialog {
            max-width: 800px;
        }
        .nav-pills {
            background: var(--bg);
            padding: 6px;
            border-radius: var(--radius);
            display: inline-flex;
            gap: 4px;
            margin-bottom: 24px;
        }
        .nav-pills .nav-link {
            color: var(--text-muted);
            font-weight: 700;
            padding: 8px 20px;
            border-radius: 8px;
            border: none;
            transition: var(--transition);
        }
        .nav-pills .nav-link.active {
            background-color: var(--primary);
            color: white;
        }
        .modal-meta-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 16px;
            margin-bottom: 24px;
        }
        .modal-meta-item {
            background: var(--bg);
            border: 1px solid var(--border);
            border-radius: var(--radius);
            padding: 14px;
        }
        .meta-label {
            font-size: 0.78rem;
            color: var(--text-muted);
            font-weight: 700;
            margin-bottom: 4px;
        }
        .meta-value {
            font-size: 0.95rem;
            font-weight: 700;
            color: var(--text);
        }
        
        .translate-box {
            background: var(--primary-light);
            border-right: 4px solid var(--primary);
            border-radius: var(--radius);
            padding: 16px;
            margin-top: 15px;
            font-size: 0.95rem;
            line-height: 1.6;
        }

        /* Log Panel styling */
        #logContainer {
            border-radius: var(--radius);
            font-size: 0.85rem;
            line-height: 1.5;
            box-shadow: inset 0 2px 8px rgba(0,0,0,0.2);
        }

        /* Pagination style */
        .pagination-container {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 8px;
            margin-top: 40px;
        }
        .page-btn {
            width: 40px;
            height: 40px;
            border-radius: 8px;
            border: 1px solid var(--border);
            background: var(--card-bg);
            color: var(--text);
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: 700;
            cursor: pointer;
            transition: var(--transition);
        }
        .page-btn.active {
            background: var(--primary);
            color: white;
            border-color: var(--primary);
        }
        .page-btn:hover:not(.active):not(:disabled) {
            background: var(--bg);
        }
        .page-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }

        /* Loading indicator spinner */
        .loader-spinner {
            display: inline-block;
            width: 1.5rem;
            height: 1.5rem;
            border: 3px solid rgba(255,255,255,.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 1s ease-in-out infinite;
        }
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
    </style>
</head>
<body>

    <!-- Header Banner -->
    <header class="dashboard-header">
        <div class="container">
            <div class="d-flex justify-content-between align-items-center flex-wrap gap-3">
                <div>
                    <h1 class="fw-extrabold mb-1 fs-3">🇪🇺 فرص التطوع في أوروبا للمغاربة 🇲🇦</h1>
                    <p class="text-white-50 mb-0 fs-6">بوابة البحث الرسمية عن الفرص المتاحة في فيلق التضامن الأوروبي (ESC) للمقيمين في المغرب</p>
                </div>
                <div class="d-flex align-items-center gap-3">
                    <button class="btn-theme-switch" id="themeToggle" onclick="toggleTheme()" title="تغيير المظهر">
                        <i class="fa-solid fa-moon"></i>
                    </button>
                    <button id="updateBtn" class="btn btn-success fw-bold px-4 py-2" style="border-radius: 12px;" onclick="startUpdate()">
                        <i class="fa-solid fa-arrows-rotate me-2"></i> تحديث البيانات
                    </button>
                </div>
            </div>
            
            <div class="mt-2 text-white-50 fs-7 text-start dir-ltr">
                تاريخ التحديث: <span id="lastUpdated" class="font-en">__LAST_UPDATED__</span>
            </div>
        </div>
    </header>

    <div class="container">
        
        <!-- Live Scraper Log Terminal -->
        <div id="logContainer" class="mb-4 p-3 bg-dark text-light rounded" style="display: none; height: 250px; overflow-y: auto; font-family: monospace; direction: ltr; text-align: left;">
            <div id="logContent"></div>
        </div>

        <!-- 3 Stats Cards Row -->
        <div class="row g-4 mb-4">
            <div class="col-md-4">
                <div class="stat-card">
                    <div class="stat-icon"><i class="fa-solid fa-briefcase"></i></div>
                    <div>
                        <div class="stat-value font-en" id="statTotal">0</div>
                        <div class="stat-label">إجمالي الفرص المتاحة</div>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="stat-card">
                    <div class="stat-icon" style="background: rgba(21da, 30, 40, 0.1); color: var(--danger);"><i class="fa-solid fa-clock"></i></div>
                    <div>
                        <div class="stat-value font-en" id="statExpiring">0</div>
                        <div class="stat-label">فرص تنتهي قريباً (&lt; 7 أيام)</div>
                    </div>
                </div>
            </div>
            <div class="col-md-4">
                <div class="stat-card">
                    <div class="stat-icon" style="background: rgba(36, 161, 72, 0.1); color: var(--success);"><i class="fa-solid fa-earth-europe"></i></div>
                    <div>
                        <div class="stat-value fs-5" id="statTopCountry">N/A</div>
                        <div class="stat-label">الدولة الأكثر طلباً</div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Main Body Grid: Filters + Opportunities Grid -->
        <div class="row g-4">
            
            <!-- Filters Sidebar -->
            <div class="col-lg-3">
                <aside class="filter-sidebar">
                    <div class="d-flex justify-content-between align-items-center mb-4">
                        <h5 class="fw-bold mb-0"><i class="fa-solid fa-filter me-2 text-primary"></i>تصفية الفرص</h5>
                        <button class="btn btn-link btn-sm text-decoration-none p-0" onclick="resetFilters()">إعادة ضبط</button>
                    </div>

                    <!-- Search Input -->
                    <div class="filter-group">
                        <label class="filter-label" for="searchFilter">بحث بالكلمات الدليلة</label>
                        <div class="position-relative">
                            <input type="text" id="searchFilter" class="custom-input" placeholder="اسم المشروع، المنظمة، المدينة...">
                        </div>
                    </div>

                    <!-- Country Filter -->
                    <div class="filter-group">
                        <label class="filter-label" for="countryFilter">الدولة المستضيفة</label>
                        <select id="countryFilter" class="custom-input">
                            <option value="all">كل الدول</option>
                        </select>
                    </div>

                    <!-- Topic Filter -->
                    <div class="filter-group">
                        <label class="filter-label" for="topicFilter">مجال التطوع</label>
                        <select id="topicFilter" class="custom-input">
                            <option value="all">كل المجالات</option>
                        </select>
                    </div>

                    <!-- Deadline Type Filter -->
                    <div class="filter-group">
                        <label class="filter-label" for="deadlineFilter">أجل التقديم</label>
                        <select id="deadlineFilter" class="custom-input">
                            <option value="all">الكل</option>
                            <option value="has_deadline">فرص محددة بأجل</option>
                            <option value="ongoing">مفتوح دائماً (بدون أجل)</option>
                            <option value="expiring_soon">تنتهي قريباً (أقل من أسبوع)</option>
                        </select>
                    </div>

                    <!-- Documents Filters -->
                    <div class="filter-group">
                        <label class="filter-label">المستندات المطلوبة</label>
                        <div class="checkbox-container">
                            <input type="checkbox" id="cvFilter" onchange="handleCheckboxChange()">
                            <span>السيرة الذاتية (CV) غير مطلوبة</span>
                        </div>
                        <div class="checkbox-container">
                            <input type="checkbox" id="motivationFilter" onchange="handleCheckboxChange()">
                            <span>الرسالة التحفيزية غير مطلوبة</span>
                        </div>
                    </div>

                    <!-- Sort -->
                    <div class="filter-group">
                        <label class="filter-label" for="sortSelect">ترتيب حسب</label>
                        <select id="sortSelect" class="custom-input" onchange="handleSortChange()">
                            <option value="start_newest">تاريخ البدء (الأحدث أولاً)</option>
                            <option value="start_oldest">تاريخ البدء (الأقدم أولاً)</option>
                            <option value="deadline_soonest">أجل التقديم (الأقرب انتهاءً)</option>
                            <option value="title_asc">اسم الفرصة (أ - ي)</option>
                        </select>
                    </div>
                </aside>
            </div>

            <!-- Opportunities Grid Content Area -->
            <div class="col-lg-9">
                <div class="row g-4" id="opportunitiesGrid">
                    <!-- Cards will be dynamically inserted here -->
                </div>
                
                <!-- Pagination container -->
                <div class="pagination-container" id="paginationControls">
                    <!-- Dynamic page buttons -->
                </div>
            </div>
            
        </div>
    </div>

    <!-- Opportunity Details Modal -->
    <div class="modal fade" id="oppModal" tabindex="-1" aria-labelledby="oppModalLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered">
            <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title fw-bold" id="oppModalLabel">اسم الفرصة</h5>
                    <button type="button" class="btn-close ms-0 me-auto" data-bs-dismiss="modal" aria-label="Close"></button>
                </div>
                <div class="modal-body">
                    
                    <!-- Navigation Pills for Tabs -->
                    <ul class="nav nav-pills" id="modalTab" role="tablist">
                        <li class="nav-item" role="presentation">
                            <button class="nav-link active" id="about-tab" data-bs-toggle="tab" data-bs-target="#about-pane" type="button" role="tab" aria-controls="about-pane" aria-selected="true">حول الفرصة</button>
                        </li>
                        <li class="nav-item" role="presentation">
                            <button class="nav-link" id="details-tab" data-bs-toggle="tab" data-bs-target="#details-pane" type="button" role="tab" aria-controls="details-pane" aria-selected="false">السكن والتدريب</button>
                        </li>
                        <li class="nav-item" role="presentation">
                            <button class="nav-link" id="contact-tab" data-bs-toggle="tab" data-bs-target="#contact-pane" type="button" role="tab" aria-controls="contact-pane" aria-selected="false">معلومات الاتصال والتقديم</button>
                        </li>
                    </ul>
                    
                    <div class="tab-content" id="modalTabContent">
                        
                        <!-- Tab 1: About -->
                        <div class="tab-pane fade show active" id="about-pane" role="tabpanel" aria-labelledby="about-tab">
                            <div class="modal-meta-grid">
                                <div class="modal-meta-item">
                                    <div class="meta-label">الدولة والمدينة</div>
                                    <div class="meta-value" id="modalLocation">N/A</div>
                                </div>
                                <div class="modal-meta-item">
                                    <div class="meta-label">تاريخ النشاط</div>
                                    <div class="meta-value font-en" id="modalDates">N/A</div>
                                </div>
                                <div class="modal-meta-item">
                                    <div class="meta-label">أجل التقديم</div>
                                    <div class="meta-value font-en" id="modalDeadline">N/A</div>
                                </div>
                            </div>
                            
                            <div class="modal-details-section">
                                <div class="modal-details-title d-flex justify-content-between align-items-center">
                                    <span>وصف مشروع التطوع</span>
                                    <button class="btn btn-sm btn-outline-primary py-1 px-3" onclick="triggerTranslate('Description')">
                                        <i class="fa-solid fa-language me-1"></i> ترجمة للعربية
                                    </button>
                                </div>
                                <div id="modalDescription" style="white-space: pre-line; line-height: 1.6;">N/A</div>
                                <div id="modalDescription_trans" class="translate-box" style="display: none;"></div>
                            </div>

                            <div class="modal-details-section">
                                <div class="modal-details-title d-flex justify-content-between align-items-center">
                                    <span>شروط المشارك / الملف الشخصي</span>
                                    <button class="btn btn-sm btn-outline-primary py-1 px-3" onclick="triggerTranslate('Profile')">
                                        <i class="fa-solid fa-language me-1"></i> ترجمة للعربية
                                    </button>
                                </div>
                                <div id="modalProfile" style="white-space: pre-line; line-height: 1.6;">N/A</div>
                                <div id="modalProfile_trans" class="translate-box" style="display: none;"></div>
                            </div>
                        </div>
                        
                        <!-- Tab 2: Logistics / Training -->
                        <div class="tab-pane fade" id="details-pane" role="tabpanel" aria-labelledby="details-tab">
                            <div class="modal-details-section">
                                <div class="modal-details-title d-flex justify-content-between align-items-center">
                                    <span>الإقامة، الإعاشة وتدبير النقل</span>
                                    <button class="btn btn-sm btn-outline-primary py-1 px-3" onclick="triggerTranslate('Boarding')">
                                        <i class="fa-solid fa-language me-1"></i> ترجمة للعربية
                                    </button>
                                </div>
                                <div id="modalBoarding" style="white-space: pre-line; line-height: 1.6;">N/A</div>
                                <div id="modalBoarding_trans" class="translate-box" style="display: none;"></div>
                            </div>

                            <div class="modal-details-section">
                                <div class="modal-details-title d-flex justify-content-between align-items-center">
                                    <span>التدريب والتكوين الموفر</span>
                                    <button class="btn btn-sm btn-outline-primary py-1 px-3" onclick="triggerTranslate('Training')">
                                        <i class="fa-solid fa-language me-1"></i> ترجمة للعربية
                                    </button>
                                </div>
                                <div id="modalTraining" style="white-space: pre-line; line-height: 1.6;">N/A</div>
                                <div id="modalTraining_trans" class="translate-box" style="display: none;"></div>
                            </div>
                        </div>
                        
                        <!-- Tab 3: Contact & Apply -->
                        <div class="tab-pane fade" id="contact-pane" role="tabpanel" aria-labelledby="contact-tab">
                            <div class="modal-details-section">
                                <div class="modal-details-title">المنظمة المستضيفة</div>
                                <div id="modalOrgName" class="fw-bold mb-2">N/A</div>
                                <div id="modalTopicsList" class="opp-badges mb-2"></div>
                            </div>
                            
                            <div class="modal-details-section">
                                <div class="modal-details-title">معلومات التواصل المباشر</div>
                                <div class="row g-3 mt-1">
                                    <div class="col-md-6">
                                        <strong>اسم المسؤول:</strong> <span id="modalContactName">N/A</span>
                                    </div>
                                    <div class="col-md-6">
                                        <strong>البريد الإلكتروني:</strong> 
                                        <span id="modalContactEmail">N/A</span>
                                        <button class="btn btn-sm btn-link text-decoration-none p-0 ms-2" onclick="copyEmail()"><i class="fa-regular fa-copy"></i> نسخ</button>
                                    </div>
                                </div>
                            </div>

                            <div class="modal-details-section">
                                <div class="modal-details-title">المستندات المطلوبة للتقديم</div>
                                <div class="d-flex gap-4 mt-2">
                                    <div>
                                        <i class="fa-solid id-card-clip text-primary me-2" id="modalCvReqIcon"></i> 
                                        السيرة الذاتية (CV): <span id="modalCvReqText" class="fw-bold">N/A</span>
                                    </div>
                                    <div>
                                        <i class="fa-solid fa-envelope-open-text text-primary me-2" id="modalMotivationReqIcon"></i>
                                        الرسالة التحفيزية (Motivation Statement): <span id="modalMotivationReqText" class="fw-bold">N/A</span>
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-outline-custom" data-bs-dismiss="modal">إغلاق</button>
                    <button type="button" class="btn btn-outline-custom" id="btnCopyDetails" onclick="copyDetailsLink()"><i class="fa-solid fa-share-nodes me-2"></i> مشاركة الرابط</button>
                    <a href="#" target="_blank" class="btn btn-primary-custom" id="btnApplyLink"><i class="fa-solid fa-circle-arrow-up-right me-2"></i> تقديم على الموقع الرسمي</a>
                </div>
            </div>
        </div>
    </div>

    <!-- Opportunities Embedded JSON Data -->
    <script id="opportunities-data" type="application/json">
/*__OPPORTUNITIES_JSON__*/
    </script>

    <!-- JS dependencies -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>

    <script>
        // Data Load & Global variables
        let opportunities = [];
        try {
            opportunities = JSON.parse(document.getElementById('opportunities-data').textContent);
        } catch (e) {
            console.error("Error reading JSON data", e);
        }

        let filteredOpportunities = [...opportunities];
        let currentPage = 1;
        const cardsPerPage = 9;

        // Current Active Filters
        let filters = {
            search: '',
            country: 'all',
            topic: 'all',
            deadlineType: 'all',
            cvNotRequired: false,
            motivationNotRequired: false
        };
        let sortBy = 'start_newest';
        
        let currentModalOpp = null; // Store references to current active modal item

        const COUNTRY_FLAGS = {
            "Austria": "🇦🇹", "Belgium": "🇧🇪", "Bulgaria": "🇧🇬", "Cyprus": "🇨🇾",
            "Czech Republic": "🇨🇿", "Germany": "🇩🇪", "Denmark": "🇩🇰", "Estonia": "🇪🇪",
            "Greece": "🇬🇷", "Spain": "🇪🇸", "Finland": "🇫🇮", "France": "🇫🇷",
            "Croatia": "🇭🇷", "Hungary": "🇭🇺", "Ireland": "🇮🇪", "Italy": "🇮🇹",
            "Lithuania": "🇱🇹", "Luxembourg": "🇱🇺", "Latvia": "🇱🇻", "Malta": "🇲🇹",
            "Netherlands": "🇳🇱", "Poland": "🇵🇱", "Portugal": "🇵🇹", "Romania": "🇷🇴",
            "Sweden": "🇸🇪", "Slovenia": "🇸🇮", "Slovakia": "🇸🇰", "Iceland": "🇮🇸",
            "Liechtenstein": "🇱🇮", "Norway": "🇳🇴", "North Macedonia": "🇲🇰",
            "Turkey": "🇹🇷", "Albania": "🇦🇱", "Armenia": "🇦🇲", "Azerbaijan": "🇦🇿",
            "Bosnia and Herzegovina": "🇧🇦", "Belarus": "🇧🇾", "Algeria": "🇩🇿",
            "Egypt": "🇪🇬", "Georgia": "🇬🇪", "Israel": "🇮🇱", "Jordan": "🇯🇴",
            "Lebanon": "🇱🇧", "Libya": "🇱🇾", "Morocco": "🇲🇦", "Moldova": "🇲🇩",
            "Montenegro": "🇲🇪", "Palestine": "🇵🇸", "Serbia": "🇷🇸", "Russia": "🇷🇺",
            "Syria": "🇸🇾", "Tunisia": "🇹🇳", "Kosovo": "🇽🇰", "Ukraine": "🇺🇦"
        };

        function getCountryFlag(countryName) {
            return COUNTRY_FLAGS[countryName] || "🌍";
        }

        // Initialize App on DOM Loaded
        document.addEventListener('DOMContentLoaded', () => {
            // Apply theme from localStorage if saved
            if (localStorage.getItem('theme') === 'dark') {
                document.body.classList.add('dark-mode');
                const themeIcon = document.querySelector('#themeToggle i');
                themeIcon.className = 'fa-solid fa-sun';
            }

            setupFiltersDropdowns();
            applyFiltersAndRender();

            // Set up search keyup event
            document.getElementById('searchFilter').addEventListener('input', (e) => {
                filters.search = e.target.value;
                currentPage = 1;
                applyFiltersAndRender();
            });

            // Set up dropdown events
            document.getElementById('countryFilter').addEventListener('change', (e) => {
                filters.country = e.target.value;
                currentPage = 1;
                applyFiltersAndRender();
            });
            document.getElementById('topicFilter').addEventListener('change', (e) => {
                filters.topic = e.target.value;
                currentPage = 1;
                applyFiltersAndRender();
            });
            document.getElementById('deadlineFilter').addEventListener('change', (e) => {
                filters.deadlineType = e.target.value;
                currentPage = 1;
                applyFiltersAndRender();
            });
        });

        // Theme Toggle
        function toggleTheme() {
            const body = document.body;
            body.classList.toggle('dark-mode');
            const themeIcon = document.querySelector('#themeToggle i');
            
            if (body.classList.contains('dark-mode')) {
                themeIcon.className = 'fa-solid fa-sun';
                localStorage.setItem('theme', 'dark');
            } else {
                themeIcon.className = 'fa-solid fa-moon';
                localStorage.setItem('theme', 'light');
            }
        }

        // Setup filter options dynamically based on scraped data
        function setupFiltersDropdowns() {
            const countries = new Set();
            const topics = new Set();

            opportunities.forEach(opp => {
                if (opp.Country) countries.add(opp.Country);
                if (opp.Topics) {
                    opp.Topics.split('|').forEach(t => {
                        const cleanT = t.trim();
                        if (cleanT && cleanT !== 'N/A') topics.add(cleanT);
                    });
                }
            });

            const countrySelect = document.getElementById('countryFilter');
            Array.from(countries).sort().forEach(country => {
                const opt = document.createElement('option');
                opt.value = country;
                opt.textContent = `${getCountryFlag(country)} ${country}`;
                countrySelect.appendChild(opt);
            });

            const topicSelect = document.getElementById('topicFilter');
            Array.from(topics).sort().forEach(topic => {
                const opt = document.createElement('option');
                opt.value = topic;
                opt.textContent = topic;
                topicSelect.appendChild(opt);
            });
        }

        function handleCheckboxChange() {
            filters.cvNotRequired = document.getElementById('cvFilter').checked;
            filters.motivationNotRequired = document.getElementById('motivationFilter').checked;
            currentPage = 1;
            applyFiltersAndRender();
        }

        function handleSortChange() {
            sortBy = document.getElementById('sortSelect').value;
            currentPage = 1;
            applyFiltersAndRender();
        }

        function resetFilters() {
            document.getElementById('searchFilter').value = '';
            document.getElementById('countryFilter').value = 'all';
            document.getElementById('topicFilter').value = 'all';
            document.getElementById('deadlineFilter').value = 'all';
            document.getElementById('cvFilter').checked = false;
            document.getElementById('motivationFilter').checked = false;
            document.getElementById('sortSelect').value = 'start_newest';

            filters = {
                search: '',
                country: 'all',
                topic: 'all',
                deadlineType: 'all',
                cvNotRequired: false,
                motivationNotRequired: false
            };
            sortBy = 'start_newest';
            currentPage = 1;
            applyFiltersAndRender();
        }

        // Helper calculations for deadline dates
        function getDeadlineDate(str) {
            if (!str || str.includes('No deadline')) return null;
            const d = new Date(str);
            return isNaN(d.getTime()) ? null : d;
        }

        function getDaysRemaining(str) {
            const d = getDeadlineDate(str);
            if (!d) return null;
            const today = new Date();
            today.setHours(0,0,0,0);
            const diffTime = d - today;
            return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        }

        // Filters + Sorting + Render pipeline
        function applyFiltersAndRender() {
            // Apply filter logic
            filteredOpportunities = opportunities.filter(opp => {
                // Search query
                if (filters.search) {
                    const q = filters.search.toLowerCase();
                    const inTitle = (opp.Title || '').toLowerCase().includes(q);
                    const inDesc = (opp.Description || '').toLowerCase().includes(q);
                    const inCity = (opp.City || '').toLowerCase().includes(q);
                    const inOrg = (opp.Organisation || '').toLowerCase().includes(q);
                    const inTopics = (opp.Topics || '').toLowerCase().includes(q);
                    if (!inTitle && !inDesc && !inCity && !inOrg && !inTopics) return false;
                }

                // Country
                if (filters.country !== 'all' && opp.Country !== filters.country) return false;

                // Topic
                if (filters.topic !== 'all' && !(opp.Topics || '').includes(filters.topic)) return false;

                // CV Required
                if (filters.cvNotRequired && opp['CV Required'] !== 'No') return false;

                // Motivation Required
                if (filters.motivationNotRequired && opp['Motivation Letter Required'] !== 'No') return false;

                // Deadline type
                if (filters.deadlineType !== 'all') {
                    const isOngoing = opp['Application Deadline'].includes('No deadline');
                    if (filters.deadlineType === 'ongoing' && !isOngoing) return false;
                    if (filters.deadlineType === 'has_deadline' && isOngoing) return false;
                    if (filters.deadlineType === 'expiring_soon') {
                        if (isOngoing) return false;
                        const rem = getDaysRemaining(opp['Application Deadline']);
                        if (rem === null || rem < 0 || rem > 7) return false;
                    }
                }

                return true;
            });

            // Sort
            if (sortBy === 'start_newest') {
                filteredOpportunities.sort((a, b) => new Date(b['Start Date'] || 0) - new Date(a['Start Date'] || 0));
            } else if (sortBy === 'start_oldest') {
                filteredOpportunities.sort((a, b) => new Date(a['Start Date'] || 0) - new Date(b['Start Date'] || 0));
            } else if (sortBy === 'deadline_soonest') {
                filteredOpportunities.sort((a, b) => {
                    const deadlineA = getDeadlineDate(a['Application Deadline']);
                    const deadlineB = getDeadlineDate(b['Application Deadline']);
                    if (!deadlineA) return 1;
                    if (!deadlineB) return -1;
                    return deadlineA - deadlineB;
                });
            } else if (sortBy === 'title_asc') {
                filteredOpportunities.sort((a, b) => (a.Title || '').localeCompare(b.Title || ''));
            }

            renderStats();
            renderOpportunitiesGrid();
            renderPagination();
        }

        // Render Stats values
        function renderStats() {
            // Total
            document.getElementById('statTotal').textContent = filteredOpportunities.length;

            // Expiring soon count
            let expiringCount = 0;
            opportunities.forEach(opp => {
                const rem = getDaysRemaining(opp['Application Deadline']);
                if (rem !== null && rem >= 0 && rem <= 7) {
                    expiringCount++;
                }
            });
            document.getElementById('statExpiring').textContent = expiringCount;

            // Top Country
            const countryCounts = {};
            opportunities.forEach(opp => {
                if (opp.Country) {
                    countryCounts[opp.Country] = (countryCounts[opp.Country] || 0) + 1;
                }
            });
            let topCountry = "N/A";
            let maxCount = 0;
            for (let c in countryCounts) {
                if (countryCounts[c] > maxCount) {
                    maxCount = countryCounts[c];
                    topCountry = `${getCountryFlag(c)} ${c} (${maxCount})`;
                }
            }
            document.getElementById('statTopCountry').textContent = topCountry;
        }

        // Render Opp cards to the UI
        function renderOpportunitiesGrid() {
            const grid = document.getElementById('opportunitiesGrid');
            grid.innerHTML = '';

            if (filteredOpportunities.length === 0) {
                grid.innerHTML = `
                    <div class="col-12 text-center py-5">
                        <div class="text-muted mb-3"><i class="fa-solid fa-magnifying-glass fs-1"></i></div>
                        <h5 class="fw-bold">لم نجد أي فرص مطابقة للبحث أو الفلاتر المحددة</h5>
                        <p class="text-muted">حاول تغيير الكلمات الدليلية أو إعادة ضبط الفلاتر.</p>
                        <button class="btn btn-primary-custom btn-sm mt-2" onclick="resetFilters()">إعادة ضبط</button>
                    </div>
                `;
                return;
            }

            const startIndex = (currentPage - 1) * cardsPerPage;
            const endIndex = Math.min(startIndex + cardsPerPage, filteredOpportunities.length);
            const pageData = filteredOpportunities.slice(startIndex, endIndex);

            pageData.forEach(opp => {
                // Determine deadline badge class
                let deadlineBadgeHtml = '';
                const isOngoing = opp['Application Deadline'].includes('No deadline');
                if (isOngoing) {
                    deadlineBadgeHtml = `<span class="opp-badge badge-ongoing"><i class="fa-solid fa-calendar-check me-1"></i> تقديم مفتوح دائماً</span>`;
                } else {
                    const daysRemaining = getDaysRemaining(opp['Application Deadline']);
                    if (daysRemaining !== null) {
                        if (daysRemaining < 0) {
                            deadlineBadgeHtml = `<span class="opp-badge badge-deadline"><i class="fa-regular fa-clock me-1"></i> منتهي</span>`;
                        } else if (daysRemaining <= 7) {
                            deadlineBadgeHtml = `<span class="opp-badge badge-deadline"><i class="fa-regular fa-clock me-1"></i> ينتهي خلال ${daysRemaining} يوم</span>`;
                        } else {
                            deadlineBadgeHtml = `<span class="opp-badge badge-deadline safe"><i class="fa-regular fa-clock me-1"></i> ينتهي ${opp['Application Deadline']}</span>`;
                        }
                    } else {
                        deadlineBadgeHtml = `<span class="opp-badge badge-deadline safe"><i class="fa-regular fa-clock me-1"></i> أجل التقديم: ${opp['Application Deadline']}</span>`;
                    }
                }

                // Render Topic badges
                let topicsHtml = '';
                if (opp.Topics && opp.Topics !== 'N/A') {
                    opp.Topics.split('|').slice(0, 2).forEach(topic => {
                        topicsHtml += `<span class="opp-badge badge-topic">${topic.trim()}</span>`;
                    });
                }

                // Required document indicators
                const cvReq = opp['CV Required'] === 'Yes' ? '<span class="opp-req-badge text-danger"><i class="fa-solid fa-circle-exclamation"></i> مطلوب CV</span>' : '<span class="opp-req-badge text-success"><i class="fa-solid fa-circle-check"></i> لا يحتاج CV</span>';
                const motReq = opp['Motivation Letter Required'] === 'Yes' ? '<span class="opp-req-badge text-danger"><i class="fa-solid fa-circle-exclamation"></i> رسالة تحفيزية</span>' : '<span class="opp-req-badge text-success"><i class="fa-solid fa-circle-check"></i> لا يحتاج رسالة</span>';

                const cardCol = document.createElement('div');
                cardCol.className = 'col-md-6 col-xl-4';
                cardCol.innerHTML = `
                    <article class="opp-card">
                        <div>
                            <div class="opp-card-header">
                                <span class="opp-location">
                                    <span>${getCountryFlag(opp.Country)}</span>
                                    <span>${opp.Country}، ${opp.City}</span>
                                </span>
                            </div>
                            
                            <h2 class="opp-title" onclick="openDetails('${opp.ID}')">${opp.Title}</h2>
                            <div class="opp-org"><i class="fa-solid fa-house-chimney text-primary-50"></i> ${opp.Organisation}</div>
                            
                            <div class="opp-badges">
                                ${topicsHtml}
                                ${deadlineBadgeHtml}
                            </div>
                        </div>
                        
                        <div class="opp-footer">
                            <div class="d-flex flex-column gap-1">
                                ${cvReq}
                                ${motReq}
                            </div>
                            <button class="btn btn-sm btn-primary-custom" onclick="openDetails('${opp.ID}')">
                                التفاصيل <i class="fa-solid fa-chevron-left ms-1 fs-8"></i>
                            </button>
                        </div>
                    </article>
                `;
                grid.appendChild(cardCol);
            });
        }

        // Render dynamic pagination button controls
        function renderPagination() {
            const container = document.getElementById('paginationControls');
            container.innerHTML = '';

            const totalPages = Math.ceil(filteredOpportunities.length / cardsPerPage);
            if (totalPages <= 1) return;

            // Previous Button
            const prevBtn = document.createElement('button');
            prevBtn.className = 'page-btn';
            prevBtn.innerHTML = '<i class="fa-solid fa-angle-right"></i>';
            prevBtn.disabled = currentPage === 1;
            prevBtn.onclick = () => {
                if (currentPage > 1) {
                    currentPage--;
                    applyFiltersAndRender();
                    window.scrollTo({ top: 300, behavior: 'smooth' });
                }
            };
            container.appendChild(prevBtn);

            // Page numbers
            const startPage = Math.max(1, currentPage - 2);
            const endPage = Math.min(totalPages, startPage + 4);

            for (let i = startPage; i <= endPage; i++) {
                const btn = document.createElement('button');
                btn.className = `page-btn ${currentPage === i ? 'active' : ''}`;
                btn.textContent = i;
                btn.onclick = () => {
                    currentPage = i;
                    applyFiltersAndRender();
                    window.scrollTo({ top: 300, behavior: 'smooth' });
                };
                container.appendChild(btn);
            }

            // Next Button
            const nextBtn = document.createElement('button');
            nextBtn.className = 'page-btn';
            nextBtn.innerHTML = '<i class="fa-solid fa-angle-left"></i>';
            nextBtn.disabled = currentPage === totalPages;
            nextBtn.onclick = () => {
                if (currentPage < totalPages) {
                    currentPage++;
                    applyFiltersAndRender();
                    window.scrollTo({ top: 300, behavior: 'smooth' });
                }
            };
            container.appendChild(nextBtn);
        }

        // Open details dialog modal
        function openDetails(oppId) {
            const opp = opportunities.find(o => String(o.ID) === String(oppId));
            if (!opp) return;

            currentModalOpp = opp;

            // Title and header details
            document.getElementById('oppModalLabel').textContent = opp.Title;
            document.getElementById('modalLocation').innerHTML = `<span class="font-en">${getCountryFlag(opp.Country)} ${opp.Country}، ${opp.City}</span>`;
            document.getElementById('modalDates').textContent = `من ${opp['Start Date']} إلى ${opp['End Date']}`;
            document.getElementById('modalDeadline').textContent = opp['Application Deadline'];

            // About tab details
            document.getElementById('modalDescription').textContent = opp.Description;
            document.getElementById('modalDescription_trans').style.display = 'none';
            document.getElementById('modalDescription_trans').innerHTML = '';

            document.getElementById('modalProfile').textContent = opp['Participant Profile / Requirements'];
            document.getElementById('modalProfile_trans').style.display = 'none';
            document.getElementById('modalProfile_trans').innerHTML = '';

            // Logistics tab details
            document.getElementById('modalBoarding').textContent = opp['Accommodation & Board'];
            document.getElementById('modalBoarding_trans').style.display = 'none';
            document.getElementById('modalBoarding_trans').innerHTML = '';

            document.getElementById('modalTraining').textContent = opp['Training Provided'];
            document.getElementById('modalTraining_trans').style.display = 'none';
            document.getElementById('modalTraining_trans').innerHTML = '';

            // Organisation & Contact details
            document.getElementById('modalOrgName').textContent = opp.Organisation;
            
            const topicsContainer = document.getElementById('modalTopicsList');
            topicsContainer.innerHTML = '';
            if (opp.Topics && opp.Topics !== 'N/A') {
                opp.Topics.split('|').forEach(t => {
                    topicsContainer.innerHTML += `<span class="opp-badge badge-topic me-1">${t.trim()}</span>`;
                });
            }

            document.getElementById('modalContactName').textContent = opp['Contact Person'] || 'N/A';
            document.getElementById('modalContactEmail').textContent = opp['Contact Email'] || 'N/A';

            // Required files details
            const isCvReq = opp['CV Required'] === 'Yes';
            document.getElementById('modalCvReqIcon').className = isCvReq ? 'fa-solid fa-circle-exclamation text-danger' : 'fa-solid fa-circle-check text-success';
            document.getElementById('modalCvReqText').textContent = isCvReq ? 'مطلوبة' : 'غير مطلوبة';

            const isMotReq = opp['Motivation Letter Required'] === 'Yes';
            document.getElementById('modalMotivationReqIcon').className = isMotReq ? 'fa-solid fa-circle-exclamation text-danger' : 'fa-solid fa-circle-check text-success';
            document.getElementById('modalMotivationReqText').textContent = isMotReq ? 'مطلوبة' : 'غير مطلوبة';

            // Set up footer URLs
            document.getElementById('btnApplyLink').href = opp['Project URL'];
            
            // Revert active tab to tab 1
            const firstTabEl = document.querySelector('#modalTab button[data-bs-target="#about-pane"]');
            const tabObj = bootstrap.Tab.getOrCreateInstance(firstTabEl);
            tabObj.show();

            const myModal = new bootstrap.Modal(document.getElementById('oppModal'));
            myModal.show();
        }

        // Translation handler using free Translate widget trick
        async function triggerTranslate(field) {
            const originalDiv = document.getElementById(`modal${field}`);
            const transDiv = document.getElementById(`modal${field}_trans`);
            const btn = originalDiv.previousElementSibling.querySelector('button');

            if (transDiv.style.display === 'block') {
                transDiv.style.display = 'none';
                btn.innerHTML = '<i class="fa-solid fa-language me-1"></i> ترجمة للعربية';
                return;
            }

            if (transDiv.innerHTML !== '') {
                transDiv.style.display = 'block';
                btn.innerHTML = '<i class="fa-solid fa-arrow-rotate-left me-1"></i> استرجاع الأصل';
                return;
            }

            btn.disabled = true;
            btn.innerHTML = `<span class="loader-spinner me-1"></span> جاري الترجمة...`;

            const text = originalDiv.textContent;
            const translation = await translateText(text, 'ar');

            transDiv.innerHTML = translation;
            transDiv.style.display = 'block';
            btn.disabled = false;
            btn.innerHTML = '<i class="fa-solid fa-arrow-rotate-left me-1"></i> استرجاع الأصل';
        }

        // Paragraph translator utility
        async function translateText(text, targetLang = 'ar') {
            if (!text || text === 'N/A') return 'لا توجد تفاصيل متوفرة للترجمة.';
            try {
                const paragraphs = text.split('\\n').map(p => p.trim()).filter(p => p.length > 0);
                const translated = [];
                for (let p of paragraphs) {
                    if (p.length > 1500) {
                        const subParts = p.match(/.{{1,1500}}/g) || [p];
                        for (let sp of subParts) {
                            const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${{targetLang}}&dt=t&q=${{encodeURIComponent(sp)}}`;
                            const res = await fetch(url);
                            const data = await res.json();
                            translated.push(data[0].map(x => x[0]).join(''));
                        }
                    } else {
                        const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl=${{targetLang}}&dt=t&q=${{encodeURIComponent(p)}}`;
                        const res = await fetch(url);
                        const data = await res.json();
                        translated.push(data[0].map(x => x[0]).join(''));
                    }
                }
                return translated.join('<br><br>');
            } catch (e) {
                console.error('Translation error:', e);
                return 'حدث خطأ أثناء الترجمة. يرجى المحاولة لاحقاً.';
            }
        }

        // Copy contact email helper
        function copyEmail() {
            if (!currentModalOpp) return;
            const email = currentModalOpp['Contact Email'];
            if (!email || email === 'N/A') return;

            navigator.clipboard.writeText(email).then(() => {
                alert('تم نسخ البريد الإلكتروني بنجاح!');
            });
        }

        // Copy dynamic URL details helper
        function copyDetailsLink() {
            if (!currentModalOpp) return;
            const url = currentModalOpp['Project URL'];
            const title = currentModalOpp['Title'];
            const textToCopy = `فرصة تطوع: ${title}\\nالرابط: ${url}`;
            
            navigator.clipboard.writeText(textToCopy).then(() => {
                alert('تم نسخ تفاصيل الفرصة إلى الحافظة!');
            });
        }

        // SSE Scraper Logs updater
        function startUpdate() {
            const btn = document.getElementById('updateBtn');
            const logContainer = document.getElementById('logContainer');
            const logContent = document.getElementById('logContent');
            
            btn.disabled = true;
            btn.innerHTML = '<span class="loader-spinner me-2"></span> جاري جلب البيانات...';
            logContainer.style.display = 'block';
            logContent.innerHTML = '<div class="text-info">البدء في تشغيل السكرابر وتحميل البيانات الجديدة...</div>';
            
            // EventSource connection (SSE stream)
            const eventSource = new EventSource('http://127.0.0.1:5000/update');
            
            eventSource.onmessage = function(event) {
                const newLog = document.createElement("div");
                newLog.textContent = event.data;
                logContent.appendChild(newLog);
                
                // Auto-scroll log window
                logContainer.scrollTop = logContainer.scrollHeight;
                
                if (event.data.includes("✨ Done!") || event.data.includes("❌ No opportunities")) {
                    eventSource.close();
                    btn.innerHTML = '✅ اكتمل التحديث!';
                    btn.classList.replace('btn-success', 'btn-primary');
                    
                    // Reload dynamic details after 3 seconds to fetch new HTML structure
                    setTimeout(() => location.reload(), 3000);
                }
            };
            
            eventSource.onerror = function(err) {
                console.error("EventSource connection failure:", err);
                eventSource.close();
                btn.disabled = false;
                btn.innerHTML = '<i class="fa-solid fa-triangle-exclamation me-2"></i> فشل الاتصال';
                const errorDiv = document.createElement("div");
                errorDiv.className = "text-danger mt-2 fw-bold";
                errorDiv.innerHTML = "❌ تعذر الاتصال بخادم التحديث. يرجى تشغيل السيرفر المحلي أولاً باستخدام أمر:<br><code>python server.py</code>";
                logContent.appendChild(errorDiv);
            };
        }
    </script>
</body>
</html>
"""
    
    # Inject opportunities data and timestamp
    html_content = html_template.replace("/*__OPPORTUNITIES_JSON__*/", json_data)
    html_content = html_content.replace("__LAST_UPDATED__", datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
        
    print(f"  ✅ Saved HTML to: {filename}")


def main():
    print()

    # Step 1: Fetch ALL opportunities
    all_hits = fetch_all_opportunities()
    print(f"\n  📥 Total fetched: {len(all_hits)} opportunities")

    # Step 2: Filter for Morocco
    morocco_opps = filter_morocco_opportunities(all_hits)
    print(f"  🇲🇦 Open to Moroccan residents: {len(morocco_opps)} opportunities")
    print()

    if not morocco_opps:
        print("  ❌ No opportunities found for Morocco. Try again later.")
        return

    # Step 3: Build DataFrame
    df = build_dataframe(morocco_opps)

    # Step 4: Export to Excel, JSON, and HTML
    print(f"  📝 Exporting {len(df)} opportunities to Excel, JSON, and HTML...")
    export_to_excel(df, OUTPUT_FILE_EXCEL)
    df.to_json(OUTPUT_FILE_JSON, orient="records", force_ascii=False, indent=4)
    print(f"  ✅ Saved JSON to: {OUTPUT_FILE_JSON}")
    export_to_html(df, OUTPUT_FILE_HTML)

    # Step 5: Print summary
    print()
    print("=" * 60)
    print("  📊 Summary")
    print("=" * 60)
    print(f"  Total opportunities for Morocco: {len(df)}")
    print(f"  Countries hosting volunteers:")

    country_counts = df["Country"].value_counts().head(10)
    for country, count in country_counts.items():
        print(f"    🏳️ {country}: {count}")

    print()
    print(f"  Topics covered:")
    # Count topics
    all_topics = []
    for topics in df["Topics"]:
        if topics and topics != "N/A":
            all_topics.extend([t.strip() for t in topics.split("|")])
    if all_topics:
        topic_series = pd.Series(all_topics).value_counts().head(10)
        for topic, count in topic_series.items():
            print(f"    📌 {topic}: {count}")

    print()
    print(f"  📁 Output Excel: {OUTPUT_FILE_EXCEL}")
    print(f"  📁 Output JSON : {OUTPUT_FILE_JSON}")
    print(f"  📁 Output HTML : {OUTPUT_FILE_HTML}")
    print("  ✨ Done!")
    print()


if __name__ == "__main__":
    main()
