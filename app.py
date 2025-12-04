import streamlit as st
import pandas as pd
import os
from datetime import datetime, date, timedelta
from dateutil import parser
import numpy as np
import re
from urllib.parse import quote_plus, quote

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Insurance Marketing CRM", layout="wide")

# --- GLOBAL STYLING (DEANS & HOMER-STYLE THEME) ---
st.markdown(
    """
    <style>
    /* -------- GLOBAL APP STYLING - DEANS & HOMER LOOK (TIGHTER, CLEANER) -------- */

    /* Main page background (custom gradient) */
    html, body, .stApp, [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #f7f9fb 0%, #e5ecf5 46%, #f7f9fb 100%) !important;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui,
                     Roboto, "Helvetica Neue", Arial, sans-serif;
    }

    /* Main content container */
    .block-container {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        box-shadow: 0 4px 16px rgba(12, 25, 48, 0.10);
        margin-top: 0.8rem;
        margin-bottom: 1rem;
        border: 1.5px solid #c1cada;
        backdrop-filter: none;
    }

    /* Headings - deep D&H navy */
    h1, h2, h3 {
        color: #0f2742;
        font-weight: 650;
        letter-spacing: -0.3px;
    }

    h1 {
        font-size: 1.65rem;
        margin-bottom: 0.2rem;
    }
    h2 {
        font-size: 1.2rem;
        margin-top: 1rem;
        margin-bottom: 0.2rem;
    }
    h3 {
        font-size: 1rem;
        margin-top: 0.8rem;
        margin-bottom: 0.2rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #eef2f7;
        border-right: 1.5px solid #c2cfdd;
    }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: #0f2742;
    }

    /* Buttons */
    .stButton>button {
        background-color: #0f2742;
        color: #ffffff;
        border-radius: 4px;
        padding: 0.42rem 0.9rem;
        border: 1px solid #0f2742;
        font-size: 0.9rem;
        font-weight: 560;
        box-shadow: inset 0 -2px 0 rgba(0,0,0,0.08);
        transition: background-color 0.16s ease, transform 0.1s ease, box-shadow 0.16s ease;
    }
    .stButton>button:hover {
        background-color: #13385e;
        color: #ffffff;
        transform: translateY(-1px);
        box-shadow: inset 0 -2px 0 rgba(0,0,0,0.12);
    }
    .stButton>button:active {
        transform: translateY(0);
        box-shadow: inset 0 2px 4px rgba(0,0,0,0.14);
    }

    /* Links */
    a, a:link, a:visited {
        color: #1a6aa8;
        text-decoration: none;
    }
    a:hover {
        color: #125083;
        text-decoration: underline;
    }

    /* DataFrames */
    .stDataFrame {
        border-radius: 8px !important;
        overflow: hidden !important;
        border: 1.5px solid #c6d0de !important;
        background-color: #ffffff;
        box-shadow: inset 0 0 0 1px #e4e9f1;
    }
    .stDataFrame table {
        border-collapse: collapse !important;
    }
    .stDataFrame th, .stDataFrame td {
        border: 1px solid #d4dbe7 !important;
        padding: 6px 8px !important;
    }

    /* Labels / inputs */
    label {
        font-weight: 550 !important;
        color: #183659 !important;
        font-size: 0.9rem !important;
    }

    .stTextInput>div>div>input,
    .stTextArea textarea,
    .stSelectbox>div>div>select,
    .stDateInput>div>div>input {
        border: 1.4px solid #c4cedd;
        border-radius: 5px;
        background-color: #f9fbfe;
        color: #0f2742;
        font-size: 0.88rem;
        padding: 0.45rem 0.55rem;
    }

    .stRadio>div>label {
        color: #183659 !important;
    }

    /* Multiselect tags/options - keep blue (avoid red accents) */
    .stMultiSelect div[data-baseweb="tag"] {
        background-color: #e7f1ff !important;
        color: #0f2742 !important;
        border: 1px solid #b7d1f8 !important;
    }
    .stMultiSelect div[aria-selected="true"] {
        background-color: #e7f1ff !important;
        color: #0f2742 !important;
    }

    .stMetric {
        background-color: #f2f5fb;
        border-radius: 6px;
        padding: 0.4rem 0.6rem;
        border: 1px solid #d6dfeb;
    }

    /* ---- COLORED PANELS FOR DIFFERENT AREAS ---- */
    .panel {
        border-radius: 8px;
        padding: 0.7rem 0.9rem;
        margin-bottom: 0.8rem;
        box-shadow: 0 2px 9px rgba(15, 39, 66, 0.07);
        border: 1.5px solid #c3cbdb;
    }
    .panel-company,
    .panel-offices,
    .panel-office-left,
    .panel-office-right,
    .panel-prod,
    .panel-activity,
    .panel-contacts,
    .panel-logs {
        background-color: #f5f7fb; /* unified pale blue-gray */
    }
    </style>
    """,
    unsafe_allow_html=True
)

# --- FILE SYSTEM SETUP ---
FILES = {
    'offices': 'crm_offices.csv',
    'employees': 'crm_employees.csv',
    'agencies': 'crm_agencies.csv',
    'contacts': 'crm_contacts.csv',
    'logs': 'crm_logs.csv',
    'production': 'crm_production.csv',  # production data
    'tasks': 'crm_tasks.csv',
}

# --- CONSTANT DEFAULTS ---
OFFICE_LABELS = {
    'BRA': 'Orange County',
    'FNO': 'Fresno',
    'LAF': 'Walnut Creek',
    'LKO': 'Portland',
    'MID': 'Mid West',
    'PAS': 'Pasadena',
    'PHX': 'Phoenix',
    'RCH': 'Roseville',
    'REN': 'Reno',
    'SDO': 'San Diego',
    'SEA': 'Seattle',
    'LVS': 'Las Vegas',
    'MHL': 'Woodland Hills',
}
DEFAULT_OFFICES = list(OFFICE_LABELS.keys())
ADMIN_CODE = os.environ.get("CRM_ADMIN_CODE", "admin123")  # set env var CRM_ADMIN_CODE to override
LOGIN_USER = os.environ.get("CRM_LOGIN_USER", "admin")
LOGIN_PASSWORD = os.environ.get("CRM_LOGIN_PASSWORD", "admin123")

def format_office(code):
    """Return display-friendly office string while keeping the code for tracking."""
    name = OFFICE_LABELS.get(str(code).strip(), str(code).strip())
    return f"{code} - {name}" if name != code else str(code)

def display_office_name(code):
    """Return the friendly office name (no code prefix)."""
    return OFFICE_LABELS.get(str(code).strip(), str(code).strip())

def display_employee_name(name):
    """Strip trailing office code suffix from employee name for display."""
    display = name
    for off in st.session_state.get('offices', []):
        suff = f" {off}"
        if str(display).endswith(suff):
            display = display[:-len(suff)]
            break
    return display

# --- PRODUCT HIGHLIGHTS LOADER ---
def load_product_highlights():
    """Load product highlights from product_highlights.json if present."""
    path = os.path.join(os.getcwd(), "product_highlights.json")
    if not os.path.exists(path):
        return []
    try:
        import json
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("product_highlights", [])
    except Exception:
        return []

PRODUCT_HIGHLIGHTS = load_product_highlights()

# --- LOAD / SAVE FUNCTIONS ---
def get_new_id(df, id_col):
    """Helper to generate a unique ID."""
    if df.empty:
        return 1
    return df[id_col].max() + 1

def save_to_csv(key):
    """Saves the specific dataframe back to its CSV file."""
    if key == 'offices':
        df = pd.DataFrame({'OfficeName': st.session_state['offices']})
        df.to_csv(FILES['offices'], index=False)
    elif key == 'tasks':
        st.session_state[key].to_csv(FILES[key], index=False)
    else:
        st.session_state[key].to_csv(FILES[key], index=False)

def load_data():
    """
    Loads data from CSVs, ensuring all required columns exist and handling NaN values.
    """
    data = {}
    allowed_offices = set(OFFICE_LABELS.keys())
    offices_changed = False
    employees_changed = False
    agencies_changed = False
    contacts_changed = False
    logs_changed = False
    production_changed = False
    tasks_changed = False

    # 1. OFFICES (List)
    if os.path.exists(FILES['offices']):
        office_df = pd.read_csv(FILES['offices'])
        data['offices'] = office_df['OfficeName'].tolist() if 'OfficeName' in office_df.columns else DEFAULT_OFFICES
    else:
        data['offices'] = DEFAULT_OFFICES
        offices_changed = True
    # Keep only known office codes and ensure all defaults exist
    data['offices'] = [o for o in data['offices'] if o in OFFICE_LABELS]
    for code in DEFAULT_OFFICES:
        if code not in data['offices']:
            data['offices'].append(code)

    # 2. EMPLOYEES
    if os.path.exists(FILES['employees']):
        data['employees'] = pd.read_csv(FILES['employees'])
    else:
        # Create 5 example employees per office: Employee 1, Employee 2, ...
        rows = []
        emp_id = 1
        for office in data['offices']:
            for i in range(1, 6):
                rows.append({
                    'EmployeeID': emp_id,
                    'Name': f"Employee {i}",
                    'Office': office
                })
                emp_id += 1
        data['employees'] = pd.DataFrame(rows, columns=['EmployeeID', 'Name', 'Office'])
        employees_changed = True
    # Drop employees assigned to offices not in allowed list (e.g., legacy SBO)
    if not data['employees'].empty:
        before = len(data['employees'])
        data['employees'] = data['employees'][data['employees']['Office'].isin(allowed_offices)]
        if len(data['employees']) != before:
            employees_changed = True
    if employees_changed:
        data['employees'].to_csv(FILES['employees'], index=False)

    default_underwriter = data['employees']['Name'].iloc[0] if not data['employees'].empty else ""

    # 3. AGENCIES
    if os.path.exists(FILES['agencies']):
        data['agencies'] = pd.read_csv(FILES['agencies'])

        # Ensure columns exist
        required_cols = ['WebAddress', 'AgencyCode', 'Notes', 'PrimaryUnderwriter']
        for col in required_cols:
            if col not in data['agencies'].columns:
                if col == 'PrimaryUnderwriter':
                    data['agencies'][col] = default_underwriter
                    agencies_changed = True
                else:
                    data['agencies'][col] = ""
                    agencies_changed = True
        # Notes: no NaN
        if data['agencies']['Notes'].isna().any():
            data['agencies']['Notes'] = data['agencies']['Notes'].fillna("")
            agencies_changed = True
        # Clean WebAddress
        if 'WebAddress' in data['agencies'].columns:
            data['agencies']['WebAddress'] = data['agencies']['WebAddress'].fillna("")
            data['agencies']['WebAddress'] = data['agencies']['WebAddress'].replace(["nan", "NaN", "None"], "")

    else:
        data['agencies'] = pd.DataFrame(
            columns=['AgencyID', 'AgencyName', 'Office', 'WebAddress', 'AgencyCode', 'Notes', 'PrimaryUnderwriter']
        )
        agencies_changed = True
    # Drop agencies tied to offices not allowed
    if not data['agencies'].empty:
        before = len(data['agencies'])
        data['agencies'] = data['agencies'][data['agencies']['Office'].isin(allowed_offices)]
        if len(data['agencies']) != before:
            agencies_changed = True

    if offices_changed:
        pd.DataFrame({'OfficeName': data['offices']}).to_csv(FILES['offices'], index=False)
    if agencies_changed:
        data['agencies'].to_csv(FILES['agencies'], index=False) 

    # 4. CONTACTS
    if os.path.exists(FILES['contacts']):
        data['contacts'] = pd.read_csv(FILES['contacts'])
        if 'Notes' not in data['contacts'].columns:
            data['contacts']['Notes'] = "" 
            contacts_changed = True
        if 'Preferences' not in data['contacts'].columns:
            data['contacts']['Preferences'] = ""
            contacts_changed = True
        if 'LinkedIn' not in data['contacts'].columns:
            data['contacts']['LinkedIn'] = ""
            contacts_changed = True
        if data['contacts']['Notes'].isna().any():
            data['contacts']['Notes'] = data['contacts']['Notes'].fillna("") 
            contacts_changed = True
        # Ensure Phone column exists
        if 'Phone' not in data['contacts'].columns:
            data['contacts']['Phone'] = ""
            contacts_changed = True
        # Normalize missing values
        for col in ['Notes', 'Email', 'Phone', 'Preferences', 'LinkedIn']:
            if col in data['contacts'].columns and data['contacts'][col].isna().any():
                data['contacts'][col] = data['contacts'][col].fillna("")
                contacts_changed = True
    else:
        data['contacts'] = pd.DataFrame(
            columns=['ContactID', 'AgencyID', 'Name', 'Role', 'Email', 'Phone', 'Notes', 'Preferences', 'LinkedIn']
        )
        contacts_changed = True
    if contacts_changed:
        data['contacts'].to_csv(FILES['contacts'], index=False)

    # 5. LOGS
    if os.path.exists(FILES['logs']):
        data['logs'] = pd.read_csv(FILES['logs'])
    else:
        data['logs'] = pd.DataFrame(columns=['Date', 'EmployeeName', 'AgencyName', 'ContactName', 'Type', 'Notes', 'Office'])
        logs_changed = True

    if not data['logs'].empty:
        # Ensure Office column exists
        if 'Office' not in data['logs'].columns:
            data['logs']['Office'] = ""
            logs_changed = True
        if 'Date' in data['logs'].columns:
            data['logs']['Date'] = pd.to_datetime(data['logs']['Date'], errors='coerce')
            before_len = len(data['logs'])
            data['logs'].dropna(subset=['Date'], inplace=True)
            if len(data['logs']) != before_len:
                logs_changed = True

    if logs_changed:
        data['logs'].to_csv(FILES['logs'], index=False)

    # 6. PRODUCTION
    if os.path.exists(FILES['production']):
        prod_df = pd.read_csv(FILES['production'])
        needed_cols = ['AgencyCode', 'AgencyName', 'Office', 'Month', 'ActiveFlag', 'AllYTDWP', 'AllYTDNB', 'PYTDWP', 'PYTDNB', 'PYTotalNB']
        for c in needed_cols:
            if c not in prod_df.columns:
                prod_df[c] = "" if c in ['AgencyCode', 'AgencyName', 'Office', 'Month', 'ActiveFlag'] else 0
                production_changed = True
        # Normalize ActiveFlag to "Active"/"Inactive"
        def normalize_active(val):
            v = str(val).strip().lower()
            if v in ["y", "yes", "active", "1", "true"]:
                return "Active"
            if v in ["n", "no", "inactive", "0", "false"]:
                return "Inactive"
            return str(val).strip() if pd.notna(val) else ""
        prod_df['ActiveFlag'] = prod_df['ActiveFlag'].apply(normalize_active)
        prod_df['AllYTDWP'] = pd.to_numeric(prod_df['AllYTDWP'], errors='coerce').fillna(0.0)
        prod_df['AllYTDNB'] = pd.to_numeric(prod_df['AllYTDNB'], errors='coerce').fillna(0).astype(int)
        prod_df['PYTDWP'] = pd.to_numeric(prod_df['PYTDWP'], errors='coerce').fillna(0.0)
        prod_df['PYTDNB'] = pd.to_numeric(prod_df['PYTDNB'], errors='coerce').fillna(0).astype(int)
        prod_df['PYTotalNB'] = pd.to_numeric(prod_df['PYTotalNB'], errors='coerce').fillna(0).astype(int)
        data['production'] = prod_df[needed_cols]
    else:
        data['production'] = pd.DataFrame(
            columns=['AgencyCode', 'AgencyName', 'Office', 'Month', 'ActiveFlag', 'AllYTDWP', 'AllYTDNB', 'PYTDWP', 'PYTDNB', 'PYTotalNB']
        )
        production_changed = True
    # Drop production rows for disallowed offices
    if not data['production'].empty:
        before = len(data['production'])
        data['production'] = data['production'][data['production']['Office'].isin(allowed_offices)]
        if len(data['production']) != before:
            production_changed = True

    if production_changed:
        data['production'].to_csv(FILES['production'], index=False)

    # 7. TASKS (follow-ups / reminders)
    if os.path.exists(FILES['tasks']):
        data['tasks'] = pd.read_csv(FILES['tasks'])
        needed = ['TaskID', 'AgencyID', 'Title', 'DueDate', 'Status', 'Owner', 'Notes']
        for c in needed:
            if c not in data['tasks'].columns:
                data['tasks'][c] = "" if c not in ['TaskID', 'AgencyID'] else 0
                tasks_changed = True
        if 'DueDate' in data['tasks'].columns:
            # normalize date strings; keep as string for display/input
            if data['tasks']['DueDate'].isna().any():
                data['tasks']['DueDate'] = data['tasks']['DueDate'].fillna("")
                tasks_changed = True
        if data['tasks'].isna().any().any():
            data['tasks'] = data['tasks'].fillna("")
            tasks_changed = True
    else:
        data['tasks'] = pd.DataFrame(
            columns=['TaskID', 'AgencyID', 'Title', 'DueDate', 'Status', 'Owner', 'Notes']
        )
        tasks_changed = True

    if tasks_changed:
        data['tasks'].to_csv(FILES['tasks'], index=False)

    return data

# --- PRODUCTION PARSER ---
def parse_production_excel(uploaded_file, office, month_str):
    """
    Parse a monthly production Excel:
    - Sheet: first sheet
    - Real header row contains 'Code' in first column.
    """
    try:
        raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return pd.DataFrame()

    first_col = raw.columns[0]
    header_idx_list = raw.index[raw[first_col] == 'Code'].tolist()
    if not header_idx_list:
        st.error("Could not find a 'Code' header in the uploaded file.")
        return pd.DataFrame()

    header_idx = header_idx_list[0]

    prod = raw.iloc[header_idx:].copy()
    prod.columns = prod.iloc[0]
    prod = prod.iloc[1:]

    # Normalize column names (strip spaces and make consistent)
    prod.columns = [str(c).strip() for c in prod.columns]

    prod = prod[prod['Code'].notna() & prod['Agency'].notna()]

    # Standardize key columns
    prod = prod.rename(columns={
        'Code': 'AgencyCode',
        'Agency': 'AgencyName',
        'Active?': 'ActiveFlag',
        'Active': 'ActiveFlag',
    })

    # Coalesce YTD WP / NB across duplicate columns
    def coalesce_numeric(df, prefixes, use_last=True):
        cols = []
        for pref in prefixes:
            cols.extend([c for c in df.columns if str(c).strip().lower().startswith(pref)])
        seen = set()
        ordered = []
        for c in cols:
            if c not in seen:
                ordered.append(c)
                seen.add(c)
        if not ordered:
            return pd.Series([0] * len(df))
        if use_last:
            coalesced = df[ordered].ffill(axis=1).iloc[:, -1]
        else:
            coalesced = df[ordered].bfill(axis=1).iloc[:, 0]
        return pd.to_numeric(coalesced, errors='coerce').fillna(0)

    # Prefer the rightmost duplicate (matches spreadsheet columns I and Q)
    prod['AllYTDWP'] = coalesce_numeric(prod, ['ytd wp'], use_last=True)
    prod['AllYTDNB'] = coalesce_numeric(prod, ['ytd nb'], use_last=True)
    prod['PYTDWP'] = coalesce_numeric(prod, ['pytd wp'], use_last=True)
    prod['PYTDNB'] = coalesce_numeric(prod, ['pytd nb'], use_last=True)
    prod['PYTotalNB'] = coalesce_numeric(prod, ['py total nb'], use_last=True)

    required_final = ['AgencyCode', 'AgencyName', 'ActiveFlag', 'AllYTDWP', 'AllYTDNB', 'PYTDWP', 'PYTDNB', 'PYTotalNB']
    prod = prod[required_final].copy()

    prod['Office'] = office
    prod['Month'] = month_str

    # Normalize ActiveFlag to "Active"/"Inactive" based on yes/no style values
    def normalize_active(val):
        v = str(val).strip().lower()
        if v in ["y", "yes", "active", "1", "true"]:
            return "Active"
        if v in ["n", "no", "inactive", "0", "false"]:
            return "Inactive"
        return str(val).strip() if pd.notna(val) else ""

    prod['ActiveFlag'] = prod['ActiveFlag'].apply(normalize_active)

    prod['AgencyCode'] = prod['AgencyCode'].astype(str).str.strip()
    prod['AgencyName'] = prod['AgencyName'].astype(str).str.strip()
    prod['ActiveFlag'] = prod['ActiveFlag'].astype(str).str.strip()

    return prod

# --- INITIALIZATION ---
if 'data_initialized' not in st.session_state:
    loaded_data = load_data()
    st.session_state.update(loaded_data)
    st.session_state['data_initialized'] = True

# --- NAVIGATION STATE ---
if 'view' not in st.session_state:
    st.session_state['view'] = 'company'
if 'selected_office' not in st.session_state:
    st.session_state['selected_office'] = None
if 'selected_agency' not in st.session_state:
    st.session_state['selected_agency'] = None
if 'selected_contact' not in st.session_state:
    st.session_state['selected_contact'] = None
if 'editing_contact_id' not in st.session_state:
    st.session_state['editing_contact_id'] = None
if 'selected_employee' not in st.session_state:
    st.session_state['selected_employee'] = None
if 'admin_unlocked' not in st.session_state:
    st.session_state['admin_unlocked'] = False
if 'authenticated' not in st.session_state:
    st.session_state['authenticated'] = False

def login_gate():
    """Simple username/password gate for the whole app."""
    if st.session_state['authenticated']:
        # Provide logout in sidebar
        if st.sidebar.button("Logout"):
            st.session_state['authenticated'] = False
            st.session_state['admin_unlocked'] = False
            st.rerun()
        return True

    st.title("Login")
    with st.form("login_form"):
        user = st.text_input("Username")
        pwd = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Sign in")
        if submitted:
            if user == LOGIN_USER and pwd == LOGIN_PASSWORD:
                st.session_state['authenticated'] = True
                st.success("Login successful.")
                st.rerun()
            else:
                st.error("Invalid username or password.")
    return False

def admin_sidebar():
    """Render admin tools for managing offices, employees, and production imports."""
    st.sidebar.header("Admin")

    # Gate the admin panel behind a simple code
    if not st.session_state['admin_unlocked']:
        with st.sidebar.form("admin_unlock_form"):
            code = st.text_input("Enter admin code", type="password")
            if st.form_submit_button("Unlock"):
                if code == ADMIN_CODE:
                    st.session_state['admin_unlocked'] = True
                    st.sidebar.success("Admin unlocked for this session.")
                    st.rerun()
                else:
                    st.sidebar.error("Incorrect code.")
        return

    if st.sidebar.button("Lock admin panel"):
        st.session_state['admin_unlocked'] = False
        st.rerun()

    # Employees
    with st.sidebar.expander("Employees"):
        employees_df = st.session_state['employees']

        # Edit existing employee offices
        if not employees_df.empty:
            emp_names = sorted(employees_df['Name'].dropna().unique().tolist())
            selected_emp = st.selectbox("Select employee", emp_names, key="edit_emp_sel")
            current_offices = employees_df[employees_df['Name'] == selected_emp]['Office'].tolist()
            new_offices = st.multiselect(
                "Offices for employee",
                st.session_state['offices'],
                default=current_offices,
                key="edit_emp_offices"
            )
            if st.sidebar.button("Update Employee", key="save_emp_offices"):
                if not new_offices:
                    st.sidebar.error("Select at least one office.")
                else:
                    # Remove current rows for this employee and recreate with selected offices
                    existing_df = employees_df[employees_df['Name'] != selected_emp].copy()

                    # Preserve the first EmployeeID for this employee, then append new IDs as needed
                    emp_rows = employees_df[employees_df['Name'] == selected_emp]
                    base_id = emp_rows['EmployeeID'].iloc[0] if not emp_rows.empty else get_new_id(existing_df, 'EmployeeID')
                    max_id = int(existing_df['EmployeeID'].max()) if not existing_df.empty else base_id

                    new_rows = []
                    for idx, off in enumerate(new_offices):
                        eid = base_id if idx == 0 else max_id + idx
                        new_rows.append({
                            'EmployeeID': eid,
                            'Name': selected_emp,
                            'Office': off
                        })

                    new_emp_df = pd.DataFrame(new_rows)
                    st.session_state['employees'] = pd.concat(
                        [existing_df, new_emp_df],
                        ignore_index=True
                    )
                    save_to_csv('employees')
                    st.sidebar.success(f"Updated {selected_emp} offices to {', '.join(new_offices)}.")
                    st.rerun()
        else:
            st.sidebar.info("No employees found. Add one below.")

        # Add new employee with offices
        with st.sidebar.form("add_employee_form"):
            emp_name = st.text_input("New employee name")
            emp_offices = st.multiselect("Offices", st.session_state['offices'], key="add_emp_offices")
            add_emp = st.form_submit_button("Add Employee")
            if add_emp:
                if not emp_name:
                    st.sidebar.error("Name is required.")
                elif not emp_offices:
                    st.sidebar.error("Select at least one office.")
                else:
                    next_id = get_new_id(st.session_state['employees'], 'EmployeeID')
                    new_rows = []
                    for idx, off in enumerate(emp_offices):
                        new_rows.append({
                            'EmployeeID': next_id + idx,
                            'Name': emp_name,
                            'Office': off
                        })
                    new_emp_df = pd.DataFrame(new_rows)
                    st.session_state['employees'] = pd.concat(
                        [st.session_state['employees'], new_emp_df],
                        ignore_index=True
                    )
                    save_to_csv('employees')
                    st.sidebar.success(f"Added {emp_name} to {', '.join(emp_offices)}.")
                    st.rerun()

        if not st.session_state['employees'].empty:
            st.sidebar.caption("Manage employee offices above; deletion disabled.")

    # Production import
    with st.sidebar.expander("Production Import"):
        with st.sidebar.form("prod_import_form"):
            office_choice = st.selectbox("Office", st.session_state['offices'], key="prod_office_sel", format_func=format_office)
            today = datetime.today()
            sel_year = st.number_input("Year", min_value=2000, max_value=2100, value=today.year, step=1)
            sel_month = st.selectbox("Month", list(range(1, 13)), index=today.month - 1)
            month_str = f"{int(sel_year):04d}-{int(sel_month):02d}"
            uploaded_file = st.file_uploader("Upload Excel", type=["xls", "xlsx"])
            import_prod = st.form_submit_button("Import")
            if import_prod:
                if not uploaded_file:
                    st.sidebar.error("Please upload a file.")
                else:
                    new_prod = parse_production_excel(uploaded_file, office_choice, month_str)
                    if not new_prod.empty:
                        # Replace existing rows for this office+month
                        existing = st.session_state['production']
                        mask = ~((existing['Office'] == office_choice) & (existing['Month'] == month_str))
                        st.session_state['production'] = pd.concat(
                            [existing[mask], new_prod],
                            ignore_index=True
                        )
                        save_to_csv('production')
                        st.sidebar.success(f"Imported {len(new_prod)} production rows for {office_choice} ({month_str}); replaced any prior import for that period.")

                        # Add any new agencies from this office that are not already listed (by code), and update ActiveFlag for existing
                        agencies = st.session_state['agencies']

                        def norm_code(val):
                            return str(val).strip().upper()

                        # Normalize existing agency codes for this office
                        agencies = agencies.copy()
                        agencies['AgencyCodeNorm'] = agencies['AgencyCode'].astype(str).str.strip().str.upper()

                        # Deduplicate by AgencyCode within this import batch
                        candidate_agencies = new_prod.copy()
                        candidate_agencies['AgencyCode'] = candidate_agencies['AgencyCode'].astype(str).str.strip()
                        candidate_agencies['AgencyName'] = candidate_agencies['AgencyName'].astype(str).str.strip()
                        candidate_agencies['AgencyCodeNorm'] = candidate_agencies['AgencyCode'].str.upper()
                        candidate_agencies = candidate_agencies.drop_duplicates(subset=['AgencyCodeNorm'])

                        existing_codes_norm = set(agencies[agencies['Office'] == office_choice]['AgencyCodeNorm'].tolist())

                        to_add = candidate_agencies[
                            ~candidate_agencies['AgencyCodeNorm'].isin(existing_codes_norm) &
                            candidate_agencies['AgencyCodeNorm'].ne("")
                        ]

                        added_count = 0
                        if not to_add.empty:
                            next_agency_id = get_new_id(agencies, 'AgencyID')
                            # pick a default underwriter if available
                            office_emps = st.session_state['employees'][
                                st.session_state['employees']['Office'] == office_choice
                            ]
                            default_uw = office_emps['Name'].iloc[0] if not office_emps.empty else ""

                            new_ag_rows = []
                            for _, row in to_add.iterrows():
                                agency_code = row['AgencyCode']
                                agency_id = next_agency_id
                                next_agency_id += 1
                                new_ag_rows.append({
                                    'AgencyID': agency_id,
                                    'AgencyName': row['AgencyName'],
                                    'Office': office_choice,
                                    'WebAddress': "",
                                    'AgencyCode': agency_code,
                                    'Notes': "",
                                    'PrimaryUnderwriter': default_uw,
                                    'ActiveFlag': row.get('ActiveFlag', '')
                                })
                            if new_ag_rows:
                                added_count = len(new_ag_rows)
                                new_ag_df = pd.DataFrame(new_ag_rows)
                                st.session_state['agencies'] = pd.concat(
                                    [st.session_state['agencies'], new_ag_df],
                                    ignore_index=True
                                )
                                # Deduplicate by AgencyCode+Office to avoid double imports
                                st.session_state['agencies']['AgencyCodeNorm'] = st.session_state['agencies']['AgencyCode'].astype(str).str.strip().str.upper()
                                st.session_state['agencies'] = st.session_state['agencies'].drop_duplicates(
                                    subset=['AgencyCodeNorm', 'Office'],
                                    keep='first'
                                )
                                st.session_state['agencies'].drop(columns=['AgencyCodeNorm'], inplace=True, errors='ignore')
                                save_to_csv('agencies')

                        # Update ActiveFlag for existing agencies based on latest import
                        if not candidate_agencies.empty:
                            agencies = st.session_state['agencies'].copy()
                            agencies['AgencyCodeNorm'] = agencies['AgencyCode'].astype(str).str.strip().str.upper()
                            for _, crow in candidate_agencies.iterrows():
                                code_norm = crow['AgencyCodeNorm']
                                if code_norm in existing_codes_norm:
                                    mask = (agencies['AgencyCodeNorm'] == code_norm) & (agencies['Office'] == office_choice)
                                    agencies.loc[mask, 'ActiveFlag'] = crow.get('ActiveFlag', agencies.loc[mask, 'ActiveFlag'])
                            agencies.drop(columns=['AgencyCodeNorm'], inplace=True, errors='ignore')
                            st.session_state['agencies'] = agencies
                            save_to_csv('agencies')

                        if added_count:
                            st.sidebar.success(f"Added {added_count} new agencies for {office_choice} from import.")
                    else:
                        st.sidebar.warning("Import produced no rows. Check the file format and header names.")

    # Agency deletion (admin-only)
    with st.sidebar.expander("Delete Agency"):
        offices = st.session_state.get('offices', [])
        agencies_df = st.session_state.get('agencies', pd.DataFrame())
        tasks_df = st.session_state.get('tasks', pd.DataFrame())
        logs_df = st.session_state.get('logs', pd.DataFrame())
        contacts_df = st.session_state.get('contacts', pd.DataFrame())

        if not offices or agencies_df.empty:
            st.sidebar.info("No offices/agencies available to delete.")
        else:
            del_office = st.selectbox("Office", offices, key="del_agency_office", format_func=format_office)
            ag_in_office = agencies_df[agencies_df['Office'] == del_office]
            if ag_in_office.empty:
                st.sidebar.info("No agencies in this office.")
            else:
                # Show agency name + code for clarity
                ag_options = ag_in_office.apply(
                    lambda r: f"{r['AgencyName']} (Code: {r.get('AgencyCode','N/A')})|{r['AgencyID']}",
                    axis=1
                ).tolist()
                ag_choice = st.selectbox("Agency to delete", ag_options, key="del_agency_choice")
                if ag_choice:
                    ag_name_part, ag_id_str = ag_choice.rsplit("|", 1)
                    ag_id = int(ag_id_str)
                    if st.button("Delete agency", key="delete_agency_btn"):
                        # Remove agency
                        st.session_state['agencies'] = agencies_df[agencies_df['AgencyID'] != ag_id]
                        # Remove contacts for this agency
                        st.session_state['contacts'] = contacts_df[contacts_df['AgencyID'] != ag_id] if not contacts_df.empty else contacts_df
                        # Remove tasks for this agency
                        if not tasks_df.empty and 'AgencyID' in tasks_df.columns:
                            st.session_state['tasks'] = tasks_df[tasks_df['AgencyID'] != ag_id]
                            save_to_csv('tasks')
                        # Remove logs for this agency
                        if not logs_df.empty and 'AgencyName' in logs_df.columns:
                            ag_name = ag_in_office[ag_in_office['AgencyID'] == ag_id]['AgencyName'].iloc[0]
                            st.session_state['logs'] = logs_df[logs_df['AgencyName'] != ag_name]
                            save_to_csv('logs')
                        save_to_csv('agencies')
                        save_to_csv('contacts')
                        st.sidebar.success(f"Deleted {ag_name_part}.")
                        st.rerun()

# --- NAV HELPERS ---
def go_to_office(name):
    st.session_state['view'] = 'office'
    st.session_state['selected_office'] = name

def go_to_agency(row):
    st.session_state['view'] = 'agency'
    st.session_state['selected_agency'] = row.to_dict()

def go_to_employee(emp_name):
    st.session_state['view'] = 'employee'
    st.session_state['selected_employee'] = emp_name

def go_to_contact(contact_id, agency_id=None):
    st.session_state['view'] = 'contact'
    st.session_state['selected_contact'] = {'ContactID': contact_id, 'AgencyID': agency_id}
    if agency_id is not None:
        agencies_df = st.session_state.get('agencies', pd.DataFrame())
        match = agencies_df[agencies_df['AgencyID'] == agency_id]
        if not match.empty:
            st.session_state['selected_agency'] = match.iloc[0].to_dict()

# --- AI HELPERS (optional OpenAI integration) ---
def ai_client_available():
    """Return True if OpenAI API key and library are available."""
    if not os.environ.get("OPENAI_API_KEY"):
        return False
    try:
        import openai  # noqa: F401
    except ImportError:
        return False
    return True

def run_ai_prompt(prompt, system_prompt=None, temperature=0.2, max_tokens=600):
    """
    Fire a lightweight chat completion against OpenAI if configured.
    Returns (text, error). Does not throw.
    """
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        return None, "Set OPENAI_API_KEY to enable in-app ChatGPT."
    try:
        from openai import OpenAI
    except ImportError:
        return None, "Install the openai package to enable in-app ChatGPT."

    try:
        client = OpenAI(api_key=api_key)
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=messages,
            temperature=temperature,
            max_tokens=max_tokens,
        )
        text = resp.choices[0].message.content
        return text, None
    except Exception as e:
        return None, f"AI call failed: {e}"

# --- SIMPLE CONTACT PARSER FOR AI SUGGESTIONS ---
def parse_contact_line(text_line):
    """Attempt to pull name, role, email, phone, linkedin from a free-text line."""
    email = None
    phone = None
    linkedin = None
    line = text_line.strip()
    # linkedin
    linkedin_match = re.search(r'(https?://[^\s]*linkedin[^\s]*)', line, re.IGNORECASE)
    if linkedin_match:
        linkedin = linkedin_match.group(1).strip().rstrip(".,);")
        line = line.replace(linkedin_match.group(1), " ")
    # email
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', line)
    if email_match:
        email = email_match.group(0)
        line = line.replace(email, " ")
    # phone (grab 10+ digits)
    phone_match = re.search(r'(\+?\d[\d\-\s\(\)]{8,}\d)', line)
    if phone_match:
        raw_phone = phone_match.group(1)
        digits = re.sub(r'\D', '', raw_phone)
        if len(digits) >= 10:
            phone = digits
        line = line.replace(raw_phone, " ")
    # split on separators to guess name/role
    parts = [p.strip() for p in re.split(r'[-•|,]', line) if p.strip()]
    name = parts[0] if parts else text_line.strip()
    role = parts[1] if len(parts) > 1 else ""
    # Ignore obviously invalid emails (nan placeholder)
    if email and "nan" in email.lower():
        email = ""
    return name, role, email or "", phone or "", linkedin or ""

def parse_contact_block(text_block):
    """Parse a multi-line AI suggestion block with Name/Role/Email/LinkedIn keys."""
    name = role = email = phone = linkedin = ""
    lines = text_block.splitlines()
    for ln in lines:
        ln_stripped = ln.strip()
        if not ln_stripped:
            continue
        if ln_stripped.lower().startswith("name:") or "**name:**" in ln_stripped.lower():
            name = re.sub(r'\*\*name:\*\*', '', ln_stripped, flags=re.I).replace("Name:", "").strip()
        elif ln_stripped.lower().startswith("role:") or "**role:**" in ln_stripped.lower():
            role = re.sub(r'\*\*role:\*\*', '', ln_stripped, flags=re.I).replace("Role:", "").strip()
        elif "linkedin" in ln_stripped.lower():
            match = re.search(r'(https?://[^\s]*linkedin[^\s]*)', ln_stripped, re.I)
            linkedin = match.group(1).strip().rstrip(".,);") if match else ln_stripped.split(":")[-1].strip()
        elif "@" in ln_stripped:
            em = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', ln_stripped)
            if em:
                email = em.group(0)
        else:
            # try phone
            pm = re.search(r'(\+?\d[\d\-\s\(\)]{8,}\d)', ln_stripped)
            if pm:
                raw_phone = pm.group(1)
                digits = re.sub(r'\D', '', raw_phone)
                if len(digits) >= 10:
                    phone = digits
    if email and "nan" in email.lower():
        email = ""
    # Fallback to single-line parse if name/role missing
    if not name:
        name, role, email_f, phone_f, linkedin_f = parse_contact_line(text_block)
        email = email or email_f
        phone = phone or phone_f
        linkedin = linkedin or linkedin_f
    return name, role, email, phone, linkedin

# --- HELPER FUNCTION: GET LAST CONTACT DATE ---
def get_last_contact_status(contact_name, agency_name=None):
    """
    Returns (status_text, is_stale_over_90d).
    Treat missing/never-contacted as stale.
    """
    logs = st.session_state['logs']
    contact_logs = logs[logs['ContactName'] == contact_name].copy()
    if agency_name:
        contact_logs = contact_logs[contact_logs['AgencyName'] == agency_name]

    today = datetime.now().date()

    if contact_logs.empty:
        return "Status: never contacted", True

    try:
        valid_logs = contact_logs.dropna(subset=['Date'])
        if valid_logs.empty:
            return "Status: no valid log dates found", True

        valid_logs['Date'] = pd.to_datetime(valid_logs['Date'], errors='coerce')
        valid_logs = valid_logs.dropna(subset=['Date'])
        if valid_logs.empty:
            return "Status: no valid log dates found", True

        last_date_dt = valid_logs['Date'].max().date()
        days_diff = (today - last_date_dt).days
        last_contact_str = last_date_dt.strftime("%Y-%m-%d")

        if days_diff > 90:
            return f"Status: last contact {last_contact_str} ({days_diff} days ago, over 90 days)", True
        elif days_diff > 0:
            return f"Status: last contact {last_contact_str} ({days_diff} days ago)", False
        else:
            return f"Status: last contact {last_contact_str} (today)", False

    except Exception as e:
        return f"Status error: {e}", True

# --- VIEW: COMPANY (Level 1) ---
def view_company():
    st.title("Company Dashboard")
    logs = st.session_state['logs']
    employees = st.session_state['employees']

    # Base stats frame: one row per employee (aggregate offices to avoid duplicates)
    if employees.empty:
        stats = pd.DataFrame(columns=['Name', 'Office'])
    else:
        stats = (
            employees[['Name', 'Office']]
            .groupby('Name')
            .agg({'Office': lambda s: ", ".join(sorted(set(s)))})
            .reset_index()
        )
    # Metrics for employee activity and production
    metrics_cols = ['InPerson_30d', 'Comm_30d', 'InPerson_YTD', 'Comm_YTD', 'YTD_WP', 'YTD_NB']
    for col in metrics_cols:
        stats[col] = 0

    if not logs.empty and 'Date' in logs.columns:
        logs_dt = logs.copy()
        logs_dt['Date'] = pd.to_datetime(logs_dt['Date'], errors='coerce')
        logs_dt = logs_dt.dropna(subset=['Date'])
        if not logs_dt.empty:
            logs_dt['TypeNorm'] = logs_dt['Type'].astype(str).str.lower()
            logs_dt['DateOnly'] = logs_dt['Date'].dt.date
            today = datetime.now().date()
            cutoff_30 = today - timedelta(days=30)
            start_year = date(today.year, 1, 1)

            mask_30 = logs_dt['DateOnly'] >= cutoff_30
            mask_ytd = logs_dt['DateOnly'] >= start_year

            inperson_30 = logs_dt[mask_30 & (logs_dt['TypeNorm'] == 'in person')] \
                .groupby('EmployeeName').size().reset_index(name='InPerson_30d')
            comm_30 = logs_dt[mask_30 & (logs_dt['TypeNorm'].isin(['call', 'email']))] \
                .groupby('EmployeeName').size().reset_index(name='Comm_30d')
            inperson_ytd = logs_dt[mask_ytd & (logs_dt['TypeNorm'] == 'in person')] \
                .groupby('EmployeeName').size().reset_index(name='InPerson_YTD')
            comm_ytd = logs_dt[mask_ytd & (logs_dt['TypeNorm'].isin(['call', 'email']))] \
                .groupby('EmployeeName').size().reset_index(name='Comm_YTD')

            stats = stats.set_index('Name')

            def add_counts(df_counts, col_name):
                if df_counts.empty:
                    return
                series = df_counts.set_index('EmployeeName')[col_name]
                stats[col_name] = stats[col_name].add(series, fill_value=0)

            add_counts(inperson_30, 'InPerson_30d')
            add_counts(comm_30, 'Comm_30d')
            add_counts(inperson_ytd, 'InPerson_YTD')
            add_counts(comm_ytd, 'Comm_YTD')

            stats = stats.reset_index()
            for col in metrics_cols:
                stats[col] = stats[col].fillna(0).astype(int)

    # Add production-based metrics (YTD premium and new business count) per underwriter
    prod_df = st.session_state.get('production', pd.DataFrame())
    tasks_df = st.session_state.get('tasks', pd.DataFrame())
    agencies_df = st.session_state.get('agencies', pd.DataFrame())
    if not prod_df.empty and not agencies_df.empty:
        prod = prod_df.copy()
        prod['AgencyCode'] = prod['AgencyCode'].astype(str).str.strip()
        prod['Month_dt'] = pd.to_datetime(prod['Month'], format="%Y-%m", errors='coerce')
        prod = prod.dropna(subset=['Month_dt', 'AgencyCode'])

        if not prod.empty:
            idx = prod.groupby('AgencyCode')['Month_dt'].idxmax()
            latest_prod = prod.loc[idx].copy()

            agencies_norm = agencies_df.copy()
            agencies_norm['AgencyCode'] = agencies_norm['AgencyCode'].astype(str).str.strip()

            uw_prod = latest_prod.merge(
                agencies_norm[['AgencyCode', 'PrimaryUnderwriter']],
                on='AgencyCode',
                how='left'
            )
            uw_prod = uw_prod.dropna(subset=['PrimaryUnderwriter'])

            if not uw_prod.empty:
                by_uw = uw_prod.groupby('PrimaryUnderwriter')[['AllYTDWP', 'AllYTDNB']].sum()
                stats = stats.set_index('Name')
                # Ensure float for premium, int for count
                if 'YTD_WP' not in stats.columns:
                    stats['YTD_WP'] = 0.0
                if 'YTD_NB' not in stats.columns:
                    stats['YTD_NB'] = 0
                stats['YTD_WP'] = stats['YTD_WP'].add(by_uw['AllYTDWP'], fill_value=0)
                stats['YTD_NB'] = stats['YTD_NB'].add(by_uw['AllYTDNB'], fill_value=0).astype(int)
                stats = stats.reset_index()

    # Sort employees by most in-person meetings in last 30 days
    stats = stats.sort_values(by='InPerson_30d', ascending=False, kind='mergesort')

    st.markdown('<div class="panel panel-company">', unsafe_allow_html=True)

    # Offices navigation moved above employee summary
    st.subheader("Offices")
    # Precompute in-person counts per office
    inperson_by_office = {}
    logs = st.session_state['logs']
    if not logs.empty and 'Date' in logs.columns and 'Type' in logs.columns and 'Office' in logs.columns:
        logs_tmp = logs.copy()
        logs_tmp['TypeNorm'] = logs_tmp['Type'].astype(str).str.lower()
        inperson = logs_tmp[logs_tmp['TypeNorm'] == 'in person']
        if not inperson.empty:
            inperson_by_office = inperson.groupby('Office').size().to_dict()

    cols = st.columns(4)
    for idx, office in enumerate(st.session_state['offices']):
        label = display_office_name(office)
        inperson_ct = inperson_by_office.get(office, 0)
        with cols[idx % 4]:
            if st.button(label, key=f"btn_{office}"):
                go_to_office(office)
                st.rerun()
            st.caption(f"In Person Marketing Calls: {inperson_ct}")

    st.subheader("Employee Activity Summary")
    # Compact, square buttons for employee rows (applies broadly, but intended here)
    st.markdown(
        """
        <style>
        div.stButton > button {
            padding: 0.35rem 0.65rem;
            border-radius: 4px;
            font-size: 0.9rem;
        }
        .emp-grid-header, .emp-grid-row {
            border: 1px solid #d5d8e0;
            border-radius: 6px;
            padding: 0.35rem 0.5rem;
            margin-bottom: 0.25rem;
            background: #f4f7fa;
        }
        .emp-grid-header {
            background: #e9edf4;
            font-weight: 600;
        }
        .panel-company {
            background: rgba(255, 255, 255, 0.92);
            border-radius: 12px;
            padding: 0.75rem 1rem;
            box-shadow: 0 8px 22px rgba(0,0,0,0.08);
            border: 1px solid #d5d9e3;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    # Header with grid styling
    with st.container():
        ch1, ch2, ch3, ch4, ch5, ch6 = st.columns([3, 1, 1, 1, 1, 1], gap="small")
        with ch1:
            st.markdown('<div class="emp-grid-header">Employee</div>', unsafe_allow_html=True)
        with ch2:
            st.markdown('<div class="emp-grid-header">In person 30d</div>', unsafe_allow_html=True)
        with ch3:
            st.markdown('<div class="emp-grid-header">Emails + Calls 30d</div>', unsafe_allow_html=True)
        with ch4:
            st.markdown('<div class="emp-grid-header">In person YTD</div>', unsafe_allow_html=True)
        with ch5:
            st.markdown('<div class="emp-grid-header">Emails + Calls YTD</div>', unsafe_allow_html=True)
        with ch6:
            st.markdown('<div class="emp-grid-header">YTD NB / WP</div>', unsafe_allow_html=True)

    for _, row in stats.iterrows():
        with st.container():
            cc1, cc2, cc3, cc4, cc5, cc6 = st.columns([3, 1, 1, 1, 1, 1], gap="small")
            display_name = row['Name']
            for off in st.session_state['offices']:
                suff = f" {off}"
                if display_name.endswith(suff):
                    display_name = display_name[:-len(suff)]
                    break
            with cc1:
                emp_clicked = st.button(display_name, key=f"comp_emp_{row['Name']}")
            cc2.markdown(f"<div class='emp-grid-row'>{row['InPerson_30d']}</div>", unsafe_allow_html=True)
            cc3.markdown(f"<div class='emp-grid-row'>{row['Comm_30d']}</div>", unsafe_allow_html=True)
            cc4.markdown(f"<div class='emp-grid-row'>{row['InPerson_YTD']}</div>", unsafe_allow_html=True)
            cc5.markdown(f"<div class='emp-grid-row'>{row['Comm_YTD']}</div>", unsafe_allow_html=True)
            cc6.markdown(f"<div class='emp-grid-row'>{row.get('YTD_NB', 0)} / ${row.get('YTD_WP', 0):,.0f}</div>", unsafe_allow_html=True)
            if emp_clicked:
                go_to_employee(row['Name'])
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --- VIEW: OFFICE (Level 2) ---
def view_office():
    office = st.session_state['selected_office']
    employees = st.session_state['employees']
    logs = st.session_state['logs']
    employees_in_office = employees[employees['Office'] == office]['Name'].tolist()

    st.button("Back to Company", on_click=lambda: st.session_state.update({'view': 'company'}))
    # Title is now just the office code, e.g. "SDO"
    st.title(display_office_name(office))

    c1, c2 = st.columns([1.2, 1.8])  # widen employee column, shrink search/agency column
    with c1:
        st.markdown('<div class="panel panel-office-left">', unsafe_allow_html=True)
        st.subheader("Employees")
        # build per-employee activity for this office (in-person, calls+emails)
        office_emp_df = employees[employees['Office'] == office][['Name']].copy()

        metrics_cols = ['InPerson_30d', 'Comm_30d', 'InPerson_YTD', 'Comm_YTD']
        for col in metrics_cols:
            office_emp_df[col] = 0

        if not logs.empty and not office_emp_df.empty:
            office_logs = logs[logs['EmployeeName'].isin(office_emp_df['Name'])].copy()
            # Keep only logs explicitly tagged to this office when available
            if 'Office' in office_logs.columns:
                office_logs = office_logs[office_logs['Office'] == office]
            office_logs['Date'] = pd.to_datetime(office_logs['Date'], errors='coerce')
            office_logs = office_logs.dropna(subset=['Date'])
            if not office_logs.empty:
                office_logs['TypeNorm'] = office_logs['Type'].astype(str).str.lower()
                office_logs['DateOnly'] = office_logs['Date'].dt.date
                today = datetime.now().date()
                cutoff_30 = today - timedelta(days=30)
                start_year = date(today.year, 1, 1)

                mask_30 = office_logs['DateOnly'] >= cutoff_30
                mask_ytd = office_logs['DateOnly'] >= start_year

                inperson_30 = office_logs[mask_30 & (office_logs['TypeNorm'] == 'in person')] \
                    .groupby('EmployeeName').size().reset_index(name='InPerson_30d')
                comm_30 = office_logs[mask_30 & (office_logs['TypeNorm'].isin(['call', 'email']))] \
                    .groupby('EmployeeName').size().reset_index(name='Comm_30d')
                inperson_ytd = office_logs[mask_ytd & (office_logs['TypeNorm'] == 'in person')] \
                    .groupby('EmployeeName').size().reset_index(name='InPerson_YTD')
                comm_ytd = office_logs[mask_ytd & (office_logs['TypeNorm'].isin(['call', 'email']))] \
                    .groupby('EmployeeName').size().reset_index(name='Comm_YTD')

                office_emp_df = office_emp_df.set_index('Name')

                def add_counts(df_counts, col_name):
                    if df_counts.empty:
                        return
                    series = df_counts.set_index('EmployeeName')[col_name]
                    office_emp_df[col_name] = office_emp_df[col_name].add(series, fill_value=0)

                add_counts(inperson_30, 'InPerson_30d')
                add_counts(comm_30, 'Comm_30d')
                add_counts(inperson_ytd, 'InPerson_YTD')
                add_counts(comm_ytd, 'Comm_YTD')

                office_emp_df = office_emp_df.reset_index()
                for col in metrics_cols:
                    office_emp_df[col] = office_emp_df[col].fillna(0).astype(int)

        stats_office = office_emp_df

        # Header row like company dashboard
        h1, h2, h3, h4, h5 = st.columns([3, 1, 1, 1, 1])
        h1.markdown("")  # remove extra "Employee" label per request
        h2.markdown("**In person 30d**")
        h3.markdown("**Emails + Calls 30d**")
        h4.markdown("**In person YTD**")
        h5.markdown("**Emails + Calls YTD**")

        # Render per-employee row with the employee name acting as the navigation button
        for _, row in stats_office.iterrows():
            with st.container():
                ec1, ec2, ec3, ec4, ec5 = st.columns([3, 1, 1, 1, 1])
                # Strip any trailing office code from stored name for display
                display_name = row['Name']
                for off in st.session_state['offices']:
                    suff = f" {off}"
                    if display_name.endswith(suff):
                        display_name = display_name[:-len(suff)]
                        break
                name_clicked = ec1.button(
                    display_name,
                    key=f"emp_btn_{office}_{row['Name']}"
                )
                ec2.markdown(str(row['InPerson_30d']))
                ec3.markdown(str(row['Comm_30d']))
                ec4.markdown(str(row['InPerson_YTD']))
                ec5.markdown(str(row['Comm_YTD']))
                if name_clicked:
                    go_to_employee(row['Name'])
                    st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    agencies = st.session_state['agencies']

    with c2:
        st.markdown('<div class="panel panel-office-right">', unsafe_allow_html=True)
        # Remove the word "Agency" per request
        st.subheader("Search")
        # Narrow search bar to reduce squeeze on employee column
        with st.container():
            st.markdown(
                """
                <style>
                /* Shrink search input width in office view */
                .office-search-container > div > div {
                    max-width: 50% !important;
                }
                </style>
                """,
                unsafe_allow_html=True
            )
            search = st.text_input(
                "Search (Start typing to see results)...",
                key="office_search_input",
                help="Search within this office",
            )

        with st.expander("Add New"):
            with st.form("add_agency_form"):
                new_ag_name = st.text_input("Name*", key='new_ag_name')
                new_ag_web = st.text_input("Web Address", key='new_ag_web')
                new_ag_code = st.text_input("Code", key='new_ag_code')

                if employees_in_office:
                    new_ag_uw = st.selectbox("Primary Underwriter", options=employees_in_office, key='new_ag_uw')
                else:
                    st.warning("No employees in this office. Cannot assign Underwriter.")
                    new_ag_uw = None

                new_ag_notes = st.text_area("Notes", key='new_ag_notes')

                if st.form_submit_button("Add"):
                    if new_ag_name and new_ag_uw:
                        new_ag = pd.DataFrame({
                            'AgencyID': [get_new_id(st.session_state['agencies'], 'AgencyID')],
                            'AgencyName': [new_ag_name],
                            'Office': [office],
                            'WebAddress': [new_ag_web],
                            'AgencyCode': [new_ag_code], 
                            'Notes': [new_ag_notes],
                            'PrimaryUnderwriter': [new_ag_uw] 
                        })
                        st.session_state['agencies'] = pd.concat(
                            [st.session_state['agencies'], new_ag],
                            ignore_index=True
                        )
                        save_to_csv('agencies')
                        st.rerun()
                    elif not new_ag_name:
                        st.error("Name is required.")

        office_agencies_search = agencies[agencies['Office'] == office].copy() 

        if search:
            office_agencies_search = office_agencies_search[
                office_agencies_search['AgencyName'].str.contains(search, case=False, na=False)
            ]

        if search:
            if office_agencies_search.empty:
                st.info(f"No matches found for '{search}' in this office.")

            for idx, row in office_agencies_search.iterrows():
                with st.container(border=True):
                    ac1, ac2, ac3 = st.columns([2, 2, 1])
                    name_clicked = ac1.button(
                        row['AgencyName'],
                        key=f"op_{row['AgencyID']}_{idx}"
                    )
                    ac2.markdown(f"Code: **{row.get('AgencyCode', 'N/A')}** | UW: **{row.get('PrimaryUnderwriter', 'Unassigned')}**")
                    web_small = str(row.get('WebAddress', '') or '').strip()
                    if web_small:
                        ac2.caption(f"[{web_small}](http://{web_small})")
                    notes_content = row.get('Notes', "")
                    if notes_content:
                        ac2.markdown(f"Notes: {str(notes_content).split('.')[0]}")
                    elif row.get('WebAddress') and row['WebAddress'] != "":
                        ac2.markdown(f"Web: [{row['WebAddress']}](http://{row['WebAddress']})", unsafe_allow_html=True)
                    ac3.markdown(f"Status: **{row.get('ActiveFlag', 'Unknown')}**")

                    if name_clicked:
                        go_to_agency(row)
                        st.rerun()
        else:
            st.info("Enter a name in the search box above to quickly find agents.")
        st.markdown('</div>', unsafe_allow_html=True)

    # OFFICE LEVEL: PRODUCTION SUMMARY (OFFICE TOTALS ONLY)
    st.markdown('<div class="panel panel-prod">', unsafe_allow_html=True)
    st.subheader("Office Production Summary")

    prod_df = st.session_state.get('production', pd.DataFrame())
    if prod_df.empty:
        st.info("No production data has been uploaded yet.")
    else:
        office_prod = prod_df[prod_df['Office'] == office].copy()
        if office_prod.empty:
            st.info("No production data for this office yet.")
        else:
            office_prod['Month_dt'] = pd.to_datetime(office_prod['Month'], format="%Y-%m", errors='coerce')
            office_prod = office_prod.dropna(subset=['Month_dt'])
            if office_prod.empty:
                st.info("Production data found, but months could not be parsed.")
            else:
                latest_dt = office_prod['Month_dt'].max()
                latest_str = latest_dt.strftime("%Y-%m")
                latest_office_prod = office_prod[office_prod['Month'] == latest_str].copy()

                total_wp = latest_office_prod['AllYTDWP'].sum()
                total_nb = latest_office_prod['AllYTDNB'].sum()
                total_py_wp = latest_office_prod['PYTDWP'].sum() if 'PYTDWP' in latest_office_prod.columns else 0
                # prefer PYTDNB, fallback to PYTotalNB
                if 'PYTDNB' in latest_office_prod.columns:
                    total_py_nb = latest_office_prod['PYTDNB'].sum()
                elif 'PYTotalNB' in latest_office_prod.columns:
                    total_py_nb = latest_office_prod['PYTotalNB'].sum()
                else:
                    total_py_nb = 0
                if total_py_wp > 0:
                    wp_growth_pct = ((total_wp - total_py_wp) / total_py_wp) * 100
                    wp_growth_str = f"{wp_growth_pct:+.1f}%"
                else:
                    wp_growth_str = "PYTD unavailable"

                st.markdown(
                    f"**{latest_str}** — YTD WP: ${total_wp:,.0f} ({wp_growth_str}) | "
                    f"YTD NB: {int(total_nb):,}"
                )

                st.markdown("**Written Premium Trend (by month, office total):**")
                monthly_agg = office_prod.groupby('Month_dt')[['AllYTDWP', 'PYTDWP']].sum().sort_index()
                monthly_agg = monthly_agg.rename(columns={'AllYTDWP': 'YTD WP', 'PYTDWP': 'PYTD WP'})
                monthly_agg.index = monthly_agg.index.strftime("%m/%y")
                st.line_chart(monthly_agg)
    st.markdown('</div>', unsafe_allow_html=True)

    # OFFICE LEVEL: RECENT ACTIVITY
    st.markdown('<div class="panel panel-activity">', unsafe_allow_html=True)
    st.subheader("Recent Activity")

    agency_names_in_office = agencies[agencies['Office'] == office]['AgencyName'].tolist()

    office_logs = st.session_state['logs'][
        st.session_state['logs']['AgencyName'].isin(agency_names_in_office)
    ].copy()

    if not office_logs.empty:
        office_logs['Date'] = pd.to_datetime(office_logs['Date'], errors='coerce')
        office_logs.dropna(subset=['Date'], inplace=True)
        office_logs.sort_values(by='Date', ascending=False, inplace=True)

        st.dataframe(
            office_logs[['Date', 'AgencyName', 'ContactName', 'Type', 'EmployeeName']].head(5),
            column_config={
                "Date": st.column_config.DatetimeColumn("Date", format="YYYY-MM-DD HH:mm"),
            },
            hide_index=True,
            use_container_width=True
        )
    else:
        st.info("No activity logs recorded for accounts in this office yet.")
    st.markdown('</div>', unsafe_allow_html=True)

    # OFFICE LEVEL: ALL AGENCIES LIST
    st.markdown('<div class="panel panel-office-right">', unsafe_allow_html=True)
    st.subheader("All in this Office")

    office_agencies_all = agencies[agencies['Office'] == office]

    if not office_agencies_all.empty:
        for idx, row in office_agencies_all.iterrows():
            with st.container(border=True):
                alc1, alc2, alc3 = st.columns([2, 2, 1])

                name_clicked = alc1.button(
                    row['AgencyName'],
                    key=f"list_op_{row['AgencyID']}_{idx}"
                )
                alc2.markdown(f"Code: **{row.get('AgencyCode', 'N/A')}** | UW: **{row.get('PrimaryUnderwriter', 'Unassigned')}**")
                web_small = str(row.get('WebAddress', '') or '').strip()
                if web_small:
                    alc2.caption(f"[{web_small}](http://{web_small})")
                notes_content = row.get('Notes', "")
                if notes_content:
                    alc2.markdown(f"Notes: {str(notes_content).split('.')[0]}")
                elif row.get('WebAddress'):
                    alc2.markdown(f"Web: [{row['WebAddress']}](http://{row['WebAddress']})", unsafe_allow_html=True)

                alc3.markdown(f"Status: **{row.get('ActiveFlag', 'Unknown')}**")

                if name_clicked:
                    go_to_agency(row)
                    st.rerun()
    else:
        st.info("No accounts are currently assigned to this office.")
    st.markdown('</div>', unsafe_allow_html=True)

# --- VIEW: AGENCY (Level 3) ---
def view_agency():
    agency_dict = st.session_state['selected_agency']
    agency_id = agency_dict['AgencyID']
    office = agency_dict['Office']
    agency_web = str(agency_dict.get('WebAddress', '') or '').strip()

    employees_in_office = st.session_state['employees'][
        st.session_state['employees']['Office'] == office
    ]['Name'].tolist()

    # Contacts for this agency (so we can build the email-all button early)
    contacts = st.session_state['contacts']
    ag_contacts = contacts[contacts['AgencyID'] == agency_id].copy()
    email_list = ""
    if not ag_contacts.empty and 'Email' in ag_contacts.columns:
        emails = ag_contacts['Email'].dropna()
        emails = emails[emails != ""]
        if not emails.empty:
            email_list = ";".join(sorted(set(emails.tolist())))

    # Top bar: back button + title (left), email-all button (right)
    st.button(f"Back to {agency_dict['Office']}", on_click=lambda: go_to_office(agency_dict['Office']))
    top_left, top_right = st.columns([3, 1])
    with top_left:
        st.title(f"{agency_dict['AgencyName']}")
        if agency_web:
            st.caption(f"[{agency_web}](http://{agency_web})")
        # Primary underwriter link (clickable name)
        puw = agency_dict.get('PrimaryUnderwriter', 'Unassigned')
        if puw and puw != 'Unassigned':
            display_puw = puw
            for off in st.session_state.get('offices', []):
                suff = f" {off}"
                if display_puw.endswith(suff):
                    display_puw = display_puw[:-len(suff)]
                    break
            if st.button(f"Primary Underwriter: {display_puw}", key=f"open_uw_{agency_id}"):
                go_to_employee(puw)
                st.rerun()
        else:
            st.markdown("Primary Underwriter: Unassigned")
    with top_right:
        # Activity counts for this agency
        agency_logs_counts = st.session_state['logs'][
            st.session_state['logs']['AgencyName'] == agency_dict['AgencyName']
        ].copy()
        calls_ct = len(agency_logs_counts[agency_logs_counts['Type'].str.lower() == 'call'])
        emails_ct = len(agency_logs_counts[agency_logs_counts['Type'].str.lower() == 'email'])
        inperson_ct = len(agency_logs_counts[agency_logs_counts['Type'].str.lower() == 'in person'])
        top_right.markdown(
            f"In person: **{inperson_ct}** | Calls: **{calls_ct}** | Emails: **{emails_ct}**"
        )
        top_right.markdown("")

        if email_list:
            mailto_link = f"mailto:{email_list}"
            top_right.markdown(
                f"""	
                <a href="{mailto_link}">
                    <button style="
                        background-color:#0f2742;
                        color:white;
                        border:none;
                        padding:0.4rem 0.8rem;
                        border-radius:6px;
                        cursor:pointer;
                        font-size:0.85rem;
                    ">
                        Email all contacts
                    </button>
                </a>
                """,
                unsafe_allow_html=True
            )
        else:
            top_right.markdown("_No contact emails available_")

    # PRODUCTION SUMMARY (no history graph, now in a row)
    prod_df = st.session_state.get('production', pd.DataFrame())
    agency_code = str(agency_dict.get('AgencyCode', "")).strip()
    if not prod_df.empty and agency_code:
        prod_match = prod_df[prod_df['AgencyCode'].astype(str).str.strip() == agency_code].copy()
        if not prod_match.empty:
            prod_match['Month_dt'] = pd.to_datetime(prod_match['Month'], format="%Y-%m", errors='coerce')
            prod_match = prod_match.sort_values('Month_dt', ascending=False)

            latest = prod_match.iloc[0]
            st.markdown('<div class="panel panel-prod">', unsafe_allow_html=True)
            st.subheader("Production Summary")

            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Latest Month", latest['Month'])
            ytd_wp_latest = float(latest.get('AllYTDWP', 0))
            py_wp_latest = float(latest.get('PYTDWP', 0))
            if py_wp_latest:
                wp_growth = ((ytd_wp_latest - py_wp_latest) / py_wp_latest) * 100
                wp_growth_str = f"{wp_growth:+.1f}%"
            else:
                wp_growth_str = "N/A"
            c2.metric("YTD WP", f"${ytd_wp_latest:,.0f}")
            c3.metric("YTD WP vs PYTD", wp_growth_str)
            c4.metric("YTD NB", int(latest['AllYTDNB']))
            c5.metric("PYTD NB", int(latest.get('PYTDNB', latest.get('PYTotalNB', 0))))

            st.markdown("</div>", unsafe_allow_html=True)

    # CONTACTS
    st.markdown('<div class="panel panel-contacts">', unsafe_allow_html=True)
    header_col, search_col = st.columns([2, 1])

    header_col.subheader("Contacts")
    with header_col.expander("Add Contact"):
        with st.form("add_ct"):
            cn = st.text_input("Name*", key=f"new_ct_name_{agency_id}")
            cr = st.text_input("Role", key=f"new_ct_role_{agency_id}")
            ce = st.text_input("Email", key=f"new_ct_email_{agency_id}")
            cp = st.text_input("Phone", key=f"new_ct_phone_{agency_id}")
            cl = st.text_input("LinkedIn URL", key=f"new_ct_linkedin_{agency_id}")
            c_notes = st.text_area("Notes", key=f"new_ct_notes_{agency_id}") 
            c_pref = st.text_area("Preferences (communication, scheduling, topics)", key=f"new_ct_pref_{agency_id}")
            if st.form_submit_button("Save Contact"):
                if not cn:
                    st.error("Name is required for a contact.")
                else:
                    new_ct = pd.DataFrame({
                        'ContactID': [get_new_id(st.session_state['contacts'], 'ContactID')],
                        'AgencyID': [agency_id],
                        'Name': [cn],
                        'Role': [cr],
                        'Email': [ce],
                        'Phone': [cp],
                        'LinkedIn': [cl],
                        'Notes': [c_notes],
                        'Preferences': [c_pref]
                    })
                    st.session_state['contacts'] = pd.concat(
                        [st.session_state['contacts'], new_ct],
                        ignore_index=True
                    )
                    save_to_csv('contacts')
                    # Clear form inputs
                    for k in [
                        f"new_ct_name_{agency_id}",
                        f"new_ct_role_{agency_id}",
                        f"new_ct_email_{agency_id}",
                        f"new_ct_phone_{agency_id}",
                        f"new_ct_linkedin_{agency_id}",
                        f"new_ct_notes_{agency_id}",
                        f"new_ct_pref_{agency_id}",
                    ]:
                        st.session_state.pop(k, None)
                    st.rerun()
    with search_col:
        search_col.markdown("**Search contacts**")
        contact_search = st.text_input(
            "",
            key=f"contact_search_{agency_id}",
            placeholder="Name, email, phone, role",
            label_visibility="collapsed"
        )

    filtered_ag_contacts = ag_contacts
    if contact_search:
        term = contact_search.lower()
        def match_contact(row):
            for col in ['Name', 'Role', 'Email', 'Phone']:
                val = str(row.get(col, "")).lower()
                if term in val:
                    return True
            return False
        filtered_ag_contacts = ag_contacts[ag_contacts.apply(match_contact, axis=1)]

    for idx, row in filtered_ag_contacts.iterrows():
        contact_id = row['ContactID']
        is_editing = st.session_state['editing_contact_id'] == contact_id
        contact_name = row['Name'] 

        with st.container(border=True):
            if is_editing:
                with st.form(f"edit_contact_form_{contact_id}"):
                    e_cn = st.text_input("Name", value=row['Name'])
                    e_cr = st.text_input("Role", value=row['Role'])
                    e_ce = st.text_input("Email", value=row['Email'])
                    e_cp = st.text_input("Phone", value=row.get('Phone', ''))
                    e_cl = st.text_input("LinkedIn URL", value=row.get('LinkedIn', ''))
                    e_c_notes = st.text_area("Notes", value=row['Notes'])
                    e_c_pref = st.text_area("Preferences", value=row.get('Preferences', ''))

                    col_save, col_cancel = st.columns(2)
                    if col_save.form_submit_button("Save"):
                        contact_index_to_update = st.session_state['contacts'][
                            st.session_state['contacts']['ContactID'] == contact_id
                        ].index[0]
                        st.session_state['contacts'].loc[
                            contact_index_to_update,
                            ['Name', 'Role', 'Email', 'Phone', 'LinkedIn', 'Notes', 'Preferences']
                        ] = [e_cn, e_cr, e_ce, e_cp, e_cl, e_c_notes, e_c_pref]
                        st.session_state['editing_contact_id'] = None
                        save_to_csv('contacts')
                        st.success("Contact updated.")
                        st.rerun()
                    if col_cancel.form_submit_button("Cancel"):
                        st.session_state['editing_contact_id'] = None
                        st.rerun()
            else:
                cca, ccb, ccc, ccd = st.columns([3, 1, 1, 1])
                cca.markdown(f"**{row['Name']}** ({row['Role']})")

                last_contact_status, is_stale = get_last_contact_status(contact_name, agency_dict['AgencyName'])
                email_display = (
                    f"Email: [{row['Email']}](mailto:{row['Email']})"
                    if row.get('Email') else "Email: none"
                )
                if is_stale:
                    status_html = f"<span style='background-color:#ffe5e5;color:#9b1b1b;padding:2px 6px;border-radius:4px;'>{last_contact_status}</span>"
                else:
                    status_html = last_contact_status
                cca.markdown(f"{status_html} | {email_display}", unsafe_allow_html=True)

                if row.get('Phone') and row['Phone'] != "":
                    cca.markdown(f"Phone: {row['Phone']}")
                if row.get('LinkedIn'):
                    cca.markdown(f"[LinkedIn]({row['LinkedIn']})", unsafe_allow_html=True)

                if ccb.button("View", key=f"view_ct_{contact_id}"):
                    go_to_contact(contact_id, agency_id)
                    st.rerun()

                if ccc.button("Edit", key=f"edit_ct_{contact_id}"):
                    st.session_state['editing_contact_id'] = contact_id
                    st.rerun()

                if ccd.button("Delete", key=f"del_ct_{contact_id}"):
                    st.session_state['contacts'] = st.session_state['contacts'][
                        st.session_state['contacts']['ContactID'] != contact_id
                    ]
                    save_to_csv('contacts')
                    st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # LOG ACTIVITY
    st.markdown('<div class="panel panel-activity">', unsafe_allow_html=True)
    st.subheader("Log Activity")

    with st.form("log_act"):
        log_date = st.date_input("Date of Contact", value=datetime.today().date(), key=f"log_date_{agency_id}")

        office_employees = employees_in_office

        employees_selected = st.multiselect(
            "Employees",
            options=office_employees,
            default=[],
            key=f"log_emps_{agency_id}",
            format_func=display_employee_name
        )

        contact_options = ag_contacts['Name'].tolist()
        contacts_selected = st.multiselect(
            "Contacts",
            options=contact_options,
            default=[],
            key=f"log_contacts_{agency_id}"
        )

        type_ = st.radio("Type", ["Call", "Email", "In person"], horizontal=True, key=f"log_type_{agency_id}")
        notes = st.text_area("Notes", key=f"log_notes_{agency_id}")

        if st.form_submit_button("Log"):
            if not office_employees:
                st.error("This office has no employees set up yet.")
            elif not employees_selected:
                st.error("Please select at least one employee.")
            elif not contact_options:
                st.error("Please add at least one contact before logging activity.")
            elif not contacts_selected:
                st.error("Please select at least one contact.")
            else:
                log_datetime = datetime.combine(log_date, datetime.now().time())
                timestamp_str = log_datetime.strftime("%Y-%m-%d %H:%M:%S")

                new_rows = []
                for emp in employees_selected:
                    for ct in contacts_selected:
                        new_rows.append({
                            'Date': timestamp_str,
                            'EmployeeName': emp,
                            'AgencyName': agency_dict['AgencyName'],
                            'ContactName': ct,
                            'Type': type_,
                            'Notes': notes,
                            'Office': agency_dict['Office']
                        })

                new_log_df = pd.DataFrame(new_rows)
                st.session_state['logs'] = pd.concat(
                    [st.session_state['logs'], new_log_df],
                    ignore_index=True
                )
                save_to_csv('logs')
                st.success(f"Logged {len(new_rows)} activities.")
                for k in [f"log_emps_{agency_id}", f"log_contacts_{agency_id}", f"log_notes_{agency_id}"]:
                    st.session_state.pop(k, None)
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # AGENCY LEVEL: ALL LOGS
    st.markdown('<div class="panel panel-logs">', unsafe_allow_html=True)
    st.subheader("Activity History")

    agency_logs = st.session_state['logs'][
        st.session_state['logs']['AgencyName'] == agency_dict['AgencyName']
    ].copy()

    if not agency_logs.empty:
        agency_logs['Date'] = pd.to_datetime(agency_logs['Date'], errors='coerce')
        agency_logs.dropna(subset=['Date'], inplace=True)

        filter_col1, filter_col2, filter_col3 = st.columns([1, 1, 2])

        unique_contacts = ['All Contacts'] + agency_logs['ContactName'].unique().tolist()
        contact_filter = filter_col1.selectbox("Filter by Contact", unique_contacts)
        if contact_filter != 'All Contacts':
            agency_logs = agency_logs[agency_logs['ContactName'] == contact_filter]

        # Ensure type filter includes the new "In person" type
        type_options = agency_logs['Type'].dropna().unique().tolist()
        # Normalize capitalization for consistency
        normalized_types = []
        for t in type_options:
            t_norm = str(t).strip()
            if t_norm.lower() == "in person":
                t_norm = "In person"
            normalized_types.append(t_norm)
        if "In person" not in normalized_types:
            normalized_types.append("In person")
        unique_types = ['All Types'] + sorted(set(normalized_types))

        type_filter = filter_col2.selectbox("Filter by Type", unique_types)
        if type_filter != 'All Types':
            agency_logs = agency_logs[agency_logs['Type'].str.strip().str.lower() == type_filter.lower()]

        search_term = filter_col3.text_input("Search Employee/Notes/Keywords")
        if search_term:
            search_term = search_term.lower()
            agency_logs = agency_logs[
                agency_logs['Notes'].astype(str).str.lower().str.contains(search_term, na=False) |
                agency_logs['EmployeeName'].astype(str).str.lower().str.contains(search_term, na=False)
            ]

        agency_logs.sort_values(by='Date', ascending=False, inplace=True)

        st.dataframe(
            agency_logs[['Date', 'ContactName', 'Type', 'EmployeeName', 'Notes']],
            column_config={
                "Date": st.column_config.DatetimeColumn("Date", format="YYYY-MM-DD HH:mm"),
            },
            hide_index=True,
            use_container_width=True
        )
    else:
        st.info("No activity history recorded for this account yet.")
    st.markdown('</div>', unsafe_allow_html=True)

    # AGENCY DETAILS (moved to bottom)
    st.markdown('<div class="panel panel-company">', unsafe_allow_html=True)
    with st.expander("Edit Agency Details"):
        st.markdown(f"**Web:** {agency_dict.get('WebAddress', 'N/A')}")
        st.markdown(f"**Code:** {agency_dict.get('AgencyCode', 'N/A')}")
        st.markdown(f"**Primary Underwriter:** {agency_dict.get('PrimaryUnderwriter', 'Unassigned')}")
        st.markdown("---")
        st.markdown("**Notes:**")
        st.write(agency_dict.get('Notes', 'No notes provided.'))
        st.markdown("---")

        current_uw = agency_dict.get('PrimaryUnderwriter')

        with st.form("edit_agency_form"):
            e_name = st.text_input("Name", value=agency_dict['AgencyName'])
            e_web = st.text_input("Web Address", value=agency_dict.get('WebAddress', ''))
            e_code = st.text_input("Code", value=agency_dict.get('AgencyCode', ''))

            try:
                uw_index = employees_in_office.index(current_uw)
            except ValueError:
                uw_index = 0 if employees_in_office else 0

            e_uw = st.selectbox(
                "Primary Underwriter", 
                options=employees_in_office if employees_in_office else [current_uw],
                index=uw_index if employees_in_office else 0
            )

            e_notes = st.text_area("Notes", value=agency_dict.get('Notes', ''))

            if st.form_submit_button("Save Changes"):
                agencies = st.session_state['agencies']
                index_to_update = agencies[agencies['AgencyID'] == agency_id].index[0]

                agencies.loc[index_to_update, 'AgencyName'] = e_name
                agencies.loc[index_to_update, 'WebAddress'] = e_web
                agencies.loc[index_to_update, 'AgencyCode'] = e_code
                agencies.loc[index_to_update, 'Notes'] = e_notes
                agencies.loc[index_to_update, 'PrimaryUnderwriter'] = e_uw 

                st.session_state['selected_agency'] = agencies.loc[index_to_update].to_dict()
                save_to_csv('agencies')
                st.success("Agency details updated.")
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --- VIEW: CONTACT DETAIL (Level 4) ---
def view_contact():
    contacts_df = st.session_state.get('contacts', pd.DataFrame())
    contact_state = st.session_state.get('selected_contact') or {}
    contact_id = contact_state.get('ContactID')

    if contact_id is None:
        st.error("No contact selected.")
        return

    contact_rows = contacts_df[contacts_df['ContactID'] == contact_id]
    if contact_rows.empty:
        st.error("Contact not found.")
        return

    contact = contact_rows.iloc[0]
    agencies_df = st.session_state.get('agencies', pd.DataFrame())
    agency_rows = agencies_df[agencies_df['AgencyID'] == contact.get('AgencyID')]
    agency = agency_rows.iloc[0] if not agency_rows.empty else None
    agency_name = agency['AgencyName'] if agency is not None else "Unassigned agency"
    office = agency['Office'] if agency is not None else ""

    if agency is not None:
        st.button(f"Back to {agency_name}", on_click=lambda ag=agency: go_to_agency(ag))
    else:
        st.button("Back to Company", on_click=lambda: st.session_state.update({'view': 'company'}))

    # Overview
    st.markdown('<div class="panel panel-company">', unsafe_allow_html=True)
    top_left, top_right = st.columns([3, 1])
    with top_left:
        st.title(contact.get('Name', 'Contact'))
        role_text = contact.get('Role', '')
        st.caption(f"{role_text} @ {agency_name}" if role_text else agency_name)
        status_text, is_stale = get_last_contact_status(contact.get('Name', ''), agency_name if agency is not None else None)
        if is_stale:
            status_html = f"<span style='background-color:#ffe5e5;color:#9b1b1b;padding:2px 6px;border-radius:4px;'>{status_text}</span>"
            st.markdown(status_html, unsafe_allow_html=True)
        else:
            st.markdown(status_text)
    with top_right:
        st.markdown(f"**Office:** {format_office(office) if office else 'N/A'}")
        if contact.get('Email'):
            st.markdown(f"[Email {contact['Name']}](mailto:{contact['Email']})")
        if contact.get('Phone'):
            st.markdown(f"Phone: {contact['Phone']}")
        if contact.get('LinkedIn'):
            st.markdown(f"[LinkedIn]({contact['LinkedIn']})", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)


    # Add log scoped to this contact
    st.markdown('<div class="panel panel-activity">', unsafe_allow_html=True)
    st.subheader("Log Contact Activity")
    with st.form("contact_log_form"):
        log_date = st.date_input("Date of contact", value=datetime.today().date())
        employees_df = st.session_state.get('employees', pd.DataFrame())
        if office:
            office_emps = employees_df[employees_df['Office'] == office]['Name'].tolist()
        else:
            office_emps = employees_df['Name'].tolist()
        employees_selected = st.multiselect(
            "Employees involved",
            options=office_emps,
            default=[],
            format_func=display_employee_name
        )
        log_type = st.radio("Type", ["Call", "Email", "In person"], horizontal=True)
        log_notes = st.text_area("Notes / summary")
        if st.form_submit_button("Add activity"):
            if not office_emps:
                st.error("No employees found for this office.")
            elif not employees_selected:
                st.error("Select at least one employee.")
            else:
                log_datetime = datetime.combine(log_date, datetime.now().time())
                timestamp_str = log_datetime.strftime("%Y-%m-%d %H:%M:%S")
                new_rows = []
                contact_name = contact.get('Name', '')
                for emp in employees_selected:
                    new_rows.append({
                        'Date': timestamp_str,
                        'EmployeeName': emp,
                        'AgencyName': agency_name,
                        'ContactName': contact_name,
                        'Type': log_type,
                        'Notes': log_notes,
                        'Office': office
                    })
                new_log_df = pd.DataFrame(new_rows)
                st.session_state['logs'] = pd.concat(
                    [st.session_state['logs'], new_log_df],
                    ignore_index=True
                )
                save_to_csv('logs')
                st.success(f"Logged {len(new_rows)} {log_type.lower()} record(s) for this contact.")
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Activity history for this contact
    st.markdown('<div class="panel panel-logs">', unsafe_allow_html=True)
    st.subheader("Activity History")
    logs_df = st.session_state.get('logs', pd.DataFrame())
    if logs_df.empty:
        st.info("No activity logged for this contact yet.")
    else:
        contact_logs = logs_df[logs_df['ContactName'] == contact.get('Name', '')].copy()
        if agency_name:
            contact_logs = contact_logs[contact_logs['AgencyName'] == agency_name]
        contact_logs['Date'] = pd.to_datetime(contact_logs['Date'], errors='coerce')
        contact_logs = contact_logs.dropna(subset=['Date'])
        if contact_logs.empty:
            st.info("No activity logged for this contact yet.")
        else:
            contact_logs = contact_logs.sort_values('Date', ascending=False)
            call_count = (contact_logs['Type'].str.lower() == 'call').sum()
            email_count = (contact_logs['Type'].str.lower() == 'email').sum()
            inperson_count = (contact_logs['Type'].str.lower() == 'in person').sum()
            st.markdown(f"Calls: **{call_count}** | Emails: **{email_count}** | In person: **{inperson_count}**")

            st.dataframe(
                contact_logs[['Date', 'Type', 'EmployeeName', 'Notes']],
                column_config={
                    "Date": st.column_config.DatetimeColumn("Date", format="YYYY-MM-DD HH:mm"),
                },
                hide_index=True,
                use_container_width=True
            )

            email_history = contact_logs[contact_logs['Type'] == 'Email']
            if not email_history.empty:
                st.markdown("**Email History**")
                st.dataframe(
                    email_history[['Date', 'EmployeeName', 'Notes']],
                    column_config={
                        "Date": st.column_config.DatetimeColumn("Date", format="YYYY-MM-DD HH:mm"),
                    },
                    hide_index=True,
                    use_container_width=True
                )
    st.markdown('</div>', unsafe_allow_html=True)

    # Notes / preferences and contact details edit (moved below activity)
    st.markdown('<div class="panel panel-contacts">', unsafe_allow_html=True)
    st.subheader("General Notes & Preferences")
    with st.form("contact_details_form"):
        e_name = st.text_input("Name", value=contact.get('Name', ''))
        e_role = st.text_input("Role", value=contact.get('Role', ''))
        e_email = st.text_input("Email", value=contact.get('Email', ''))
        e_phone = st.text_input("Phone", value=contact.get('Phone', ''))
        e_linkedin = st.text_input("LinkedIn URL", value=contact.get('LinkedIn', ''))
        e_notes = st.text_area("Notes", value=contact.get('Notes', ''))
        e_pref = st.text_area("Preferences (cadence, topics, communication)", value=contact.get('Preferences', ''))
        if st.form_submit_button("Save details"):
            contact_idx = st.session_state['contacts'][
                st.session_state['contacts']['ContactID'] == contact_id
            ].index[0]
            old_name = contact.get('Name', '')
            st.session_state['contacts'].loc[
                contact_idx, ['Name', 'Role', 'Email', 'Phone', 'LinkedIn', 'Notes', 'Preferences']
            ] = [e_name, e_role, e_email, e_phone, e_linkedin, e_notes, e_pref]
            save_to_csv('contacts')

            # Keep logs aligned with updated name for this agency
            if old_name != e_name:
                logs_df = st.session_state.get('logs', pd.DataFrame())
                if not logs_df.empty:
                    mask = logs_df['ContactName'] == old_name
                    if agency_name:
                        mask = mask & (logs_df['AgencyName'] == agency_name)
                    st.session_state['logs'].loc[mask, 'ContactName'] = e_name
                    save_to_csv('logs')
            st.success("Contact details saved.")
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# --- VIEW: EMPLOYEE DETAIL (NEW) ---
def view_employee():
    emp_name = st.session_state['selected_employee']
    employees = st.session_state['employees']
    logs = st.session_state['logs']
    agencies = st.session_state['agencies']
    prod_df = st.session_state.get('production', pd.DataFrame())

    # Get employee office
    emp_row = employees[employees['Name'] == emp_name]
    emp_office = emp_row['Office'].iloc[0] if not emp_row.empty else "Unknown"

    st.button(f"Back to {emp_office}", on_click=lambda: go_to_office(emp_office))
    # Strip any trailing office code from stored name for display
    display_emp = emp_name
    for off in st.session_state['offices']:
        suff = f" {off}"
        if display_emp.endswith(suff):
            display_emp = display_emp[:-len(suff)]
            break

    st.title(f"{display_emp}")
    st.write(f"Office: **{format_office(emp_office)}**")

    # ---- PRODUCTION PANEL ----
    st.markdown('<div class="panel panel-prod">', unsafe_allow_html=True)
    st.subheader("New Business Count")

    # Agencies where this employee is PrimaryUnderwriter
    uw_agencies = agencies[agencies['PrimaryUnderwriter'] == emp_name].copy()
    if not uw_agencies.empty:
        uw_agencies = uw_agencies.sort_values('AgencyID').drop_duplicates(subset=['AgencyCode'], keep='first')
    agency_rows_for_later = None
    if not uw_agencies.empty and not prod_df.empty:
        codes = uw_agencies['AgencyCode'].astype(str).str.strip()
        prod_emp = prod_df[prod_df['AgencyCode'].astype(str).str.strip().isin(codes)].copy()

        if prod_emp.empty:
            st.info("No production data found for agencies under this employee.")
        else:
            prod_emp['Month_dt'] = pd.to_datetime(prod_emp['Month'], format="%Y-%m", errors='coerce')
            prod_emp = prod_emp.dropna(subset=['Month_dt'])
            if prod_emp.empty:
                st.info("Production rows found, but months could not be parsed.")
            else:
                # For each agency code, keep latest Month_dt
                idx = prod_emp.groupby('AgencyCode')['Month_dt'].idxmax()
                latest_prod = prod_emp.loc[idx].copy()
                latest_prod = latest_prod.drop_duplicates(subset=['AgencyCode'])

                # Normalize merge keys to strings to avoid dtype mismatches
                latest_prod['AgencyCode'] = latest_prod['AgencyCode'].astype(str).str.strip()
                latest_prod['AgencyName'] = latest_prod['AgencyName'].astype(str).str.strip()
                uw_agencies = uw_agencies.copy()
                uw_agencies['AgencyCode'] = uw_agencies['AgencyCode'].astype(str).str.strip()
                uw_agencies['AgencyName'] = uw_agencies['AgencyName'].astype(str).str.strip()

                # Merge in Office from uw_agencies, but keep a clean 'Office' column
                latest_prod = latest_prod.merge(
                    uw_agencies[['AgencyCode', 'AgencyName', 'Office']],
                    on=['AgencyCode', 'AgencyName'],
                    how='left',
                    suffixes=('', '_uw')
                )

                # Coalesce Office / Office_uw into a single 'Office' column
                if 'Office_uw' in latest_prod.columns:
                    if 'Office' in latest_prod.columns:
                        latest_prod['Office'] = latest_prod['Office_uw'].fillna(latest_prod['Office'])
                    else:
                        latest_prod['Office'] = latest_prod['Office_uw']
                    latest_prod.drop(columns=['Office_uw'], inplace=True)

                # Totals row
                total_wp = latest_prod['AllYTDWP'].sum()
                total_nb = latest_prod['AllYTDNB'].sum()

                st.markdown(
                    f"Agencies under this employee: **{len(latest_prod)}**  |  "
                    f"Total YTD Written Premium: **${total_wp:,.0f}**  |  "
                    f"Total YTD New Business Count: **{int(total_nb)}**"
                )
                st.markdown("")

                agency_rows_for_later = latest_prod.sort_values('AllYTDWP', ascending=False)
    elif uw_agencies.empty:
        st.info("This employee is not set as Primary Underwriter for any agencies.")
    else:
        st.info("No production data has been uploaded yet.")
    st.markdown('</div>', unsafe_allow_html=True)

    # ---- TASKS / FOLLOW-UPS ----
    # ---- ACTIVITY COUNTS PANEL ----
    st.markdown('<div class="panel panel-activity">', unsafe_allow_html=True)
    st.subheader("In person / Call / Email Activity")

    logs_dt = logs.copy()
    if not logs_dt.empty and 'Date' in logs_dt.columns:
        logs_dt['Date'] = pd.to_datetime(logs_dt['Date'], errors='coerce')
        logs_dt = logs_dt.dropna(subset=['Date'])
        logs_emp = logs_dt[logs_dt['EmployeeName'] == emp_name].copy()
        logs_emp['TypeNorm'] = logs_emp['Type'].astype(str).str.lower()
        today = datetime.now().date()
        cutoff_30 = today - timedelta(days=30)
        logs_emp['DateOnly'] = logs_emp['Date'].dt.date
        mask_30 = logs_emp['DateOnly'] >= cutoff_30
        mask_ytd = logs_emp['DateOnly'] >= date(today.year, 1, 1)

        inperson_30 = len(logs_emp[mask_30 & (logs_emp['TypeNorm'] == 'in person')])
        calls_30 = len(logs_emp[mask_30 & (logs_emp['TypeNorm'] == 'call')])
        emails_30 = len(logs_emp[mask_30 & (logs_emp['TypeNorm'] == 'email')])
        inperson_ytd = len(logs_emp[mask_ytd & (logs_emp['TypeNorm'] == 'in person')])
        calls_ytd = len(logs_emp[mask_ytd & (logs_emp['TypeNorm'] == 'call')])
        emails_ytd = len(logs_emp[mask_ytd & (logs_emp['TypeNorm'] == 'email')])

        c1, c2, c3, c4, c5, c6 = st.columns(6)
        c1.metric("In person 30d", inperson_30)
        c2.metric("Calls 30d", calls_30)
        c3.metric("Emails 30d", emails_30)
        c4.metric("In person YTD", inperson_ytd)
        c5.metric("Calls YTD", calls_ytd)
        c6.metric("Emails YTD", emails_ytd)
    else:
        st.info("No activity logged for this employee yet.")
    st.markdown('</div>', unsafe_allow_html=True)

    # ---- CONTACTS NEEDING FOLLOW-UP ----
    st.markdown('<div class="panel panel-activity">', unsafe_allow_html=True)
    cutoff = datetime.now().date() - timedelta(days=90)
    stale_contacts = []
    if not uw_agencies.empty:
        contacts_df = st.session_state.get('contacts', pd.DataFrame())
        logs_df = st.session_state.get('logs', pd.DataFrame())
        agency_lookup = uw_agencies.set_index('AgencyID')['AgencyName'].to_dict()
        uw_contact_rows = contacts_df[contacts_df['AgencyID'].isin(uw_agencies['AgencyID'])] if not contacts_df.empty else pd.DataFrame()
        if not uw_contact_rows.empty:
            for _, crow in uw_contact_rows.iterrows():
                ag_name = agency_lookup.get(crow['AgencyID'], "")
                if logs_df is not None and not logs_df.empty:
                    logs_match = logs_df[
                        (logs_df['AgencyName'] == ag_name) &
                        (logs_df['ContactName'] == crow['Name'])
                    ].copy()
                    logs_match['Date'] = pd.to_datetime(logs_match['Date'], errors='coerce')
                    logs_match = logs_match.dropna(subset=['Date'])
                    last_date = logs_match['Date'].max().date() if not logs_match.empty else None
                else:
                    last_date = None
                if last_date is None or last_date < cutoff:
                    days_ago = (datetime.now().date() - last_date).days if last_date else None
                    stale_contacts.append({
                        'ContactName': crow['Name'],
                        'AgencyName': ag_name,
                        'LastContact': last_date.strftime("%Y-%m-%d") if last_date else "Never",
                        'DaysSince': days_ago if days_ago is not None else "Never"
                    })

    if stale_contacts:
        st.subheader("Contacts Needing Follow-up")
        stale_df = pd.DataFrame(stale_contacts)
        st.dataframe(
            stale_df[['ContactName', 'AgencyName', 'LastContact', 'DaysSince']],
            hide_index=True,
            use_container_width=True
        )
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # ---- CALL HISTORY PANEL ----
    st.markdown('<div class="panel panel-company">', unsafe_allow_html=True)
    st.subheader("Last 20 Calls")

    if logs.empty:
        st.info("No logs recorded yet.")
    else:
        emp_logs = logs[logs['EmployeeName'] == emp_name].copy()
        emp_calls = emp_logs[emp_logs['Type'] == 'Call'].copy()
        if emp_calls.empty:
            st.info("This employee has no call activity logged yet.")
        else:
            emp_calls['Date'] = pd.to_datetime(emp_calls['Date'], errors='coerce')
            emp_calls = emp_calls.dropna(subset=['Date'])
            emp_calls = emp_calls.sort_values('Date', ascending=False).head(20)
            st.dataframe(
                emp_calls[['Date', 'AgencyName', 'ContactName', 'Notes']],
                column_config={
                    "Date": st.column_config.DatetimeColumn("Date", format="YYYY-MM-DD HH:mm"),
                },
                hide_index=True,
                use_container_width=True
            )
    st.markdown('</div>', unsafe_allow_html=True)

    # ---- AGENCIES UNDER THIS EMPLOYEE (moved to bottom) ----
    if agency_rows_for_later is not None and not agency_rows_for_later.empty:
        st.markdown('<div class="panel panel-offices">', unsafe_allow_html=True)
        st.subheader("Agencies under this employee")
        for idx_row, prow in agency_rows_for_later.iterrows():
            with st.container(border=True):
                pc1, pc2, pc3, pc4, pc5 = st.columns([3, 1, 1, 1, 1])
                name_clicked = pc1.button(
                    f"{prow.get('AgencyName', '')}  \nCode: {prow.get('AgencyCode', '')}",
                    key=f"emp_prod_open_{prow.get('AgencyCode','')}_{emp_name}_{idx_row}"
                )
                pc2.markdown(f"Office: {format_office(prow.get('Office', ''))}")
                pc3.markdown(f"YTD WP: ${prow.get('AllYTDWP', 0):,.0f}")
                pc4.markdown(f"YTD NB: {int(prow.get('AllYTDNB', 0))}")
                pc5.markdown(f"Active: {prow.get('ActiveFlag', '')}")
                if name_clicked:
                    agencies_match = agencies[
                        agencies['AgencyCode'].astype(str).str.strip() == str(prow.get('AgencyCode', '')).strip()
                    ]
                    if not agencies_match.empty:
                        go_to_agency(agencies_match.iloc[0])
                        st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

# --- ROUTER ---
if login_gate():
    admin_sidebar()
    if st.session_state['view'] == 'company':
        view_company()
    elif st.session_state['view'] == 'office':
        view_office()
    elif st.session_state.get('view') == 'agency':
        view_agency()
    elif st.session_state.get('view') == 'contact':
        view_contact()
    elif st.session_state.get('view') == 'employee':
        view_employee()
    else:
        st.error(f"Unknown view: {st.session_state.get('view')}")
