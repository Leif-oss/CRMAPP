import streamlit as st
import pandas as pd
import os
from datetime import datetime, date, timedelta
from dateutil import parser
import numpy as np
from urllib.parse import quote_plus

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Insurance Marketing CRM", layout="wide")

# --- GLOBAL STYLING (DEANS & HOMER-STYLE THEME) ---
st.markdown(
    """
    <style>
    /* -------- GLOBAL APP STYLING - DEANS & HOMER LOOK -------- */

    /* Main page background */
    .stApp {
        background: linear-gradient(135deg, #f4f7fb 0%, #e4edf7 50%, #dde6f2 100%);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", system-ui,
                     Roboto, "Helvetica Neue", Arial, sans-serif;
    }

    /* Main content container */
    .block-container {
        background-color: rgba(255, 255, 255, 0.9);
        border-radius: 14px;
        padding: 1.8rem 2.2rem;
        box-shadow: 0 8px 26px rgba(11, 34, 66, 0.08);
        margin-top: 1.2rem;
        margin-bottom: 1.8rem;
        border: 1px solid #d0d9e6;
    }

    /* Headings - deep D&H navy */
    h1, h2, h3 {
        color: #0f2742;
        font-weight: 650;
        letter-spacing: -0.3px;
    }

    h1 {
        font-size: 1.8rem;
        margin-bottom: 0.4rem;
    }
    h2 {
        font-size: 1.3rem;
        margin-top: 1.3rem;
    }
    h3 {
        font-size: 1.1rem;
        margin-top: 1.0rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #e6edf5;
        border-right: 1px solid #c2cfdd;
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
        border-radius: 6px;
        padding: 0.4rem 0.9rem;
        border: none;
        font-size: 0.9rem;
        font-weight: 500;
        transition: background-color 0.18s ease, transform 0.12s ease;
    }
    .stButton>button:hover {
        background-color: #1e3c63;
        color: #ffffff;
        transform: translateY(-1px);
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
        border: 1px solid #d0d9e6 !important;
        background-color: #ffffff;
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
        border: 1px solid #c4cedd;
        border-radius: 6px;
        background-color: #f8fafc;
        color: #0f2742;
        font-size: 0.9rem;
    }

    .stRadio>div>label {
        color: #183659 !important;
    }

    .stMetric {
        background-color: #f2f5fb;
        border-radius: 8px;
        padding: 0.5rem 0.75rem;
    }

    /* ---- COLORED PANELS FOR DIFFERENT AREAS ---- */
    .panel {
        border-radius: 10px;
        padding: 0.9rem 1.1rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 4px 14px rgba(15, 39, 66, 0.06);
        border: 1px solid #d5dfec;
    }
    .panel-company {
        background-color: #f4fbff; /* light blue */
    }
    .panel-offices {
        background-color: #fdf7ec; /* light warm */
    }
    .panel-office-left {
        background-color: #f4fbff; /* light blue for employees */
    }
    .panel-office-right {
        background-color: #f6f2ff; /* soft purple for search/list */
    }
    .panel-prod {
        background-color: #f0f7f4; /* soft green tint for production */
    }
    .panel-activity {
        background-color: #f6f8fc; /* light gray-blue for activity logs */
    }
    .panel-contacts {
        background-color: #fff7fb; /* light pink for contacts */
    }
    .panel-logs {
        background-color: #f4fffb; /* aqua tint for activity history/logs */
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
}

# --- CONSTANT DEFAULTS ---
DEFAULT_OFFICES = ['BRA', 'FNO', 'LAF', 'LKO', 'MID', 'PAS', 'PHX', 'RCH', 'REN', 'SBO', 'SDO', 'SEA', 'LVS']
ADMIN_CODE = os.environ.get("CRM_ADMIN_CODE", "admin123")  # set env var CRM_ADMIN_CODE to override
LOGIN_USER = os.environ.get("CRM_LOGIN_USER", "admin")
LOGIN_PASSWORD = os.environ.get("CRM_LOGIN_PASSWORD", "admin123")

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
    else:
        st.session_state[key].to_csv(FILES[key], index=False)

def load_data():
    """
    Loads data from CSVs, ensuring all required columns exist and handling NaN values.
    """
    data = {}
    offices_changed = False
    employees_changed = False
    agencies_changed = False
    contacts_changed = False
    logs_changed = False
    production_changed = False

    # 1. OFFICES (List)
    if os.path.exists(FILES['offices']):
        office_df = pd.read_csv(FILES['offices'])
        data['offices'] = office_df['OfficeName'].tolist() if 'OfficeName' in office_df.columns else DEFAULT_OFFICES
    else:
        data['offices'] = DEFAULT_OFFICES
        offices_changed = True

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

    else:
        data['agencies'] = pd.DataFrame(
            columns=['AgencyID', 'AgencyName', 'Office', 'WebAddress', 'AgencyCode', 'Notes', 'PrimaryUnderwriter']
        )
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
        if data['contacts']['Notes'].isna().any():
            data['contacts']['Notes'] = data['contacts']['Notes'].fillna("") 
            contacts_changed = True
        # Ensure Phone column exists
        if 'Phone' not in data['contacts'].columns:
            data['contacts']['Phone'] = ""
            contacts_changed = True
        # Normalize missing values
        for col in ['Notes', 'Email', 'Phone']:
            if col in data['contacts'].columns and data['contacts'][col].isna().any():
                data['contacts'][col] = data['contacts'][col].fillna("")
                contacts_changed = True
    else:
        data['contacts'] = pd.DataFrame(
            columns=['ContactID', 'AgencyID', 'Name', 'Role', 'Email', 'Phone', 'Notes']
        )
        contacts_changed = True
    if contacts_changed:
        data['contacts'].to_csv(FILES['contacts'], index=False)

    # 5. LOGS
    if os.path.exists(FILES['logs']):
        data['logs'] = pd.read_csv(FILES['logs'])
    else:
        data['logs'] = pd.DataFrame(columns=['Date', 'EmployeeName', 'AgencyName', 'ContactName', 'Type', 'Notes'])
        logs_changed = True

    if not data['logs'].empty and 'Date' in data['logs'].columns:
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
        needed_cols = ['AgencyCode', 'AgencyName', 'Office', 'Month', 'ActiveFlag', 'AllYTDWP', 'AllYTDNB']
        for c in needed_cols:
            if c not in prod_df.columns:
                prod_df[c] = "" if c in ['AgencyCode', 'AgencyName', 'Office', 'Month', 'ActiveFlag'] else 0
                production_changed = True
        prod_df['AllYTDWP'] = pd.to_numeric(prod_df['AllYTDWP'], errors='coerce').fillna(0.0)
        prod_df['AllYTDNB'] = pd.to_numeric(prod_df['AllYTDNB'], errors='coerce').fillna(0).astype(int)
        data['production'] = prod_df[needed_cols]
    else:
        data['production'] = pd.DataFrame(
            columns=['AgencyCode', 'AgencyName', 'Office', 'Month', 'ActiveFlag', 'AllYTDWP', 'AllYTDNB']
        )
        production_changed = True

    if production_changed:
        data['production'].to_csv(FILES['production'], index=False)

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

    # Allow common variants of column labels
    col_map = {
        'code': 'AgencyCode',
        'agency': 'AgencyName',
        'active?': 'ActiveFlag',
        'active': 'ActiveFlag',
        'ytd wp': 'AllYTDWP',
        'ytd wp  ': 'AllYTDWP',
        'ytd wp #': 'AllYTDWP',
        'ytd nb #': 'AllYTDNB',
        'ytd nb  ': 'AllYTDNB',
    }

    normalized_cols = {}
    for c in prod.columns:
        key = str(c).strip().lower()
        if key in col_map:
            normalized_cols[c] = col_map[key]

    prod = prod.rename(columns=normalized_cols)

    # Remove duplicate columns that can appear in some exports
    prod = prod.loc[:, ~prod.columns.duplicated()]

    required_final = ['AgencyCode', 'AgencyName', 'ActiveFlag', 'AllYTDWP', 'AllYTDNB']
    missing_final = [c for c in required_final if c not in prod.columns]
    if missing_final:
        st.error(f"Missing expected columns in the file: {missing_final}")
        return pd.DataFrame()

    prod = prod[required_final].copy()

    prod['Office'] = office
    prod['Month'] = month_str
    # Some files can include duplicate column names; ensure we operate on Series
    def as_series(col):
        return col.iloc[:, 0] if isinstance(col, pd.DataFrame) else col

    prod['AllYTDWP'] = pd.to_numeric(as_series(prod['AllYTDWP']), errors='coerce').fillna(0.0)
    prod['AllYTDNB'] = pd.to_numeric(as_series(prod['AllYTDNB']), errors='coerce').fillna(0).astype(int)
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

    # Offices
    with st.sidebar.expander("Offices"):
        with st.sidebar.form("add_office_form"):
            new_office = st.text_input("New office code")
            add_office = st.form_submit_button("Add Office")
            if add_office and new_office:
                if new_office in st.session_state['offices']:
                    st.sidebar.warning("Office already exists.")
                else:
                    st.session_state['offices'].append(new_office)
                    save_to_csv('offices')
                    st.sidebar.success(f"Added office {new_office}.")

        if st.session_state['offices']:
            del_office = st.selectbox("Delete office", st.session_state['offices'], key="del_office_sel")
            if st.sidebar.button("Delete Selected Office"):
                st.session_state['offices'] = [o for o in st.session_state['offices'] if o != del_office]
                save_to_csv('offices')
                st.sidebar.success(f"Deleted office {del_office}.")

    # Employees
    with st.sidebar.expander("Employees"):
        with st.sidebar.form("add_employee_form"):
            emp_name = st.text_input("Name")
            emp_office = st.selectbox("Office", st.session_state['offices'])
            add_emp = st.form_submit_button("Add Employee")
            if add_emp and emp_name:
                new_emp = pd.DataFrame({
                    'EmployeeID': [get_new_id(st.session_state['employees'], 'EmployeeID')],
                    'Name': [emp_name],
                    'Office': [emp_office]
                })
                st.session_state['employees'] = pd.concat(
                    [st.session_state['employees'], new_emp],
                    ignore_index=True
                )
                save_to_csv('employees')
                st.sidebar.success(f"Added employee {emp_name}.")

        if not st.session_state['employees'].empty:
            del_emp = st.selectbox("Delete employee", st.session_state['employees']['Name'], key="del_emp_sel")
            if st.sidebar.button("Delete Selected Employee"):
                st.session_state['employees'] = st.session_state['employees'][
                    st.session_state['employees']['Name'] != del_emp
                ]
                save_to_csv('employees')
                st.sidebar.success(f"Deleted employee {del_emp}.")

    # Production import
    with st.sidebar.expander("Production Import"):
        with st.sidebar.form("prod_import_form"):
            office_choice = st.selectbox("Office", st.session_state['offices'], key="prod_office_sel")
            month_input = st.date_input("Month (any day)", value=datetime.today())
            month_str = month_input.strftime("%Y-%m")
            uploaded_file = st.file_uploader("Upload Excel", type=["xls", "xlsx"])
            import_prod = st.form_submit_button("Import")
            if import_prod:
                if not uploaded_file:
                    st.sidebar.error("Please upload a file.")
                else:
                    new_prod = parse_production_excel(uploaded_file, office_choice, month_str)
                    if not new_prod.empty:
                        st.session_state['production'] = pd.concat(
                            [st.session_state['production'], new_prod],
                            ignore_index=True
                        )
                        save_to_csv('production')
                        st.sidebar.success(f"Imported {len(new_prod)} production rows for {office_choice} ({month_str}).")

                        # Add any new agencies from this office that are not already listed
                        agencies = st.session_state['agencies']
                        existing_codes = agencies[agencies['Office'] == office_choice]['AgencyCode'] \
                            .astype(str).str.strip().tolist()

                        # Deduplicate by AgencyCode within this import batch
                        candidate_agencies = new_prod.copy()
                        candidate_agencies['AgencyCode'] = candidate_agencies['AgencyCode'].astype(str).str.strip()
                        candidate_agencies['AgencyName'] = candidate_agencies['AgencyName'].astype(str).str.strip()
                        candidate_agencies = candidate_agencies.drop_duplicates(subset=['AgencyCode'])

                        missing_mask = ~candidate_agencies['AgencyCode'].isin(existing_codes) & \
                                       candidate_agencies['AgencyCode'].ne("")
                        to_add = candidate_agencies[missing_mask]

                        if not to_add.empty:
                            next_agency_id = get_new_id(agencies, 'AgencyID')
                            # pick a default underwriter if available
                            office_emps = st.session_state['employees'][
                                st.session_state['employees']['Office'] == office_choice
                            ]
                            default_uw = office_emps['Name'].iloc[0] if not office_emps.empty else ""

                            new_ag_rows = []
                            for _, row in to_add.iterrows():
                                agency_id = next_agency_id
                                next_agency_id += 1
                                new_ag_rows.append({
                                    'AgencyID': agency_id,
                                    'AgencyName': row['AgencyName'],
                                    'Office': office_choice,
                                    'WebAddress': "",
                                    'AgencyCode': row['AgencyCode'],
                                    'Notes': "",
                                    'PrimaryUnderwriter': default_uw
                                })

                            new_ag_df = pd.DataFrame(new_ag_rows)
                            st.session_state['agencies'] = pd.concat(
                                [st.session_state['agencies'], new_ag_df],
                                ignore_index=True
                            )
                            save_to_csv('agencies')
                            st.sidebar.success(f"Added {len(new_ag_df)} new agencies for {office_choice} from import.")
                    else:
                        st.sidebar.warning("Import produced no rows. Check the file format and header names.")

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

# --- HELPER FUNCTION: GET LAST CONTACT DATE ---
def get_last_contact_status(contact_name):
    logs = st.session_state['logs']
    contact_logs = logs[logs['ContactName'] == contact_name].copy()

    today = datetime.now().date()

    if contact_logs.empty:
        return "Status: never contacted"

    try:
        valid_logs = contact_logs.dropna(subset=['Date'])
        if valid_logs.empty:
            return "Status: no valid log dates found"

        valid_logs['Date'] = pd.to_datetime(valid_logs['Date'], errors='coerce')
        valid_logs = valid_logs.dropna(subset=['Date'])
        if valid_logs.empty:
            return "Status: no valid log dates found"

        last_date_dt = valid_logs['Date'].max().date()
        days_diff = (today - last_date_dt).days
        last_contact_str = last_date_dt.strftime("%Y-%m-%d")

        if days_diff > 90:
            return f"Status: last contact {last_contact_str} ({days_diff} days ago, over 90 days)"
        elif days_diff > 0:
            return f"Status: last contact {last_contact_str} ({days_diff} days ago)"
        else:
            return f"Status: last contact {last_contact_str} (today)"

    except Exception as e:
        return f"Status error: {e}"

# --- VIEW: COMPANY (Level 1) ---
def view_company():
    st.title("Company Dashboard")
    logs = st.session_state['logs']
    employees = st.session_state['employees']

    # Base stats frame: one row per employee
    stats = employees[['Name', 'Office']].copy()
    metrics_cols = ['Calls_30d', 'Emails_30d', 'Calls_YTD', 'Emails_YTD', 'YTD_WP', 'YTD_NB']
    for col in metrics_cols:
        stats[col] = 0

    if not logs.empty and 'Date' in logs.columns:
        logs_dt = logs.copy()
        logs_dt['Date'] = pd.to_datetime(logs_dt['Date'], errors='coerce')
        logs_dt = logs_dt.dropna(subset=['Date'])
        if not logs_dt.empty:
            logs_dt['DateOnly'] = logs_dt['Date'].dt.date
            today = datetime.now().date()
            cutoff_30 = today - timedelta(days=30)
            start_year = date(today.year, 1, 1)

            mask_30 = logs_dt['DateOnly'] >= cutoff_30
            mask_ytd = logs_dt['DateOnly'] >= start_year

            calls_30 = logs_dt[mask_30 & (logs_dt['Type'] == 'Call')] \
                .groupby('EmployeeName').size().reset_index(name='Calls_30d')
            emails_30 = logs_dt[mask_30 & (logs_dt['Type'] == 'Email')] \
                .groupby('EmployeeName').size().reset_index(name='Emails_30d')
            calls_ytd = logs_dt[mask_ytd & (logs_dt['Type'] == 'Call')] \
                .groupby('EmployeeName').size().reset_index(name='Calls_YTD')
            emails_ytd = logs_dt[mask_ytd & (logs_dt['Type'] == 'Email')] \
                .groupby('EmployeeName').size().reset_index(name='Emails_YTD')

            stats = stats.set_index('Name')

            def add_counts(df_counts, col_name):
                if df_counts.empty:
                    return
                series = df_counts.set_index('EmployeeName')[col_name]
                stats[col_name] = stats[col_name].add(series, fill_value=0)

            add_counts(calls_30, 'Calls_30d')
            add_counts(emails_30, 'Emails_30d')
            add_counts(calls_ytd, 'Calls_YTD')
            add_counts(emails_ytd, 'Emails_YTD')

            stats = stats.reset_index()
            for col in metrics_cols:
                stats[col] = stats[col].fillna(0).astype(int)

    # Add production-based metrics (YTD premium and new business count) per underwriter
    prod_df = st.session_state.get('production', pd.DataFrame())
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

    # Sort employees by most calls in last 30 days
    stats = stats.sort_values(by='Calls_30d', ascending=False, kind='mergesort')

    st.markdown('<div class="panel panel-company">', unsafe_allow_html=True)
    st.subheader("Employee Activity Summary")
    st.dataframe(
        stats[['Name', 'Office', 'Calls_30d', 'Emails_30d', 'Calls_YTD', 'Emails_YTD', 'YTD_NB', 'YTD_WP']],
        use_container_width=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="panel panel-offices">', unsafe_allow_html=True)
    st.subheader("Offices")
    cols = st.columns(4)
    for idx, office in enumerate(st.session_state['offices']):
        if cols[idx % 4].button(f"{office}", key=f"btn_{office}"):
            go_to_office(office)
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
    st.title(f"{office}")

    c1, c2 = st.columns([1.2, 1.8])  # widen employee column, shrink search/agency column
    with c1:
        st.markdown('<div class="panel panel-office-left">', unsafe_allow_html=True)
        st.subheader("Employees")
        # build per-employee calls/emails for this office
        office_emp_df = employees[employees['Office'] == office][['Name']].copy()

        metrics_cols = ['Calls_30d', 'Emails_30d', 'Calls_YTD', 'Emails_YTD']
        for col in metrics_cols:
            office_emp_df[col] = 0

        if not logs.empty and not office_emp_df.empty:
            office_logs = logs[logs['EmployeeName'].isin(office_emp_df['Name'])].copy()
            office_logs['Date'] = pd.to_datetime(office_logs['Date'], errors='coerce')
            office_logs = office_logs.dropna(subset=['Date'])
            if not office_logs.empty:
                office_logs['DateOnly'] = office_logs['Date'].dt.date
                today = datetime.now().date()
                cutoff_30 = today - timedelta(days=30)
                start_year = date(today.year, 1, 1)

                mask_30 = office_logs['DateOnly'] >= cutoff_30
                mask_ytd = office_logs['DateOnly'] >= start_year

                calls_30 = office_logs[mask_30 & (office_logs['Type'] == 'Call')] \
                    .groupby('EmployeeName').size().reset_index(name='Calls_30d')
                emails_30 = office_logs[mask_30 & (office_logs['Type'] == 'Email')] \
                    .groupby('EmployeeName').size().reset_index(name='Emails_30d')
                calls_ytd = office_logs[mask_ytd & (office_logs['Type'] == 'Call')] \
                    .groupby('EmployeeName').size().reset_index(name='Calls_YTD')
                emails_ytd = office_logs[mask_ytd & (office_logs['Type'] == 'Email')] \
                    .groupby('EmployeeName').size().reset_index(name='Emails_YTD')

                office_emp_df = office_emp_df.set_index('Name')

                def add_counts(df_counts, col_name):
                    if df_counts.empty:
                        return
                    series = df_counts.set_index('EmployeeName')[col_name]
                    office_emp_df[col_name] = office_emp_df[col_name].add(series, fill_value=0)

                add_counts(calls_30, 'Calls_30d')
                add_counts(emails_30, 'Emails_30d')
                add_counts(calls_ytd, 'Calls_YTD')
                add_counts(emails_ytd, 'Emails_YTD')

                office_emp_df = office_emp_df.reset_index()
                for col in metrics_cols:
                    office_emp_df[col] = office_emp_df[col].fillna(0).astype(int)

        stats_office = office_emp_df

        # Header row like company dashboard
        h1, h2, h3, h4, h5 = st.columns([3, 1, 1, 1, 1])
        h1.markdown("")  # remove extra "Employee" label per request
        h2.markdown("**Calls 30d**")
        h3.markdown("**Emails 30d**")
        h4.markdown("**Calls YTD**")
        h5.markdown("**Emails YTD**")

        # Render per-employee row with a "View" button to open employee screen
        for _, row in stats_office.iterrows():
            with st.container():
                ec1, ec2, ec3, ec4, ec5 = st.columns([3, 1, 1, 1, 1])
                # Strip trailing office code from display only
                display_name = row['Name']
                suffix = f" {office}"
                if display_name.endswith(suffix):
                    display_name = display_name[:-len(suffix)]
                ec1.markdown(f"**{display_name}**")
                ec2.markdown(str(row['Calls_30d']))
                ec3.markdown(str(row['Emails_30d']))
                ec4.markdown(str(row['Calls_YTD']))
                ec5.markdown(str(row['Emails_YTD']))
                if ec5.button("View", key=f"emp_view_{office}_{row['Name']}"):
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
                    ac1, ac2, ac3 = st.columns([3, 1, 1])
                    ac1.markdown(f"**{row['AgencyName']}**")
                    web_small = str(row.get('WebAddress', '') or '').strip()
                    if web_small:
                        ac1.caption(f"[{web_small}](http://{web_small})")
                    ac1.markdown(f"Code: {row.get('AgencyCode', 'N/A')} | Underwriter: {row.get('PrimaryUnderwriter', 'Unassigned')}")

                    notes_content = row.get('Notes', "")
                    if notes_content:
                        ac1.markdown(f"Notes: {str(notes_content).split('.')[0]}")
                    elif row.get('WebAddress') and row['WebAddress'] != "":
                        ac1.markdown(f"Web: [{row['WebAddress']}](http://{row['WebAddress']})", unsafe_allow_html=True)

                    if ac2.button("Open", key=f"op_{row['AgencyID']}_{idx}"):
                        go_to_agency(row)
                        st.rerun()
                    if ac3.button("Delete", key=f"del_ag_{row['AgencyID']}_{idx}"):
                        st.session_state['agencies'] = st.session_state['agencies'][
                            st.session_state['agencies']['AgencyID'] != row['AgencyID']
                        ]
                        st.session_state['contacts'] = st.session_state['contacts'][
                            st.session_state['contacts']['AgencyID'] != row['AgencyID']
                        ]
                        save_to_csv('agencies')
                        save_to_csv('contacts')
                        st.rerun()
        else:
            st.info("Enter a name in the search box above to quickly find accounts.")
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
                total_agencies = len(latest_office_prod)
                active_mask = latest_office_prod['ActiveFlag'].astype(str).str.upper().isin(
                    ['Y', 'YES', 'ACTIVE', '1', 'TRUE']
                )
                active_count = active_mask.sum()

                # Single-line summary for office production
                st.markdown(
                    f"**{latest_str}** - Total YTD WP: **${total_wp:,.0f}** | YTD NB: **{int(total_nb)}** | Active: **{active_count}/{total_agencies}**"
                )

                st.markdown("**YTD Written Premium Trend (by month, office total):**")
                monthly_agg = office_prod.groupby('Month_dt')[['AllYTDWP', 'AllYTDNB']].sum().sort_index()
                st.line_chart(monthly_agg[['AllYTDWP']])
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
                alc1, alc2, alc3 = st.columns([4, 1, 1])

                alc1.markdown(f"**{row['AgencyName']}**")
                web_small = str(row.get('WebAddress', '') or '').strip()
                if web_small:
                    alc1.caption(f"[{web_small}](http://{web_small})")
                alc1.markdown(f"Code: {row.get('AgencyCode', 'N/A')} | Underwriter: {row.get('PrimaryUnderwriter', 'Unassigned')}")

                notes_content = row.get('Notes', "")
                if notes_content:
                    alc1.markdown(f"Notes: {str(notes_content).split('.')[0]}")
                elif row.get('WebAddress'):
                    alc1.markdown(f"Web: [{row['WebAddress']}](http://{row['WebAddress']})", unsafe_allow_html=True)

                if alc2.button("Open", key=f"list_op_{row['AgencyID']}_{idx}"):
                    go_to_agency(row)
                    st.rerun()

                if alc3.button("Delete", key=f"list_del_{row['AgencyID']}_{idx}"):
                    st.session_state['agencies'] = st.session_state['agencies'][
                        st.session_state['agencies']['AgencyID'] != row['AgencyID']
                    ]
                    st.session_state['contacts'] = st.session_state['contacts'][
                        st.session_state['contacts']['AgencyID'] != row['AgencyID']
                    ]
                    save_to_csv('agencies')
                    save_to_csv('contacts')
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
    with top_right:
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

        if agency_web:
            search_q = quote_plus(
                f"{agency_web} list as many agency contacts and employees as possible with any emails phones or profiles you can find"
            )
            chatgpt_url = f"https://chat.openai.com/?q={search_q}"
            top_right.markdown(
                f'<a href="{chatgpt_url}" target="_blank" rel="noopener noreferrer">'
                f'<button style="margin-top:0.4rem;background-color:#1a6aa8;color:white;border:none;'
                f'padding:0.35rem 0.65rem;border-radius:6px;cursor:pointer;font-size:0.8rem;">'
                f"Search ChatGPT for contacts"
                f"</button></a>",
                unsafe_allow_html=True,
            )

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

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Latest Month", latest['Month'])
            c2.metric("YTD WP", f"${latest['AllYTDWP']:,.0f}")
            c3.metric("YTD New Business", int(latest['AllYTDNB']))
            c4.metric("Active Flag", str(latest['ActiveFlag']))

            if len(prod_match) >= 2:
                prev = prod_match.iloc[1]
                delta = latest['AllYTDWP'] - prev['AllYTDWP']
                if delta > 0:
                    trend_text = f"Increasing (+${delta:,.0f} vs prior month)"
                elif delta < 0:
                    trend_text = f"Decreasing (${delta:,.0f} vs prior month)"
                else:
                    trend_text = "Flat (no change vs prior month)"
                st.write(f"Trend: **{trend_text}**")
            st.markdown("</div>", unsafe_allow_html=True)

    # AGENCY DETAILS
    st.markdown('<div class="panel panel-company">', unsafe_allow_html=True)
    with st.expander("Details and Edit"):
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

    # CONTACTS
    st.markdown('<div class="panel panel-contacts">', unsafe_allow_html=True)
    st.subheader("Contacts")

    with st.expander("Add Contact"):
        with st.form("add_ct"):
            cn = st.text_input("Name*")
            cr = st.text_input("Role")
            ce = st.text_input("Email")
            cp = st.text_input("Phone")
            c_notes = st.text_area("Notes") 
            if st.form_submit_button("Save Contact"):
                new_ct = pd.DataFrame({
                    'ContactID': [get_new_id(st.session_state['contacts'], 'ContactID')],
                    'AgencyID': [agency_id],
                    'Name': [cn],
                    'Role': [cr],
                    'Email': [ce],
                    'Phone': [cp],
                    'Notes': [c_notes]
                })
                st.session_state['contacts'] = pd.concat(
                    [st.session_state['contacts'], new_ct],
                    ignore_index=True
                )
                save_to_csv('contacts')
                st.rerun()

    for idx, row in ag_contacts.iterrows():
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
                    e_c_notes = st.text_area("Notes", value=row['Notes'])

                    col_save, col_cancel = st.columns(2)
                    if col_save.form_submit_button("Save"):
                        contact_index_to_update = st.session_state['contacts'][
                            st.session_state['contacts']['ContactID'] == contact_id
                        ].index[0]
                        st.session_state['contacts'].loc[
                            contact_index_to_update,
                            ['Name', 'Role', 'Email', 'Phone', 'Notes']
                        ] = [e_cn, e_cr, e_ce, e_cp, e_c_notes]
                        st.session_state['editing_contact_id'] = None
                        save_to_csv('contacts')
                        st.success("Contact updated.")
                        st.rerun()
                    if col_cancel.form_submit_button("Cancel"):
                        st.session_state['editing_contact_id'] = None
                        st.rerun()
            else:
                cca, ccb, ccc = st.columns([3, 1, 1])
                cca.markdown(f"**{row['Name']}** ({row['Role']})")

                last_contact_status = get_last_contact_status(contact_name)
                cca.markdown(last_contact_status)

                if row.get('Email') and row['Email'] != "":
                    cca.markdown(f"Email: [{row['Email']}](mailto:{row['Email']})", unsafe_allow_html=True)
                if row.get('Phone') and row['Phone'] != "":
                    cca.markdown(f"Phone: {row['Phone']}")

                cca.markdown(f"Notes: {row['Notes']}")

                if ccb.button("Edit", key=f"edit_ct_{contact_id}"):
                    st.session_state['editing_contact_id'] = contact_id
                    st.rerun()

                if ccc.button("Delete", key=f"del_ct_{contact_id}"):
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
        log_date = st.date_input("Date of Contact", value=datetime.today().date())

        office_employees = employees_in_office

        employees_selected = st.multiselect(
            "Employees",
            options=office_employees,
            default=office_employees[:1] if office_employees else []
        )

        contact_options = ag_contacts['Name'].tolist()
        contacts_selected = st.multiselect(
            "Contacts",
            options=contact_options,
            default=contact_options[:1] if contact_options else []
        )

        type_ = st.radio("Type", ["Call", "Email"], horizontal=True)
        notes = st.text_area("Notes")

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
                            'Notes': notes
                        })

                new_log_df = pd.DataFrame(new_rows)
                st.session_state['logs'] = pd.concat(
                    [st.session_state['logs'], new_log_df],
                    ignore_index=True
                )
                save_to_csv('logs')
                st.success(f"Logged {len(new_rows)} activities.")
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

        unique_types = ['All Types'] + agency_logs['Type'].unique().tolist()
        type_filter = filter_col2.selectbox("Filter by Type", unique_types)
        if type_filter != 'All Types':
            agency_logs = agency_logs[agency_logs['Type'] == type_filter]

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
    st.title(f"Employee: {emp_name}")
    st.write(f"Office: **{emp_office}**")

    # ---- PRODUCTION PANEL ----
    st.markdown('<div class="panel panel-prod">', unsafe_allow_html=True)
    st.subheader("Production - Agencies This Employee Underwrites")

    # Agencies where this employee is PrimaryUnderwriter
    uw_agencies = agencies[agencies['PrimaryUnderwriter'] == emp_name].copy()
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

                st.write(f"Agencies under this employee: **{len(latest_prod)}**")
                st.write(f"Total YTD Written Premium (latest month per agency): **${total_wp:,.0f}**")
                st.write(f"Total YTD New Business Count (latest month per agency): **{int(total_nb)}**")
                st.markdown("")

                # Only show columns that actually exist to avoid KeyErrors
                desired_cols = ['AgencyName', 'AgencyCode', 'Office', 'Month',
                                'AllYTDWP', 'AllYTDNB', 'ActiveFlag']
                display_cols = [c for c in desired_cols if c in latest_prod.columns]

                st.dataframe(
                    latest_prod[display_cols].sort_values('AllYTDWP', ascending=False),
                    use_container_width=True
                )
    elif uw_agencies.empty:
        st.info("This employee is not set as Primary Underwriter for any agencies.")
    else:
        st.info("No production data has been uploaded yet.")
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

# --- ROUTER ---
if login_gate():
    admin_sidebar()
    if st.session_state['view'] == 'company':
        view_company()
    elif st.session_state['view'] == 'office':
        view_office()
    elif st.session_state.get('view') == 'agency':
        view_agency()
    elif st.session_state.get('view') == 'employee':
        view_employee()
    else:
        st.error(f"Unknown view: {st.session_state.get('view')}")
