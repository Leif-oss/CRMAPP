import streamlit as st
import pandas as pd
import os
from datetime import datetime
from dateutil import parser 
import numpy as np 

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="Insurance Marketing CRM (Final)", layout="wide")

# --- FILE SYSTEM SETUP ---
FILES = {
    'offices': 'crm_offices.csv',
    'employees': 'crm_employees.csv',
    'agencies': 'crm_agencies.csv',
    'contacts': 'crm_contacts.csv',
    'logs': 'crm_logs.csv'
}

# --- LOAD / SAVE FUNCTIONS ---
def get_new_id(df, id_col):
    """Helper to generate a unique ID."""
    if df.empty: return 1
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
    
    # 1. OFFICES (List)
    if os.path.exists(FILES['offices']):
        office_df = pd.read_csv(FILES['offices'])
        data['offices'] = office_df['OfficeName'].tolist() if 'OfficeName' in office_df.columns else []
    else:
        data['offices'] = ['New York HQ', 'Chicago Branch', 'Austin Hub']
        pd.DataFrame({'OfficeName': data['offices']}).to_csv(FILES['offices'], index=False)

    # 2. EMPLOYEES
    if os.path.exists(FILES['employees']):
        data['employees'] = pd.read_csv(FILES['employees'])
    else:
        data['employees'] = pd.DataFrame({
            'EmployeeID': [101, 102, 103, 104], 'Name': ['Alice Smith', 'Bob Jones', 'Charlie Day', 'Diana Prince'],
            'Office': ['New York HQ', 'New York HQ', 'Chicago Branch', 'Austin Hub']
        })
    data['employees'].to_csv(FILES['employees'], index=False)
    
    default_underwriter = data['employees']['Name'].iloc[0] if not data['employees'].empty else ""

    # 3. AGENCIES
    if os.path.exists(FILES['agencies']):
        data['agencies'] = pd.read_csv(FILES['agencies'])
        
        # --- FIX/ENSURE: Add missing columns for PrimaryUnderwriter, WebAddress, etc. ---
        required_cols = ['WebAddress', 'AgencyCode', 'Notes', 'PrimaryUnderwriter']
        
        for col in required_cols:
            if col not in data['agencies'].columns:
                if col == 'PrimaryUnderwriter':
                    data['agencies'][col] = default_underwriter
                else:
                    data['agencies'][col] = "" 
        # --- END FIX/ENSURE ---
        
        # --- FIX: Ensure Notes column has no NaN values (float) ---
        # This is the primary defense against the AttributeError
        data['agencies']['Notes'] = data['agencies']['Notes'].fillna("") 
        # --- END FIX ---

    else:
        # --- INITIAL AGENCIES DATA ---
        default_underwriter_ny = 'Alice Smith'
        default_underwriter_ch = 'Charlie Day'
        default_underwriter_au = 'Diana Prince'

        data['agencies'] = pd.DataFrame({
            'AgencyID': [501, 502, 503, 504], 'AgencyName': ['SafeHands Insurance', 'RapidCover Agency', 'Windy City Brokers', 'Lone Star Life'],
            'Office': ['New York HQ', 'New York HQ', 'Chicago Branch', 'Austin Hub'],
            'WebAddress': ['safehands.com', '', 'windycity.com', ''],
            'AgencyCode': ['SHI-001', 'RCA-002', 'WCB-003', 'LSL-004'],
            'Notes': ["Top performer agency.", "New account, needs attention.", "", "Rural specialist."],
            'PrimaryUnderwriter': [default_underwriter_ny, default_underwriter_ny, default_underwriter_ch, default_underwriter_au] 
        })
        
    data['agencies'].to_csv(FILES['agencies'], index=False) 

    # 4. CONTACTS
    if os.path.exists(FILES['contacts']):
        data['contacts'] = pd.read_csv(FILES['contacts'])
        if 'Notes' not in data['contacts'].columns:
            data['contacts']['Notes'] = "" 
        data['contacts']['Notes'] = data['contacts']['Notes'].fillna("") 
    else:
        data['contacts'] = pd.DataFrame({
            'ContactID': [901, 902, 903, 904], 'AgencyID': [501, 501, 503, 504],
            'Name': ['John Doe', 'Jane Doe', 'Mike Ditka', 'Willie Nelson'],
            'Role': ['Owner', 'Manager', 'Agent', 'Owner'],
            'Email': ['john@safehands.com', 'jane@safehands.com', 'mike@windy.com', 'willie@lonestar.com'],
            'Notes': ["Gatekeeper, prefer email.", "Primary decision maker.", "", "Call after 2 PM."]
        })
    data['contacts'].to_csv(FILES['contacts'], index=False)

    # 5. LOGS
    if os.path.exists(FILES['logs']):
        data['logs'] = pd.read_csv(FILES['logs'])
    else:
        data['logs'] = pd.DataFrame(columns=['Date', 'EmployeeName', 'AgencyName', 'ContactName', 'Type', 'Notes'])
    
    if not data['logs'].empty and 'Date' in data['logs'].columns:
        data['logs']['Date'] = pd.to_datetime(data['logs']['Date'], errors='coerce')
        data['logs'].dropna(subset=['Date'], inplace=True)

    data['logs'].to_csv(FILES['logs'], index=False)
    
    return data

# --- INITIALIZATION ---
if 'data_initialized' not in st.session_state:
    loaded_data = load_data()
    st.session_state.update(loaded_data)
    st.session_state['data_initialized'] = True

# --- SIDEBAR: ADMIN (Global Data Management) ---
with st.sidebar:
    st.header("⚙️ Admin Panel")
    
    with st.expander("Manage Offices"):
        new_office = st.text_input("New Office Name")
        if st.button("Add Office"):
            if new_office and new_office not in st.session_state['offices']:
                st.session_state['offices'].append(new_office)
                save_to_csv('offices')
                st.rerun()
        st.markdown("---")
        office_to_del = st.selectbox("Delete Office", ["Select..."] + st.session_state['offices'])
        if st.button("Delete Selected Office"):
            if office_to_del != "Select...":
                st.session_state['offices'].remove(office_to_del)
                save_to_csv('offices') 
                st.rerun()

    with st.expander("Manage Employees"):
        e_name = st.text_input("New Employee Name")
        e_office = st.selectbox("Assign Office", st.session_state['offices'])
        if st.button("Add Employee"):
            if e_name:
                new_emp = pd.DataFrame({
                    'EmployeeID': [get_new_id(st.session_state['employees'], 'EmployeeID')],
                    'Name': [e_name], 'Office': [e_office]
                })
                st.session_state['employees'] = pd.concat([st.session_state['employees'], new_emp], ignore_index=True)
                save_to_csv('employees')
                st.rerun()
        st.markdown("---")
        emp_to_del = st.selectbox("Delete Employee", ["Select..."] + st.session_state['employees']['Name'].tolist())
        if st.button("Delete Selected Employee"):
            if emp_to_del != "Select...":
                st.session_state['employees'] = st.session_state['employees'][st.session_state['employees']['Name'] != emp_to_del]
                save_to_csv('employees')
                st.rerun()

# --- NAVIGATION ---
if 'view' not in st.session_state: st.session_state['view'] = 'company'
if 'selected_office' not in st.session_state: st.session_state['selected_office'] = None
if 'selected_agency' not in st.session_state: st.session_state['selected_agency'] = None
if 'editing_contact_id' not in st.session_state: st.session_state['editing_contact_id'] = None

def go_to_office(name):
    st.session_state['view'] = 'office'
    st.session_state['selected_office'] = name

def go_to_agency(row):
    st.session_state['view'] = 'agency'
    st.session_state['selected_agency'] = row.to_dict()

# --- HELPER FUNCTION: GET LAST CONTACT DATE ---
def get_last_contact_status(contact_name):
    """
    Checks the logs for the last contact date for a given contact and returns a status string.
    """
    logs = st.session_state['logs']
    contact_logs = logs[logs['ContactName'] == contact_name].copy()
    
    today = datetime.now().date()
    
    if contact_logs.empty:
        return "🚨 **NEVER CONTACTED!**"

    try:
        valid_logs = contact_logs.dropna(subset=['Date'])
        if valid_logs.empty: return "Error: No valid log dates found."
        
        last_date_dt = valid_logs['Date'].max().date() 
        
        time_difference = today - last_date_dt
        days_diff = time_difference.days
        
        last_contact_str = last_date_dt.strftime("%Y-%m-%d")
        
        if days_diff > 90:
            return f"🚨 **Last Contact:** {last_contact_str} ({days_diff} days ago). **OVER 90 DAYS!**"
        elif days_diff > 0:
            return f"⚠️ **Last Contact:** {last_contact_str} ({days_diff} days ago)."
        else:
            return f"✅ **Last Contact:** {last_contact_str} (Today)."
            
    except Exception as e:
        return f"Error processing logs: {e}"

# --- VIEW: COMPANY (Level 1) ---
def view_company():
    st.title("🏢 Company Dashboard")
    logs = st.session_state['logs']
    employees = st.session_state['employees']
    
    calls = logs[logs['Type'] == 'Call'].groupby('EmployeeName').size().reset_index(name='Calls')
    emails = logs[logs['Type'] == 'Email'].groupby('EmployeeName').size().reset_index(name='Emails')
    
    stats = pd.merge(employees, calls, left_on='Name', right_on='EmployeeName', how='left')
    stats = pd.merge(stats, emails, left_on='Name', right_on='EmployeeName', how='left')
    stats[['Calls', 'Emails']] = stats[['Calls', 'Emails']].fillna(0).astype(int)
    
    st.dataframe(stats[['Name', 'Office', 'Calls', 'Emails']], use_container_width=True)
    st.subheader("Select Office")
    cols = st.columns(4)
    for idx, office in enumerate(st.session_state['offices']):
        if cols[idx % 4].button(f"📂 {office}", key=f"btn_{office}"):
            go_to_office(office)
            st.rerun()

# --- VIEW: OFFICE (Level 2) ---
def view_office():
    office = st.session_state['selected_office']
    employees_in_office = st.session_state['employees'][st.session_state['employees']['Office'] == office]['Name'].tolist()
    
    st.button("← Back to Company", on_click=lambda: st.session_state.update({'view': 'company'}))
    st.title(f"📍 Office: {office}")
    
    c1, c2 = st.columns([1, 2])
    with c1:
        st.subheader("Employees")
        st.table(st.session_state['employees'][st.session_state['employees']['Office'] == office]['Name'])

    with c2:
        st.subheader("Insurance Agencies (Search)")
        with st.expander("➕ Add New Agency"):
            with st.form("add_agency_form"):
                new_ag_name = st.text_input("Agency Name*", key='new_ag_name')
                new_ag_web = st.text_input("Web Address", key='new_ag_web')
                new_ag_code = st.text_input("Agency Code", key='new_ag_code')
                
                if employees_in_office:
                    new_ag_uw = st.selectbox("Primary Underwriter", options=employees_in_office, key='new_ag_uw')
                else:
                    st.warning("No employees in this office. Cannot assign Underwriter.")
                    new_ag_uw = None
                
                new_ag_notes = st.text_area("Notes", key='new_ag_notes')
                
                if st.form_submit_button("Add Agency"):
                    if new_ag_name and new_ag_uw:
                        new_ag = pd.DataFrame({
                            'AgencyID': [get_new_id(st.session_state['agencies'], 'AgencyID')],
                            'AgencyName': [new_ag_name], 'Office': [office],
                            'WebAddress': [new_ag_web], 'AgencyCode': [new_ag_code], 
                            'Notes': [new_ag_notes],
                            'PrimaryUnderwriter': [new_ag_uw] 
                        })
                        st.session_state['agencies'] = pd.concat([st.session_state['agencies'], new_ag], ignore_index=True)
                        save_to_csv('agencies')
                        st.rerun()
                    elif not new_ag_name:
                        st.error("Agency Name is required.")
        
        # --- Agency Search Logic (Only showing searched results) ---
        search = st.text_input("Search Agencies (Start typing to see results)...")
        
        agencies = st.session_state['agencies']
        office_agencies_search = agencies[agencies['Office'] == office].copy() 
        
        if search:
            office_agencies_search = office_agencies_search[office_agencies_search['AgencyName'].str.contains(search, case=False, na=False)]
        
        if search:
            if office_agencies_search.empty:
                 st.info(f"No agencies found matching '{search}' in this office.")
            
            for idx, row in office_agencies_search.iterrows():
                with st.container(border=True):
                    ac1, ac2, ac3 = st.columns([3, 1, 1])
                    ac1.markdown(f"**{row['AgencyName']}**")
                    
                    # --- DISPLAY AGENCY CODE AND UNDERWRITER (in search results) ---
                    ac1.markdown(f"Code: **{row.get('AgencyCode', 'N/A')}** | UW: **{row.get('PrimaryUnderwriter', 'Unassigned')}**")
                    
                    # --- Agency Preview Notes/Web Link (FIXED: Handles any remaining float issue) ---
                    notes_content = row.get('Notes', "")
                    
                    if notes_content:
                        # Safely convert to string before split, as load_data should have made it a string
                        ac1.markdown(f"Notes: *{str(notes_content).split('.')[0]}*") 
                    elif row.get('WebAddress') and row['WebAddress'] != "":
                        ac1.markdown(f"Web: [{row['WebAddress']}](http://{row['WebAddress']})", unsafe_allow_html=True)
                    
                    if ac2.button("Open", key=f"op_{row['AgencyID']}"):
                        go_to_agency(row)
                        st.rerun()
                    if ac3.button("Delete", key=f"del_ag_{row['AgencyID']}"):
                        st.session_state['agencies'] = st.session_state['agencies'][st.session_state['agencies']['AgencyID'] != row['AgencyID']]
                        st.session_state['contacts'] = st.session_state['contacts'][st.session_state['contacts']['AgencyID'] != row['AgencyID']]
                        save_to_csv('agencies')
                        save_to_csv('contacts')
                        st.rerun()
        else:
            st.info("Enter an agency name in the search box above to quickly find agencies.")

    st.markdown("---")
    
    ### OFFICE LEVEL: RECENT ACTIVITY
    st.subheader("📅 Recent Agency Activity in this Office")
    
    agency_names_in_office = st.session_state['agencies'][st.session_state['agencies']['Office'] == office]['AgencyName'].tolist()
    
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
        st.info("No activity logs recorded for agencies in this office yet.")

    st.markdown("---")
    
    ### OFFICE LEVEL: ALL AGENCIES LIST (Clickable List)
    st.subheader("📋 All Agencies in this Office")

    office_agencies_all = agencies[agencies['Office'] == office]

    if not office_agencies_all.empty:
        for idx, row in office_agencies_all.iterrows():
            with st.container(border=True):
                alc1, alc2, alc3 = st.columns([4, 1, 1])
                
                alc1.markdown(f"**{row['AgencyName']}**")
                alc1.markdown(f"Code: **{row.get('AgencyCode', 'N/A')}** | UW: **{row.get('PrimaryUnderwriter', 'Unassigned')}**")
                
                # --- Agency Preview Notes/Web Link (FIXED: Handles any remaining float issue) ---
                notes_content = row.get('Notes', "")
                if notes_content:
                     # Safely convert to string before split
                     alc1.markdown(f"Notes: *{str(notes_content).split('.')[0]}*") 
                elif row.get('WebAddress'):
                    alc1.markdown(f"Web: [{row['WebAddress']}](http://{row['WebAddress']})", unsafe_allow_html=True)
                
                if alc2.button("Open", key=f"list_op_{row['AgencyID']}"):
                    go_to_agency(row)
                    st.rerun()
                
                if alc3.button("Delete", key=f"list_del_{row['AgencyID']}"):
                    st.session_state['agencies'] = st.session_state['agencies'][st.session_state['agencies']['AgencyID'] != row['AgencyID']]
                    st.session_state['contacts'] = st.session_state['contacts'][st.session_state['contacts']['AgencyID'] != row['AgencyID']]
                    save_to_csv('agencies')
                    save_to_csv('contacts')
                    st.rerun()
    else:
        st.info("No agencies are currently assigned to this office.")


# --- VIEW: AGENCY (Level 3) ---
def view_agency():
    agency_dict = st.session_state['selected_agency']
    agency_id = agency_dict['AgencyID']
    office = agency_dict['Office']
    
    employees_in_office = st.session_state['employees'][st.session_state['employees']['Office'] == office]['Name'].tolist()
    
    st.button(f"← Back to {agency_dict['Office']}", on_click=lambda: go_to_office(agency_dict['Office']))
    st.title(f"📑 {agency_dict['AgencyName']}")
    
    # --- AGENCY EDIT ---
    with st.expander("ℹ️ Agency Details and Edit"):
        st.markdown(f"**Web:** {agency_dict.get('WebAddress', 'N/A')}")
        st.markdown(f"**Code:** {agency_dict.get('AgencyCode', 'N/A')}")
        st.markdown(f"**Primary Underwriter:** **{agency_dict.get('PrimaryUnderwriter', 'Unassigned')}**") 
        st.markdown("---")
        # Notes display is safe because load_data fills NaNs
        st.markdown("**Notes:**")
        st.write(agency_dict.get('Notes', 'No notes provided.'))
        st.markdown("---")

        current_uw = agency_dict.get('PrimaryUnderwriter')
        
        with st.form("edit_agency_form"):
            e_name = st.text_input("Agency Name", value=agency_dict['AgencyName'])
            e_web = st.text_input("Web Address", value=agency_dict.get('WebAddress', ''))
            e_code = st.text_input("Agency Code", value=agency_dict.get('AgencyCode', ''))
            
            try:
                uw_index = employees_in_office.index(current_uw)
            except ValueError:
                uw_index = 0

            e_uw = st.selectbox("Primary Underwriter", 
                                options=employees_in_office, 
                                index=uw_index)
            
            e_notes = st.text_area("Notes", value=agency_dict.get('Notes', ''))
            
            if st.form_submit_button("Save Agency Changes"):
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

    lc, rc = st.columns(2)
    with lc:
        st.subheader("Contacts")
        
        # --- ADD CONTACT FORM ---
        with st.expander("➕ Add Contact"):
            with st.form("add_ct"):
                cn = st.text_input("Name*")
                cr = st.text_input("Role")
                ce = st.text_input("Email")
                c_notes = st.text_area("Notes") 
                if st.form_submit_button("Save"):
                    new_ct = pd.DataFrame({
                        'ContactID': [get_new_id(st.session_state['contacts'], 'ContactID')],
                        'AgencyID': [agency_id],
                        'Name': [cn], 'Role': [cr], 'Email': [ce], 'Notes': [c_notes]
                    })
                    st.session_state['contacts'] = pd.concat([st.session_state['contacts'], new_ct], ignore_index=True)
                    save_to_csv('contacts')
                    st.rerun()
        
        # --- LIST & EDIT CONTACTS ---
        contacts = st.session_state['contacts']
        ag_contacts = contacts[contacts['AgencyID'] == agency_id].copy()

        for idx, row in ag_contacts.iterrows():
            contact_id = row['ContactID']
            is_editing = st.session_state['editing_contact_id'] == contact_id
            contact_name = row['Name'] 
            
            with st.container(border=True):
                if is_editing:
                    # Edit Form logic
                    with st.form(f"edit_contact_form_{contact_id}"):
                        e_cn = st.text_input("Name", value=row['Name'])
                        e_cr = st.text_input("Role", value=row['Role'])
                        e_ce = st.text_input("Email", value=row['Email'])
                        e_c_notes = st.text_area("Notes", value=row['Notes'])
                        
                        col_save, col_cancel = st.columns(2)
                        if col_save.form_submit_button("Save Contact"):
                            contact_index_to_update = st.session_state['contacts'][st.session_state['contacts']['ContactID'] == contact_id].index[0]
                            st.session_state['contacts'].loc[contact_index_to_update, ['Name', 'Role', 'Email', 'Notes']] = [e_cn, e_cr, e_ce, e_c_notes]
                            st.session_state['editing_contact_id'] = None
                            save_to_csv('contacts')
                            st.success("Contact updated.")
                            st.rerun()
                        if col_cancel.form_submit_button("Cancel"):
                            st.session_state['editing_contact_id'] = None
                            st.rerun()
                else:
                    # Display Mode
                    cca, ccb, ccc = st.columns([3, 1, 1])
                    cca.markdown(f"**{row['Name']}** ({row['Role']})")
                    
                    # --- LAST CONTACT STATUS DISPLAY ---
                    last_contact_status = get_last_contact_status(contact_name)
                    cca.markdown(last_contact_status)
                    
                    # --- CONTACT EMAIL LINK ---
                    if row.get('Email') and row['Email'] != "":
                        cca.markdown(f"Email: [{row['Email']}](mailto:{row['Email']})", unsafe_allow_html=True)
                    
                    cca.markdown(f"**Notes:** {row['Notes']}")
                    
                    if ccb.button("✏️ Edit", key=f"edit_ct_{contact_id}"):
                        st.session_state['editing_contact_id'] = contact_id
                        st.rerun()
                        
                    if ccc.button("🗑️", key=f"del_ct_{contact_id}"):
                        st.session_state['contacts'] = st.session_state['contacts'][st.session_state['contacts']['ContactID'] != contact_id]
                        save_to_csv('contacts')
                        st.rerun()

    with rc:
        st.subheader("📞 Log Activity")
        
        with st.form("log_act"):
            log_date = st.date_input("Date of Contact", value=datetime.today().date())
            
            user = st.selectbox("Employee", st.session_state['employees']['Name'].unique())
            
            contact_options = ag_contacts['Name'].tolist()
            target = st.selectbox("Contact", contact_options if contact_options else ["Unknown"])
            
            type_ = st.radio("Type", ["Call", "Email"], horizontal=True)
            notes = st.text_area("Notes")
            
            if st.form_submit_button("Log"):
                if target == "Unknown":
                    st.error("Please add a contact before logging activity.")
                else:
                    log_datetime = datetime.combine(log_date, datetime.now().time())
                    
                    new_log = pd.DataFrame({
                        'Date': [log_datetime.strftime("%Y-%m-%d %H:%M:%S")], 
                        'EmployeeName': [user],
                        'AgencyName': [agency_dict['AgencyName']],
                        'ContactName': [target],
                        'Type': [type_],
                        'Notes': [notes]
                    })
                    st.session_state['logs'] = pd.concat([st.session_state['logs'], new_log], ignore_index=True)
                    save_to_csv('logs')
                    st.success("Logged!")
                    st.rerun()
    
    st.markdown("---")
    ### AGENCY LEVEL: ALL LOGS (Filterable & Sorted)
    st.subheader("🕰️ Agency Activity History")
    
    agency_logs = st.session_state['logs'][
        st.session_state['logs']['AgencyName'] == agency_dict['AgencyName']
    ].copy()

    if not agency_logs.empty:
        agency_logs['Date'] = pd.to_datetime(agency_logs['Date'], errors='coerce')
        agency_logs.dropna(subset=['Date'], inplace=True)
        
        # --- FILTERING UI ---
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
        # --------------------

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
        st.info("No activity history recorded for this agency yet.")


# --- ROUTER ---
if st.session_state['view'] == 'company':
    view_company()
elif st.session_state['view'] == 'office':
    view_office()
elif st.session_state.get('view') == 'agency':
    view_agency()
else:
    st.error(f"Unknown view: {st.session_state.get('view')}")


