import re
import json
import urllib.request
import unicodedata
from datetime import date, datetime

import pandas as pd
import streamlit as st

# =========================================================
# CO-PO MANAGEMENT SYSTEM - COMBINED APPLICATION
# =========================================================
st.set_page_config(
    page_title="CO-PO Management System",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

TEACHER_FILE = "teachers.xlsx"
LOGIN_STUDENT_FILE = "students.csv"       # Institutional login register
MARKS_STUDENT_FILE = "csv-4.csv"           # Course-wise marks-entry roster

ASSESSMENT_TYPES = ["PA1", "PA2", "LE", "LM", "ND"]
COMPONENT_TYPES = ["Theory", "Practical", "Project", "Viva", "Internship"]

# --------------------------- Theme ---------------------------
st.markdown("""
<style>
.stApp {background:#F0FAF9;}
[data-testid="stHeader"] {background:#F0FAF9;}
h1,h2,h3 {color:#115E59;}
.main-banner {background:#168C8C;padding:18px 26px;margin-bottom:18px;border-radius:15px;}
.main-banner .heading {color:white;font-size:29px;font-weight:700;}
.main-banner .subheading {color:white;font-size:15px;margin-top:5px;}
.main-banner .credit {color:#FFE08A;font-size:13px;margin-top:12px;}
div[data-testid="stMetric"] {background:#FFFFFF;border:1px solid #99D5CD;padding:12px;border-radius:10px;}
div.stButton > button {background:#0F766E;color:white;border-radius:8px;border:1px solid #0F766E;font-weight:600;min-height:40px;}
div.stButton > button:hover {background:#115E59;color:white;border-color:#115E59;}
div[data-baseweb="tab-list"] {gap:8px;}
div[data-baseweb="tab"] {background:#D9F0EC;border-radius:8px 8px 0 0;color:#115E59;}
div[data-baseweb="tab"][aria-selected="true"] {background:#0F766E;color:white;}
div[data-testid="stSelectbox"] label,div[data-testid="stTextInput"] label {font-weight:600;color:#334155;}
.app-footer {text-align:center;color:#115E59;padding:18px 0 8px;font-size:13px;}
</style>
""", unsafe_allow_html=True)

# ----------------------- Common helpers -----------------------
def clean_text(value):
    if pd.isna(value):
        return ""
    value = unicodedata.normalize("NFKC", str(value)).replace("\u00a0", " ")
    return re.sub(r"\s+", " ", value).strip()

def normalise(value):
    value = clean_text(value).casefold()
    return re.sub(r"[\W_]+", "", value, flags=re.UNICODE)

def normalize_email(value):
    return clean_text(value).lower()

def find_column(df, possible_names):
    lookup = {normalise(c): c for c in df.columns}
    for name in possible_names:
        if normalise(name) in lookup:
            return lookup[normalise(name)]
    return None

def apps_script_post(payload):
    url = st.secrets.get("apps_script_url", "")
    if not url:
        return False, "Google Sheets URL is missing from Streamlit Secrets.", {}
    try:
        request = urllib.request.Request(
            url,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(request, timeout=30) as response:
            result = json.loads(response.read().decode("utf-8"))
        return result.get("status") == "success" or result.get("status") in ("clear", "duplicate"), result.get("message", ""), result
    except Exception as exc:
        return False, str(exc), {}

def check_google_sheets_duplicates(department, values):
    ok, message, result = apps_script_post({
        "action": "check_duplicates",
        "department": department,
        "values": values,
    })
    return result.get("status", "error"), message or result.get("message", "Duplicate check failed.")

def upload_to_google_sheets(department, values, is_edit=False):
    ok, message, result = apps_script_post({
        "action": "save",
        "department": department,
        "values": values,
        "is_edit": is_edit,
    })
    if result.get("status") == "success":
        return True, message or "Google Sheets updated successfully!"
    return False, message or result.get("message", "Google Sheets upload failed.")

def fetch_course_assessments(program, course_code, course_name):
    ok, message, result = apps_script_post({
        "action": "fetch_assessments",
        "program": program,
        "course_code": course_code,
        "course_name": course_name,
    })
    if result.get("status") == "success":
        return True, message or "Records loaded successfully.", result.get("records", [])
    return False, message or result.get("message", "Unable to fetch records."), []

def update_course_assessments(updates):
    ok, message, result = apps_script_post({
        "action": "update_assessments",
        "updates": updates,
    })
    if result.get("status") == "success":
        return True, message or "Assessment records updated successfully."
    return False, message or result.get("message", "Assessment update failed.")

# --------------------------- Load files ---------------------------
@st.cache_data(show_spinner=False)
def load_login_data():
    teachers = pd.read_excel(TEACHER_FILE, dtype=str).fillna("")
    students = pd.read_csv(
        LOGIN_STUDENT_FILE, dtype=str, encoding="utf-8-sig",
        keep_default_na=False,
    )
    teachers.columns = [clean_text(c).lower() for c in teachers.columns]
    students.columns = [clean_text(c).lower() for c in students.columns]
    return teachers, students

@st.cache_data(show_spinner=False)
def load_marks_roster():
    df = pd.read_csv(MARKS_STUDENT_FILE, dtype=str, encoding="utf-8-sig").fillna("")
    df.columns = [clean_text(c) for c in df.columns]
    aliases = {
        "Student ID": ["Student ID", "StudentID", "Student Id"],
        "Admission No": ["Admission No", "Admission Number", "AdmissionNo"],
        "Student Name": ["Student Name", "Name"],
        "Batch": ["Batch"],
        "Category": ["Category", "Programme Type", "Program Type"],
        "Subject": ["Subject", "Department"],
        "Program Title": ["Program Title", "Programme Title", "Course", "Course Title"],
    }
    rename = {}
    for standard, aliases_for_col in aliases.items():
        found = find_column(df, aliases_for_col)
        if found:
            rename[found] = standard
    df = df.rename(columns=rename)
    required = ["Student ID", "Admission No", "Student Name", "Batch", "Category", "Subject", "Program Title"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("csv-4.csv is missing columns: " + ", ".join(missing))
    for col in required:
        df[col] = df[col].apply(clean_text)
    return df

# ------------------------- Assessment helpers -------------------------
def generate_assessment_id(assessment_type, assessment_date, co, component_type):
    assessment_codes = {"PA1": "P1", "PA2": "P2", "LE": "LE", "LM": "LM", "ND": "ND"}
    component_codes = {"Theory": "T", "Practical": "P", "Project": "R", "Viva": "V", "Internship": "I"}
    if assessment_type not in assessment_codes or component_type not in component_codes:
        raise ValueError("Invalid assessment type or component type.")
    return f"{assessment_codes[assessment_type]}{assessment_date.strftime('%d%m%y')}{co}{component_codes[component_type]}"

def extract_maximum_from_heading(column_name):
    match = re.search(r"/\s*(\d+(?:\.\d+)?)\s*$", clean_text(column_name))
    if match and float(match.group(1)) > 0:
        return float(match.group(1))
    return None

def find_maximum_mark_column(df, obtained_column=None):
    names = [
        "Maximum CO Mark", "Max CO Mark", "Maximum CO", "Max CO",
        "Maximum Marks", "Maximum Mark", "Max Marks", "Max Mark",
        "Maximum Score", "Max Score", "Points Possible", "Possible Points",
        "Maximum Points", "Maximum Grade", "Grade Maximum", "Grade Out Of",
        "Marks Out Of", "Out Of",
    ]
    found = find_column(df, names)
    return found if found and found != obtained_column else None

def find_total_mark_table_value(dataframes):
    labels = {
        "maximummarks", "maximummark", "maxmarks", "maxmark",
        "maximumcomark", "maxcomark", "totalmaximummarks",
        "totalmaximummark", "maximumscore", "maxscore",
        "pointspossible", "possiblepoints", "maximumpoints",
        "totalmarks", "totalmark",
    }
    for df in dataframes:
        if df is None or df.empty:
            continue
        values = df.fillna("").astype(str)
        for r in range(len(values)):
            for c in range(len(values.columns)):
                if normalise(values.iat[r, c]) not in labels:
                    continue
                for offset in range(1, 4):
                    for rr, cc in ((r, c + offset), (r + offset, c)):
                        if rr < len(values) and cc < len(values.columns):
                            number = pd.to_numeric(values.iat[rr, cc], errors="coerce")
                            if pd.notna(number) and number > 0:
                                return float(number)
    return None

def get_upload_maximum_marks(uploaded_df, obtained_column, other_sheets=None):
    other_sheets = other_sheets or []
    max_col = find_maximum_mark_column(uploaded_df, obtained_column)
    if max_col:
        vals = pd.to_numeric(uploaded_df[max_col].astype(str).str.replace(",", "", regex=False), errors="coerce")
        if vals.notna().any():
            return vals, f"Maximum marks detected from column: {max_col}"
    heading_max = extract_maximum_from_heading(obtained_column)
    if heading_max is not None:
        return pd.Series([heading_max] * len(uploaded_df), index=uploaded_df.index, dtype=float), f"Maximum marks detected from heading '{obtained_column}': {heading_max}"
    table_max = find_total_mark_table_value(other_sheets)
    if table_max is not None:
        return pd.Series([table_max] * len(uploaded_df), index=uploaded_df.index, dtype=float), f"Maximum marks detected from the total-mark table: {table_max}"
    return None, "Maximum marks could not be identified. Include a heading such as Grade/15, a maximum-mark column, or a separate maximum-mark table."

def process_marks_upload(uploaded_file, roster, assessment_date):
    other_sheets = []
    if uploaded_file.name.lower().endswith(".csv"):
        uploaded_df = pd.read_csv(uploaded_file, dtype=str, encoding="utf-8-sig")
        other_sheets = [uploaded_df.copy()]
    else:
        sheets = pd.read_excel(uploaded_file, sheet_name=None, dtype=str)
        if not sheets:
            raise ValueError("The uploaded Excel file is empty.")
        uploaded_df = None
        for sheet in sheets.values():
            if sheet is None or sheet.empty:
                continue
            sheet.columns = [clean_text(c) or "Unnamed Column" for c in sheet.columns]
            email_col = find_column(sheet, ["Email", "Email Address", "Student Email", "User Email", "Email ID", "E-mail", "E-mail Address"])
            obtained_col = find_column(sheet, ["Obtained CO", "Obtained Marks", "Marks Obtained", "Grade", "Score", "Points Earned", "Earned Points", "Marks"])
            if obtained_col is None:
                obtained_col = next((c for c in sheet.columns if extract_maximum_from_heading(c) is not None), None)
            if email_col and obtained_col:
                uploaded_df = sheet.copy()
                break
        if uploaded_df is None:
            uploaded_df = next(iter(sheets.values())).copy()
            uploaded_df.columns = [clean_text(c) or "Unnamed Column" for c in uploaded_df.columns]
        other_sheets = [s.copy() for s in sheets.values() if s is not None and not s.empty]

    if uploaded_df.empty:
        raise ValueError("The uploaded file contains no student records.")
    # Make duplicate column headings unique.
    seen, new_cols = {}, []
    for col in uploaded_df.columns:
        name = clean_text(col) or "Unnamed Column"
        count = seen.get(name, 0)
        seen[name] = count + 1
        new_cols.append(name if count == 0 else f"{name}.{count}")
    uploaded_df.columns = new_cols

    email_col = find_column(uploaded_df, ["Email", "Email Address", "Student Email", "User Email", "Email ID", "E-mail", "E-mail Address"])
    if not email_col:
        raise ValueError("Email Address column was not found in the uploaded file.")
    obtained_col = find_column(uploaded_df, ["Obtained CO", "Obtained Marks", "Marks Obtained", "Grade", "Score", "Points Earned", "Earned Points", "Marks"])
    if obtained_col is None:
        obtained_col = next((c for c in uploaded_df.columns if extract_maximum_from_heading(c) is not None), None)
    if obtained_col is None:
        raise ValueError("Obtained CO / Grade / Score column was not found.")

    maximum_values, maximum_message = get_upload_maximum_marks(uploaded_df, obtained_col, other_sheets)
    if maximum_values is None:
        raise ValueError(maximum_message)

    roster = roster[["Admission No", "Student Name"]].drop_duplicates().copy()
    roster["Admission No"] = roster["Admission No"].apply(clean_text)
    roster["_key"] = roster["Admission No"].apply(normalise)
    roster_keys = set(roster["_key"])
    def extract_admission(email):
        local = clean_text(email).split("@", 1)[0]
        groups = re.findall(r"(\d+)", local)
        for candidate in reversed(groups):
            if normalise(candidate) in roster_keys:
                return candidate
        return groups[-1] if groups else ""
    work = uploaded_df.copy()
    work["_admission"] = work[email_col].apply(extract_admission)
    work["_key"] = work["_admission"].apply(normalise)
    names = dict(zip(roster["_key"], roster["Student Name"]))
    admissions = dict(zip(roster["_key"], roster["Admission No"]))
    obtained = pd.to_numeric(work[obtained_col].astype(str).str.replace(",", "", regex=False), errors="coerce")
    maximum = pd.to_numeric(maximum_values, errors="coerce")
    result = pd.DataFrame({
        "Admission No": work["_key"].map(admissions).fillna(work["_admission"]),
        "Student Name": work["_key"].map(names).fillna(""),
        "Obtained CO": obtained,
        "Max CO": maximum,
        "Date": assessment_date.strftime("%d/%m/%Y"),
    }).reset_index(drop=True)
    result.attrs["maximum_message"] = maximum_message
    return result

def make_entry_columns(selected_cos):
    columns = {
        "Admission No": st.column_config.TextColumn("Admission No", disabled=True, width="medium"),
        "Student Name": st.column_config.TextColumn("Student Name", disabled=True, width="large"),
    }
    for co in selected_cos:
        columns[f"{co} Mark Obtained"] = st.column_config.NumberColumn(f"{co} Mark Obtained", min_value=0.0, step=0.5, format="%.2f")
        columns[f"{co} Maximum Mark"] = st.column_config.NumberColumn(f"{co} Maximum Mark", min_value=0.0, step=0.5, format="%.2f")
    return columns

def make_blank_entry(student_df, selected_cos, co_max_marks):
    result = student_df[["Admission No", "Student Name"]].copy()
    for co in selected_cos:
        result[f"{co} Mark Obtained"] = 0.0
        result[f"{co} Maximum Mark"] = float(co_max_marks.get(co, 0.0))
    return result

def show_final_report(key):
    if key in st.session_state:
        st.markdown("### Final Submitted Table")
        st.caption("The submitted student marks are displayed below. The records are saved in Google Sheets.")
        st.dataframe(st.session_state[key], use_container_width=True, hide_index=True)

# -------------------------- Session state --------------------------
defaults = {
    "logged_in": False, "user_role": None, "user_name": None,
    "user_email": None, "teacher_courses": None,
    "selected_course_details": None, "show_data_entry": False,
}
for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value

def logout_user():
    for key in list(st.session_state.keys()):
        if key.startswith(("teacher_", "direct_", "upload_", "edit_assessment_", "assessment_editor_")):
            st.session_state.pop(key, None)
    for key in defaults:
        st.session_state[key] = defaults[key]
    st.logout()

# -------------------------- Header --------------------------
st.markdown("""
<div class="main-banner">
  <div class="heading">📊 CO-PO Management System</div>
  <div class="subheading">College Academic Management Portal</div>
  <div class="credit">Developed by MIMAMSA | Department of Statistics, B.C.M. College, Kottayam (Autonomous)</div>
</div>
""", unsafe_allow_html=True)

# ----------------------- Required files -----------------------
required_files = [TEACHER_FILE, LOGIN_STUDENT_FILE, MARKS_STUDENT_FILE]
missing = [name for name in required_files if not __import__("pathlib").Path(name).exists()]
if missing:
    st.error("Required file(s) not found in the app folder: " + ", ".join(missing))
    st.info("Keep teachers.xlsx, students.csv, csv-4.csv and app.py in the same project folder.")
    st.stop()

try:
    teachers, login_students = load_login_data()
    marks_roster = load_marks_roster()
except Exception as exc:
    st.error(f"Unable to load the required data files: {exc}")
    st.stop()

teacher_required = ["course code", "semester", "course name", "course type", "offering department", "teacher name", "email id"]
login_required = ["first name [required]", "last name [required]", "email address [required]"]
missing_teacher = [c for c in teacher_required if c not in teachers.columns]
missing_login = [c for c in login_required if c not in login_students.columns]
if missing_teacher or missing_login:
    st.error("Required columns are missing in the login data files.")
    if missing_teacher:
        st.write("Missing teachers.xlsx columns:", missing_teacher)
    if missing_login:
        st.write("Missing students.csv columns:", missing_login)
    st.stop()

# ---------------------- Google authentication ----------------------
if not st.user.is_logged_in:
    st.session_state["logged_in"] = False
    st.session_state["user_role"] = None
    st.session_state["user_name"] = None
    st.session_state["user_email"] = None
    st.session_state["teacher_courses"] = None
    st.session_state["selected_course_details"] = None
    st.session_state["show_data_entry"] = False
    for key in list(st.session_state.keys()):
        if key.startswith("teacher_"):
            st.session_state.pop(key, None)
    st.subheader("Login")
    st.write("Sign in using your institutional Google account to access the academic management portal.")
    if st.button("Login using Institutional Mail ID", use_container_width=True):
        st.login()
    st.stop()

email = normalize_email(st.user.get("email", ""))
verified = st.user.get("email_verified", False)
if not email or verified is not True or not email.endswith("@bcmcollege.ac.in"):
    st.session_state["logged_in"] = False
    st.error("Access denied. Please sign in using your verified BCM College institutional Google account.")
    if st.button("Sign out"):
        logout_user()
    st.stop()

teacher_record = teachers[teachers["email id"].apply(normalize_email) == email].copy()
student_record = login_students[login_students["email address [required]"].apply(normalize_email) == email].copy()

if not teacher_record.empty and not student_record.empty:
    st.session_state["logged_in"] = False
    st.error("This email appears in both teacher and student records. Please contact the administrator.")
    if st.button("Sign out"):
        logout_user()
    st.stop()
elif not teacher_record.empty:
    teacher_name = clean_text(teacher_record["teacher name"].iloc[0])
    if not teacher_name:
        st.error("Teacher name is missing in teachers.xlsx.")
        if st.button("Sign out"):
            logout_user()
        st.stop()
    st.session_state["logged_in"] = True
    st.session_state["user_role"] = "Teacher"
    st.session_state["user_name"] = teacher_name
    st.session_state["user_email"] = email
    st.session_state["teacher_courses"] = teacher_record
elif not student_record.empty:
    student = student_record.iloc[0]
    st.session_state["logged_in"] = True
    st.session_state["user_role"] = "Student"
    st.session_state["user_name"] = (clean_text(student["first name [required]"]) + " " + clean_text(student["last name [required]"])).strip()
    st.session_state["user_email"] = email
    st.session_state["teacher_courses"] = None
    st.session_state["selected_course_details"] = None
    st.session_state["show_data_entry"] = False
else:
    st.session_state["logged_in"] = False
    st.error("Your institutional email is authenticated, but it is not registered in the college teacher or student records.")
    if st.button("Sign out"):
        logout_user()
    st.stop()

# ------------------------- Logged-in header -------------------------
head_col, logout_col = st.columns([8, 1])
with logout_col:
    if st.button("Logout", key="top_logout", use_container_width=True):
        logout_user()
st.success(f"Welcome, {st.session_state['user_name']}")
st.caption(f"Role: {st.session_state['user_role']}")

# ------------------------- Teacher course selection -------------------------
if st.session_state["user_role"] == "Teacher":
    st.divider()
    st.subheader("Teacher Dashboard")
    assigned = st.session_state["teacher_courses"].copy()
    for col in ["course type", "semester", "course name", "course code", "offering department"]:
        assigned[col] = assigned[col].astype(str).str.strip()
    assigned = assigned[
        (assigned["course type"] != "") & (assigned["semester"] != "") &
        (assigned["course name"] != "") & (assigned["course code"] != "")
    ].copy()

    if assigned.empty:
        st.error("No complete course information is available for this teacher. Please check teachers.xlsx.")
        st.stop()

    def reset_course_type():
        for key in ["teacher_semester", "teacher_course_name"]:
            st.session_state.pop(key, None)
        st.session_state["selected_course_details"] = None
        st.session_state["show_data_entry"] = False

    def reset_semester():
        st.session_state.pop("teacher_course_name", None)
        st.session_state["selected_course_details"] = None
        st.session_state["show_data_entry"] = False

    with st.container(border=True):
        st.markdown("### 📚 COURSE DETAILS")
        st.write("Select your assigned course details.")
        types = sorted(assigned["course type"].unique().tolist(), key=str.casefold)
        selected_type = st.selectbox("Course Type", types, index=None, placeholder="Select Course Type", key="teacher_course_type", on_change=reset_course_type)

        def semester_sort(value):
            match = re.search(r"\d+", str(value))
            return (0, int(match.group())) if match else (1, str(value).lower())

        type_rows = assigned[assigned["course type"] == selected_type].copy() if selected_type else assigned.iloc[0:0].copy()
        semesters = sorted(type_rows["semester"].unique().tolist(), key=semester_sort)
        selected_semester = st.selectbox("Semester", semesters, index=None, placeholder="Select Semester", key="teacher_semester", disabled=selected_type is None, on_change=reset_semester)

        semester_rows = type_rows[type_rows["semester"] == selected_semester].copy() if selected_semester else type_rows.iloc[0:0].copy()
        course_names = sorted(semester_rows["course name"].unique().tolist(), key=str.casefold)
        selected_name = st.selectbox("Course Name", course_names, index=None, placeholder="Select Course Name", key="teacher_course_name", disabled=selected_semester is None)

        chosen = semester_rows[semester_rows["course name"] == selected_name].copy() if selected_name else semester_rows.iloc[0:0].copy()
        if not chosen.empty:
            codes = sorted(chosen["course code"].dropna().astype(str).str.strip().unique().tolist())
            departments = chosen["offering department"].dropna().astype(str).str.strip().unique().tolist()
            if len(codes) == 1 and len(departments) == 1:
                course_code = codes[0]
                department = departments[0]
                st.text_input("Course Code", value=course_code, disabled=True, key="teacher_course_code_display")
                st.session_state["selected_course_details"] = {
                    "Teacher Name": st.session_state["user_name"],
                    "Email ID": st.session_state["user_email"],
                    "Offering Department": department,
                    "Course Type": selected_type,
                    "Semester": selected_semester,
                    "Course Name": selected_name,
                    "Course Code": course_code,
                }
            else:
                st.session_state["selected_course_details"] = None
                st.error("More than one course code or offering department matches this course selection. Please check teachers.xlsx.")
        else:
            st.text_input("Course Code", value="", disabled=True, placeholder="Select a course first", key="teacher_course_code_empty")

        st.divider()
        if st.button("Continue to Data Entry", use_container_width=True, key="continue_to_data_entry"):
            if st.session_state.get("selected_course_details"):
                st.session_state["show_data_entry"] = True
                st.rerun()
            else:
                st.warning("Please select valid course details.")

    if not st.session_state.get("show_data_entry", False):
        st.info("Select your assigned course and click **Continue to Data Entry** to open the marks-entry dashboard.")
        st.stop()

    # Map teacher-assigned course fields to the existing CSV-4 roster fields.
    details = st.session_state["selected_course_details"]
    selected_department = clean_text(details["Offering Department"])
    selected_programme_type = clean_text(details["Course Type"])
    selected_course = clean_text(details["Course Name"])
    course_code = clean_text(details["Course Code"])

    # Department -> Subject; Course Type -> Category; Course Name -> Program Title.
    matching_students = marks_roster[
        (marks_roster["Subject"].apply(normalise) == normalise(selected_department)) &
        (marks_roster["Category"].apply(normalise) == normalise(selected_programme_type)) &
        (marks_roster["Program Title"].apply(normalise) == normalise(selected_course))
    ].copy()

    if matching_students.empty:
        st.warning("No students matched this teacher course in csv-4.csv. Check that Offering Department matches Subject, Course Type matches Category, and Course Name matches Program Title.")
    else:
        st.caption(f"Selected course: {selected_course} · {course_code} · {len(matching_students)} students")
    st.divider()
    st.title("CO-PO Data Entry Dashboard")
    st.caption("Course Outcome Assessment Data Entry System")

    # --------------------------- Tabs ---------------------------
    tab_entry, tab_edit = st.tabs(["Mark Entry", "Edit score"])

    with tab_entry:
        st.subheader("Assessment and Data Entry")
        c1, c2, c3, c4 = st.columns([1.0, 1.25, 1.35, 1.45])
        with c1:
            assessment_type = st.selectbox("Assessment Type", [""] + ASSESSMENT_TYPES, format_func=lambda x: "Select Assessment" if x == "" else x, key="assessment_type_v3")
        with c2:
            assessment_date = st.date_input("Assessment Date", value=date.today(), format="DD/MM/YYYY", key="assessment_date_v3")
        with c3:
            component_type = st.selectbox("Component Type", [""] + COMPONENT_TYPES, format_func=lambda x: "Select Component" if x == "" else x, key="component_type_v3")
        with c4:
            entry_method = st.selectbox("Data Entry Method", ["", "Direct Entry", "Google Classroom", "Moodle"], format_func=lambda x: "Select Entry Method" if x == "" else x, key="entry_method_v4")

        selected_cos, co_max_marks, upload_selected_co = [], {}, None
        number_of_cos = 4
        if entry_method == "Direct Entry":
            st.markdown("#### Direct Entry Settings")
            d1, d2 = st.columns([1, 2])
            with d1:
                number_of_cos = int(st.number_input("Number of COs in course", min_value=1, max_value=20, value=4, step=1, key="number_of_cos_v3"))
            with d2:
                selected_cos = st.multiselect("Select COs in assessment", [f"CO{i}" for i in range(1, number_of_cos + 1)], default=[], placeholder="Select one or more COs", key=f"selected_cos_v3_{normalise(selected_course)}_{normalise(assessment_type)}")
            if selected_cos:
                st.write("Maximum mark for each selected CO")
                max_cols = st.columns(min(len(selected_cos), 5))
                for i, co in enumerate(selected_cos):
                    with max_cols[i % len(max_cols)]:
                        co_max_marks[co] = st.number_input(f"{co} Maximum Mark", min_value=0.0, value=0.0, step=0.5, format="%.2f", key=f"direct_max_mark_{normalise(selected_course)}_{normalise(assessment_type)}_{co}")
        elif entry_method in ("Google Classroom", "Moodle"):
            upload_selected_co = st.selectbox("Select CO in assessment", [f"CO{i}" for i in range(1, 21)], key=f"upload_selected_co_{normalise(selected_course)}_{normalise(assessment_type)}_{entry_method.lower().replace(' ', '_')}")

        context = "_".join([normalise(selected_department), normalise(selected_programme_type), normalise(selected_course), normalise(assessment_type), assessment_date.strftime("%Y%m%d"), normalise(component_type), normalise(entry_method), normalise(upload_selected_co or "")])

        if not entry_method:
            st.info("Please select a Data Entry Method.")
        elif not selected_course:
            st.info("Please select a course.")
        elif not assessment_type:
            st.info("Please select an Assessment Type.")
        elif not component_type:
            st.info("Please select a Component Type.")
        elif matching_students.empty:
            st.warning("There are no matching students for this course in csv-4.csv.")
        elif entry_method == "Direct Entry":
            st.markdown("#### Direct Entry")
            submit_key = f"direct_submitted_{context}"
            report_key = f"direct_report_{context}"
            table_key = f"direct_marks_{context}"
            edit_key = f"direct_edit_mode_{context}"
            submitted = st.session_state.get(submit_key, False)

            if not selected_cos:
                st.info("Please select at least one CO.")
            elif any(float(co_max_marks.get(co, 0)) <= 0 for co in selected_cos):
                st.warning("Enter a maximum mark greater than zero for every selected CO.")
            else:
                roster_students = matching_students[["Admission No", "Student Name"]].drop_duplicates().reset_index(drop=True)
                expected = ["Admission No", "Student Name"] + [x for co in selected_cos for x in (f"{co} Mark Obtained", f"{co} Maximum Mark")]
                if table_key not in st.session_state:
                    st.session_state[table_key] = make_blank_entry(roster_students, selected_cos, co_max_marks)
                current = st.session_state[table_key].copy()
                if list(current.columns) != expected or len(current) != len(roster_students):
                    rebuilt = make_blank_entry(roster_students, selected_cos, co_max_marks)
                    for col in expected:
                        if col in current.columns:
                            for i, value in enumerate(current[col].tolist()[:len(rebuilt)]):
                                rebuilt.at[i, col] = value
                    st.session_state[table_key] = rebuilt
                st.write("Enter each student's obtained marks. You can adjust the maximum mark for an individual student in the table.")
                edited = st.data_editor(st.session_state[table_key], use_container_width=True, hide_index=True, num_rows="fixed", column_config=make_entry_columns(selected_cos), disabled=submitted, key=f"direct_marks_editor_{context}")
                st.session_state[table_key] = edited.copy()
                invalid, invalid_max = [], []
                for co in selected_cos:
                    obtained = pd.to_numeric(edited[f"{co} Mark Obtained"], errors="coerce")
                    maximum = pd.to_numeric(edited[f"{co} Maximum Mark"], errors="coerce")
                    bad_max = maximum.isna() | (maximum <= 0)
                    bad_obtained = obtained.isna() | (obtained < 0) | (obtained > maximum)
                    invalid_max.extend(edited.loc[bad_max, "Admission No"].astype(str).tolist())
                    invalid.extend(edited.loc[bad_obtained, "Admission No"].astype(str).tolist())
                if invalid_max:
                    st.error("Maximum marks must be greater than zero. Correct: " + ", ".join(sorted(set(invalid_max))))
                if invalid:
                    st.error("Obtained marks cannot be negative or greater than the maximum mark. Correct: " + ", ".join(sorted(set(invalid))))
                if not submitted and st.button("Submit", key=f"direct_submit_button_{context}"):
                    if invalid or invalid_max:
                        st.error("Please correct invalid marks before submitting.")
                    else:
                        date_text = assessment_date.strftime("%d/%m/%Y")
                        report_rows, sheet_values = [], []
                        for _, student_row in edited.iterrows():
                            for co in selected_cos:
                                obtained = float(student_row[f"{co} Mark Obtained"])
                                maximum = float(student_row[f"{co} Maximum Mark"])
                                report_rows.append({
                                    "Admission No": student_row["Admission No"],
                                    "Student Name": student_row["Student Name"],
                                    "Program": selected_programme_type,
                                    "Course Code": course_code,
                                    "Course Name": selected_course,
                                    "Component Type": component_type,
                                    "Assessment Type": assessment_type,
                                    "Assessment Date": date_text,
                                    "CO Mark Obtained": obtained,
                                    "CO Maximum Mark": maximum,
                                })
                                sheet_values.append([
                                    str(student_row["Admission No"]), str(student_row["Student Name"]),
                                    str(selected_programme_type), str(course_code), str(selected_course),
                                    generate_assessment_id(assessment_type, assessment_date, co, component_type),
                                    obtained, maximum, "", "Pending",
                                ])
                        is_edit = st.session_state.get(edit_key, False)
                        if not is_edit:
                            dup_status, dup_message = check_google_sheets_duplicates(selected_department, sheet_values)
                            if dup_status == "duplicate":
                                st.warning("This data has already been recorded for this assessment and date.")
                            elif dup_status != "clear":
                                st.error(f"Duplicate check failed: {dup_message}")
                            else:
                                success, message = upload_to_google_sheets(selected_department, sheet_values, is_edit=False)
                                if success:
                                    st.session_state[report_key] = pd.DataFrame(report_rows)
                                    st.session_state[submit_key] = True
                                    st.session_state[edit_key] = False
                                    st.session_state["_google_sheet_upload_status"] = (True, message)
                                    st.rerun()
                                else:
                                    st.error(f"Google Sheets upload failed: {message}")
                        else:
                            success, message = upload_to_google_sheets(selected_department, sheet_values, is_edit=True)
                            if success:
                                st.session_state[report_key] = pd.DataFrame(report_rows)
                                st.session_state[submit_key] = True
                                st.session_state[edit_key] = False
                                st.session_state["_google_sheet_upload_status"] = (True, message)
                                st.rerun()
                            else:
                                st.error(f"Google Sheets upload failed: {message}")
                if submitted:
                    show_final_report(report_key)
                    if st.button("Edit CO Table", key=f"direct_edit_co_{context}"):
                        st.session_state[edit_key] = True
                        st.session_state[submit_key] = False
                        st.rerun()

        elif entry_method in ("Google Classroom", "Moodle"):
            st.markdown(f"#### {entry_method} Data Upload")
            upload_submit_key = f"upload_submitted_{context}"
            upload_report_key = f"upload_report_{context}"
            upload_data_key = f"upload_processed_data_{context}"
            submitted = st.session_state.get(upload_submit_key, False)
            uploaded_file = st.file_uploader(f"Upload {entry_method} marks file", type=["csv", "xlsx"], key=f"upload_file_{context}", disabled=submitted)
            fingerprint_key = f"{upload_data_key}_fingerprint"
            if uploaded_file is not None and not submitted:
                fingerprint = (uploaded_file.name, uploaded_file.size)
                if st.session_state.get(fingerprint_key) != fingerprint:
                    try:
                        uploaded_file.seek(0)
                        processed = process_marks_upload(uploaded_file, matching_students, assessment_date)
                        st.session_state[upload_data_key] = processed.copy()
                        st.session_state[f"{upload_data_key}_filename"] = uploaded_file.name
                        st.session_state[f"{upload_data_key}_maximum_message"] = processed.attrs.get("maximum_message", "Maximum marks detected.")
                        st.session_state[fingerprint_key] = fingerprint
                        st.success(f"{entry_method} file uploaded and processed successfully.")
                    except Exception as exc:
                        st.error(f"Unable to process the uploaded file: {exc}")
            if upload_data_key in st.session_state:
                processed = st.session_state[upload_data_key].copy()
                st.info(st.session_state.get(f"{upload_data_key}_maximum_message", "Maximum marks detected."))
                st.metric("Number of uploaded records", len(processed))
                unmatched = int((processed["Student Name"].astype(str).str.strip() == "").sum())
                if unmatched:
                    st.warning(f"{unmatched} record(s) could not be matched with csv-4.csv. Check the admission numbers.")
                if not submitted:
                    st.subheader("Processed Student Details")
                    processed = st.data_editor(
                        processed, use_container_width=True, hide_index=True, num_rows="fixed",
                        key=f"upload_processed_editor_{context}",
                        column_config={
                            "Admission No": st.column_config.TextColumn("Admission No"),
                            "Student Name": st.column_config.TextColumn("Student Name"),
                            "Obtained CO": st.column_config.NumberColumn("Obtained CO", format="%.2f", min_value=0),
                            "Max CO": st.column_config.NumberColumn("Max CO", format="%.2f", min_value=0, disabled=True),
                            "Date": st.column_config.TextColumn("Date", disabled=True),
                        },
                    )
                    st.session_state[upload_data_key] = processed.copy()
                    if st.button("Submit", key=f"upload_submit_button_{context}"):
                        obtained = pd.to_numeric(processed["Obtained CO"], errors="coerce")
                        maximum = pd.to_numeric(processed["Max CO"], errors="coerce")
                        if obtained.isna().any():
                            st.error("Please correct missing or invalid obtained marks.")
                        elif maximum.isna().any() or (maximum <= 0).any():
                            st.error("Maximum marks are missing or invalid.")
                        elif (obtained < 0).any() or (obtained > maximum).any():
                            st.error("Obtained marks cannot be negative or greater than maximum marks.")
                        else:
                            date_text = assessment_date.strftime("%d/%m/%Y")
                            report = pd.DataFrame({
                                "Admission No": processed["Admission No"],
                                "Student Name": processed["Student Name"],
                                "Program": selected_programme_type,
                                "Course Code": course_code,
                                "Course Name": selected_course,
                                "Component Type": component_type,
                                "Assessment Type": assessment_type,
                                "Assessment Date": date_text,
                                "CO Mark Obtained": obtained,
                                "CO Maximum Mark": maximum,
                            })
                            assessment_id = generate_assessment_id(assessment_type, assessment_date, upload_selected_co, component_type)
                            values = [[str(row["Admission No"]), str(row["Student Name"]), str(row["Program"]), str(row["Course Code"]), str(row["Course Name"]), assessment_id, float(row["CO Mark Obtained"]), float(row["CO Maximum Mark"]), "", "Pending"] for _, row in report.iterrows()]
                            dup_status, dup_message = check_google_sheets_duplicates(selected_department, values)
                            if dup_status == "duplicate":
                                st.warning("This data has already been recorded for this assessment and date.")
                            elif dup_status != "clear":
                                st.error(f"Duplicate check failed: {dup_message}")
                            else:
                                success, message = upload_to_google_sheets(selected_department, values, is_edit=False)
                                if success:
                                    st.session_state[upload_report_key] = report
                                    st.session_state[upload_submit_key] = True
                                    st.session_state["_google_sheet_upload_status"] = (True, message)
                                    st.rerun()
                                else:
                                    st.error(f"Google Sheets upload failed: {message}")
                else:
                    show_final_report(upload_report_key)
            else:
                st.info(f"Upload the {entry_method} marks file after selecting the course and CO. Maximum marks are detected from the uploaded file.")

        if "_google_sheet_upload_status" in st.session_state:
            status, message = st.session_state["_google_sheet_upload_status"]
            if status:
                st.success("Confirmed – Google Sheets updated successfully!")
            else:
                st.error(f"Google Sheets upload failed: {message}")

    # --------------------------- Edit score tab ---------------------------
    with tab_edit:
        st.subheader("Edit score")
        st.caption("Find submitted student assessment records and update Obtained Mark or Maximum Mark.")
        course_context = "_".join([normalise(selected_programme_type), normalise(selected_course)])
        records_key = f"edit_assessment_records_{course_context}"
        edit_mode_key = f"edit_assessment_mode_{course_context}"

        if st.button("Load Assessment Details", key=f"load_assessment_{course_context}"):
            with st.spinner("Loading records from Google Sheets..."):
                ok, message, records = fetch_course_assessments(selected_programme_type, course_code, selected_course)
            if ok:
                parsed = []
                assessment_map = {"P1": "PA1", "P2": "PA2", "LE": "LE", "LM": "LM", "ND": "ND"}
                component_map = {"T": "Theory", "P": "Practical", "R": "Project", "V": "Viva", "I": "Internship"}
                pattern = r"^(P1|P2|LE|LM|ND)(\d{6})(CO\d+)(T|P|R|V|I)$"
                for record in records:
                    vals = record.get("values", [])
                    if len(vals) < 8:
                        continue
                    assessment_id = str(vals[5]).strip()
                    match = re.fullmatch(pattern, assessment_id)
                    if not match:
                        continue
                    code, date_code, co, component_code = match.groups()
                    try:
                        date_text = datetime.strptime(date_code, "%d%m%y").strftime("%d/%m/%Y")
                    except ValueError:
                        continue
                    obtained = pd.to_numeric(vals[6], errors="coerce")
                    maximum = pd.to_numeric(vals[7], errors="coerce")
                    parsed.append({
                        "rowNumber": int(record["rowNumber"]), "expected": vals[:6],
                        "Admission No": str(vals[0]), "Student Name": str(vals[1]),
                        "CO": co, "Assessment ID": assessment_id,
                        "Assessment Type": assessment_map[code], "Assessment Date": date_text,
                        "Component Type": component_map[component_code],
                        "Obtained Mark": float(obtained) if pd.notna(obtained) else 0.0,
                        "Maximum Mark": float(maximum) if pd.notna(maximum) else 0.0,
                    })
                st.session_state[records_key] = parsed
                st.session_state[edit_mode_key] = False
                st.success(message)
            else:
                st.error(message)

        if records_key in st.session_state:
            all_records = st.session_state[records_key]
            if not all_records:
                st.warning("No valid assessment records were found for this course.")
            else:
                records_df = pd.DataFrame(all_records)
                f1, f2, f3 = st.columns(3)
                with f1:
                    assessment_options = sorted(records_df["Assessment Type"].unique().tolist())
                    selected_edit_assessment = st.selectbox("Assessment Type", assessment_options, key=f"edit_assessment_type_{course_context}")
                with f2:
                    date_options = sorted(
                        records_df.loc[records_df["Assessment Type"] == selected_edit_assessment, "Assessment Date"].unique().tolist(),
                        key=lambda x: datetime.strptime(x, "%d/%m/%Y"),
                    )
                    selected_edit_date = st.selectbox("Assessment Date", date_options, key=f"edit_assessment_date_{course_context}_{selected_edit_assessment}")
                with f3:
                    component_options = sorted(records_df.loc[
                        (records_df["Assessment Type"] == selected_edit_assessment) &
                        (records_df["Assessment Date"] == selected_edit_date),
                        "Component Type",
                    ].unique().tolist())
                    selected_edit_component = st.selectbox("Component Type", component_options, key=f"edit_component_{course_context}_{selected_edit_assessment}_{selected_edit_date}")

                filtered = records_df[
                    (records_df["Assessment Type"] == selected_edit_assessment) &
                    (records_df["Assessment Date"] == selected_edit_date) &
                    (records_df["Component Type"] == selected_edit_component)
                ].copy()
                display_cols = ["Admission No", "Student Name", "CO", "Obtained Mark", "Maximum Mark"]
                st.markdown("### Student Assessment Records")
                st.metric("Number of records", len(filtered))
                originals = {int(row["rowNumber"]): row for row in all_records}

                if not st.session_state.get(edit_mode_key, False):
                    st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)
                    if st.button("Edit", key=f"enable_edit_{course_context}"):
                        st.session_state[edit_mode_key] = True
                        st.rerun()
                else:
                    st.info("You can now change the Obtained Mark and Maximum Mark columns.")
                    editable = filtered[display_cols].copy()
                    editable.index = filtered["rowNumber"].astype(int)
                    edited = st.data_editor(
                        editable, use_container_width=True, hide_index=True, num_rows="fixed",
                        key=f"assessment_editor_{course_context}",
                        column_config={
                            "Admission No": st.column_config.TextColumn("Admission No", disabled=True),
                            "Student Name": st.column_config.TextColumn("Student Name", disabled=True),
                            "CO": st.column_config.TextColumn("CO", disabled=True),
                            "Obtained Mark": st.column_config.NumberColumn("Obtained Mark", min_value=0.0, step=0.5, format="%.2f"),
                            "Maximum Mark": st.column_config.NumberColumn("Maximum Mark", min_value=0.0, step=0.5, format="%.2f"),
                        },
                    )
                    save_col, cancel_col = st.columns(2)
                    with save_col:
                        if st.button("Save Changes", key=f"save_assessment_changes_{course_context}"):
                            updates, errors = [], []
                            for row_num, row in edited.iterrows():
                                try:
                                    obtained, maximum = float(row["Obtained Mark"]), float(row["Maximum Mark"])
                                except (ValueError, TypeError):
                                    errors.append(f"Invalid marks in row {row_num}.")
                                    continue
                                if pd.isna(obtained) or pd.isna(maximum) or obtained < 0 or maximum <= 0 or obtained > maximum:
                                    errors.append(f"Check marks for Admission No {row['Admission No']}.")
                                    continue
                                original = originals.get(int(row_num))
                                if original is None:
                                    errors.append(f"Original record not found for row {row_num}.")
                                    continue
                                if obtained != float(original["Obtained Mark"]) or maximum != float(original["Maximum Mark"]):
                                    updates.append({"rowNumber": int(row_num), "expected": original["expected"], "obtained": obtained, "maximum": maximum})
                            if errors:
                                for error in errors:
                                    st.error(error)
                            elif not updates:
                                st.info("No changes were made.")
                            else:
                                with st.spinner("Saving changes to Google Sheets..."):
                                    ok, message = update_course_assessments(updates)
                                if ok:
                                    st.success(message)
                                    st.session_state.pop(records_key, None)
                                    st.session_state.pop(edit_mode_key, None)
                                    st.rerun()
                                else:
                                    st.error(message)
                    with cancel_col:
                        if st.button("Cancel", key=f"cancel_assessment_edit_{course_context}"):
                            st.session_state[edit_mode_key] = False
                            st.rerun()

# -------------------------- Student page --------------------------
elif st.session_state["user_role"] == "Student":
    st.divider()
    st.subheader("Student Dashboard")
    st.info("Student login successful. The final marksheet module can be connected here.")

st.markdown('<div class="app-footer">Developed by MIMAMSA | Department of Statistics</div>', unsafe_allow_html=True)

