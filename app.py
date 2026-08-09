import io
import os
import pandas as pd
import streamlit as st
from datetime import datetime

# Page configuration
st.set_page_config(page_title="Mirage Absence Portal", page_icon="📅", layout="wide")

# --- GLOBAL SYSTEM FILES ---
SHARED_FILE = "shared_absences.xlsx"
STATUS_FILE = "absence_portal_status.txt"  
ADMIN_PASSWORD = "Mirage_Absence_Secured_2026!#$xK9"

# --- INITIALIZE SESSION STATES ---
if "admin_logged_in" not in st.session_state:
    st.session_state.admin_logged_in = False
if "logged_in_user" not in st.session_state:
    st.session_state.logged_in_user = None
if "checked_name" not in st.session_state:
    st.session_state.checked_name = None
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0

# --- CORE LOGIC: PORTAL STATUS GATEKEEPER ---
def is_portal_open():
    if not os.path.exists(SHARED_FILE):
        return False
    if not os.path.exists(STATUS_FILE):
        return False
    try:
        with open(STATUS_FILE, "r") as f:
            return f.read().strip() == "OPEN"
    except Exception:
        return False

def set_portal_status(is_open: bool):
    with open(STATUS_FILE, "w") as f:
        f.write("OPEN" if is_open else "CLOSED")

# --- Language Translations Dictionary ---
translations = {
    "English": {
        "title": "📅 Mirage Employee Attendance & Absence Portal",
        "subtitle": "🆔 Please enter your Name to check your attendance records.",
        "admin_header": "🛠️ Admin Control Panel",
        "admin_pass_label": "🔑 Enter Admin Password:",
        "admin_pass_btn": "🔓 Unlock Admin Panel",
        "admin_access_denied": "❌ Incorrect Admin Password.",
        "admin_panel_unlocked": "✨ Admin Panel Unlocked Successfully!",
        "portal_master_toggle": "🔓 Enable Absence Portal Access",
        "portal_locked_msg": "⚠️ PORTAL LOCKED: Employee login is currently disabled by the Administrator.",
        "upload_label": "📁 Upload Attendance Excel File (.xlsx or .xls)",
        "download_btn": "📥 Download Updated Absence Database",
        "remove_btn": "🗑️ Remove Excel Sheet (Lock Portal & Wipe Data)",
        "refresh_btn": "🔄 Refresh Data",
        "refresh_success": "✅ Data refreshed successfully!",
        "upload_success": "✅ Attendance file uploaded successfully! Portal unlocked.",
        "remove_success": "🗑️ File removed. Portal locked and data wiped.",
        "input_label": "🆔 Employee Name (الاسم):",
        "check_id_btn": "➡️ Next / Verify Name",
        "password_input_label": "🔒 Password (كلمة المرور):",
        "new_password_label": "✨ Create Your Password (أنشئ كلمة المرور):",
        "confirm_password_label": "✔️ Confirm Password (تأكيد كلمة المرور):",
        "register_btn": "🚀 Register & Login",
        "login_btn": "🔑 Login",
        "logout_btn": "🚪 Logout",
        "back_btn": "⬅️ Back",
        "empty_input": "⚠️ Please fill in all required fields.",
        "pass_mismatch": "❌ Passwords do not match. Please try again.",
        "pass_taken": "⚠️ This password is already taken. Please choose another.",
        "error_name": "⚠️ Employee Name not found. Please check and try again.",
        "error_login": "❌ Incorrect Password. Please check and try again.",
        "register_success": "🎉 Password created successfully! Welcome.",
        "error_read": "❌ Error reading file: {error}",
        "dashboard_title": "📊 Employee Absence Breakdown",
        "welcome_banner": "👋 Welcome, {name}!",
        "name_display": "👤 Employee Name:",
        "total_absences_label": "📌 Total Absence Days (إجمالي أيام الغياب):",
        "absence_details_header": "🔍 Unique Absence Dates & Details",
        "table_col_date": "📅 Date / Period",
        "table_col_status": "🔴 Status / Reason",
        "admin_employees_header": "👥 Unique Employee Management",
        "reset_pass_btn": "🔄 Reset Password",
        "reset_success": "✅ Password successfully reset for {name}.",
        "no_absences_msg": "🌟 Excellent attendance record! No recorded absences found."
    },
    "العربية": {
        "title": "📅 بوابة متابعة غياب وحضور العاملين بـ شركة ميراج",
        "subtitle": "🆔 الرجاء إدخال اسم الموظف للاطلاع على سجل الغياب.",
        "admin_header": "🛠️ لوحة تحكم المسؤول (Admin)",
        "admin_pass_label": "🔑 أدخل كلمة مرور المسؤول:",
        "admin_pass_btn": "🔓 فتح لوحة المسؤول",
        "admin_access_denied": "❌ كلمة مرور المسؤول غير صحيحة.",
        "admin_panel_unlocked": "✨ تم فتح لوحة المسؤول بنجاح!",
        "portal_master_toggle": "🔓 تفعيل بوابة الغياب للموظفين",
        "portal_locked_msg": "⚠️ البوابة مغلقة: تسجيل دخول الموظفين معطل حالياً من قِبل المسؤول.",
        "upload_label": "📁 رفع ملف Excel للغياب والحضور (.xlsx أو .xls)",
        "download_btn": "📥 تحميل قاعدة بيانات الغياب الآمنة",
        "remove_btn": "🗑️ حذف ملف الـ Excel (إغلاق البوابة ومسح البيانات)",
        "refresh_btn": "🔄 تحديث البيانات",
        "refresh_success": "✅ تم تحديث البيانات بنجاح!",
        "upload_success": "✅ تم رفع ملف الحضور بنجاح! تم فتح البوابة تلقائياً.",
        "remove_success": "🗑️ تم حذف الملف وإغلاق البوابة ومسح البيانات.",
        "input_label": "🆔 اسم الموظف (Name):",
        "check_id_btn": "➡️ التالي / التحقق من الاسم",
        "password_input_label": "🔒 كلمة المرور (Password):",
        "new_password_label": "✨ أنشئ كلمة المرور الخاصة بك:",
        "confirm_password_label": "✔️ تأكيد كلمة المرور:",
        "register_btn": "🚀 التسجيل والدخول",
        "login_btn": "🔑 تسجيل الدخول",
        "logout_btn": "🚪 تسجيل الخروج",
        "back_btn": "⬅️ رجوع",
        "empty_input": "⚠️ الرجاء ملء جميع الحقول المطلوبة.",
        "pass_mismatch": "❌ كلمتا المرور غير متطابقتين. يرجى المحاولة مرة أخرى.",
        "pass_taken": "⚠️ كلمة المرور هذه مستخدمة من قبل موظف آخر.",
        "error_name": "⚠️ اسم الموظف غير موجود. يرجى التحقق والمحاولة.",
        "error_login": "❌ كلمة المرور غير صحيحة. يرجى التحقق.",
        "register_success": "🎉 تم إنشاء كلمة المرور بنجاح! أهلاً بك.",
        "error_read": "❌ خطأ في قراءة الملف: {error}",
        "dashboard_title": "📊 سجل غياب الموظف والتفاصيل اليومية",
        "welcome_banner": "👋 أهلاً بك يا {name}!",
        "name_display": "👤 اسم الموظف:",
        "total_absences_label": "📌 إجمالي أيام الغياب:",
        "absence_details_header": "🔍 تواريخ وتفاصيل أيام الغياب الفريدة",
        "table_col_date": "📅 التاريخ / الفترة",
        "table_col_status": "🔴 الحالة / سبب الغياب",
        "admin_employees_header": "👥 إدارة الموظفين الفريدين وكلمات المرور",
        "reset_pass_btn": "🔄 إعادة تعيين كلمة المرور",
        "reset_success": "✅ تم إعادة تعيين كلمة المرور للموظف {name} بنجاح.",
        "no_absences_msg": "🌟 سجل حضور ممتاز! لا توجد أيام غياب مسجلة."
    },
}

selected_lang = st.sidebar.selectbox("🌐 Choose Language / اللغة", ["العربية", "English"])
t = translations[selected_lang]

# --- Helper Functions (Cached for High Performance) ---
def read_excel_file(file_path_or_buffer):
    try:
        return pd.read_excel(file_path_or_buffer, dtype=str)
    except Exception as e:
        raise Exception(f"Could not read the Excel file: {e}")

@st.cache_data
def load_excel_df():
    if not os.path.exists(SHARED_FILE):
        return None
    try:
        df = read_excel_file(SHARED_FILE)
        df.columns = df.columns.str.strip()

        if "Password" not in df.columns:
            df["Password"] = ""
        else:
            df["Password"] = (
                df["Password"]
                .fillna("")
                .astype(str)
                .str.replace(r"\.0$", "", regex=True)
                .str.strip()
            )
            df.loc[df["Password"].isin(["nan", "None", ""]), "Password"] = ""

        name_col = next((c for c in df.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
        if name_col:
            df[name_col] = df[name_col].fillna("").astype(str).str.strip()
        return df
    except Exception:
        if os.path.exists(SHARED_FILE):
            try:
                os.remove(SHARED_FILE)
            except Exception:
                pass
        if os.path.exists(STATUS_FILE):
            try:
                os.remove(STATUS_FILE)
            except Exception:
                pass
        return None

def save_excel_safely(df):
    name_col = next((c for c in df.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
    if name_col:
        df[name_col] = df[name_col].astype(str).str.strip()
    if "Password" in df.columns:
        df["Password"] = df["Password"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
        df.loc[df["Password"].isin(["nan", "None", ""]), "Password"] = ""

    df.to_excel(SHARED_FILE, index=False)
    st.cache_data.clear()

def get_employee_absences(df, employee_name):
    """
    Scans all rows belonging to a specific employee Name and aggregates 
    unique absence dates, filtering out duplicates and non-absence entries.
    """
    name_col = next((c for c in df.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
    date_col = next((c for c in df.columns if any(k in c.lower() for k in ["date", "التاريخ", "day"])), None)
    status_col = next((c for c in df.columns if any(k in c.lower() for k in ["status", "الحالة", "البيان", "notes"])), None)

    if not name_col:
        return 0, []

    emp_rows = df[df[name_col].astype(str).str.strip().str.lower() == str(employee_name).strip().lower()]
    absence_records = []
    seen_dates = set()

    for _, row in emp_rows.iterrows():
        row_text = " ".join([str(val).lower() for val in row.values])
        
        is_absent = any(kw in row_text for kw in ["غياب", "absent", "leave", "vacation"]) or \
                    (status_col and any(kw in str(row[status_col]).lower() for kw in ["غياب", "absent"]))

        if is_absent:
            date_val = str(row[date_col]).strip().replace(".0", "") if date_col and pd.notna(row[date_col]) else "Recorded Absence"
            status_val = str(row[status_col]).strip() if status_col and pd.notna(row[status_col]) else "Absent / غياب"

            if date_val not in seen_dates:
                seen_dates.add(date_val)
                absence_records.append({
                    t["table_col_date"]: date_val,
                    t["table_col_status"]: status_val
                })

    return len(absence_records), absence_records

# --- ADMIN SECTION (Sidebar) ---
st.sidebar.markdown("---")
st.sidebar.header(t["admin_header"])

if not st.session_state.admin_logged_in:
    with st.sidebar.form(key="admin_login_form"):
        admin_pass_input = st.text_input(t["admin_pass_label"], type="password")
        submit_admin = st.form_submit_button(t["admin_pass_btn"])
        
        if submit_admin:
            if admin_pass_input == ADMIN_PASSWORD:
                st.session_state.admin_logged_in = True
                st.success(t["admin_panel_unlocked"])
                st.rerun()
            else:
                st.sidebar.error(t["admin_access_denied"])
else:
    has_file = os.path.exists(SHARED_FILE)
    if has_file:
        current_status = is_portal_open()
        master_toggle = st.sidebar.checkbox(t["portal_master_toggle"], value=current_status)
        if master_toggle != current_status:
            set_portal_status(master_toggle)
            st.rerun()
    else:
        st.sidebar.warning("⚠️ Upload an Excel attendance sheet to enable portal access.")

    uploaded_file = st.sidebar.file_uploader(
        t["upload_label"], 
        type=["xlsx", "xls"], 
        key=f"uploader_{st.session_state.uploader_key}"
    )

    if uploaded_file is not None:
        try:
            df_upload = read_excel_file(uploaded_file)
            df_upload.columns = df_upload.columns.str.strip()

            existing_passwords = {}
            if os.path.exists(SHARED_FILE):
                df_old = load_excel_df()
                name_c_old = next((c for c in df_old.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
                if df_old is not None and name_c_old and "Password" in df_old.columns:
                    for _, row in df_old.iterrows():
                        ename = str(row[name_c_old]).strip().lower()
                        pwd = str(row["Password"]).strip()
                        if pwd and pwd.lower() not in ["nan", "none", ""]:
                            existing_passwords[ename] = pwd

            name_c_up = next((c for c in df_upload.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
            pass_col = []
            for _, row in df_upload.iterrows():
                ename = str(row.get(name_c_up, "")).strip().lower() if name_c_up else ""
                pass_col.append(existing_passwords.get(ename, ""))
            df_upload["Password"] = pass_col

            save_excel_safely(df_upload)
            set_portal_status(True)
            
            st.session_state.uploader_key += 1
            st.sidebar.success(t["upload_success"])
            st.rerun()
        except Exception as e:
            st.sidebar.error(t["error_read"].format(error=e))

    if os.path.exists(SHARED_FILE):
        st.sidebar.markdown("---")
        st.sidebar.subheader(t["admin_employees_header"])
        df_admin = load_excel_df()
        if df_admin is not None:
            name_c_adm = next((c for c in df_admin.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
            
            if name_c_adm:
                unique_employees = df_admin.drop_duplicates(subset=[name_c_adm])
                
                for idx, row in unique_employees.iterrows():
                    emp_name = str(row.get(name_c_adm, f"Employee")).strip()
                    current_pwd = str(row.get("Password", "")).strip()
                    has_pass = current_pwd not in ["", "nan", "None"]
                    status_text = "🔒 Registered" if has_pass else "⏳ Not Registered"

                    with st.sidebar.expander(f"👤 {emp_name} ({status_text})"):
                        if has_pass:
                            if st.button(t["reset_pass_btn"], key=f"reset_{emp_name}_{idx}"):
                                df_admin.loc[df_admin[name_c_adm].astype(str).str.strip().str.lower() == emp_name.lower(), "Password"] = ""
                                save_excel_safely(df_admin)
                                st.success(t["reset_success"].format(name=emp_name))
                                st.rerun()
                        else:
                            st.info("ℹ️ No password set yet.")

            st.sidebar.markdown("---")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df_admin.to_excel(writer, index=False)
            excel_bytes = output.getvalue()

            st.sidebar.download_button(
                label=t["download_btn"],
                data=excel_bytes,
                file_name="mirage_absence_database.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.sidebar.markdown("---")
    if st.sidebar.button(t["remove_btn"]):
        if os.path.exists(SHARED_FILE):
            os.remove(SHARED_FILE)
        if os.path.exists(STATUS_FILE):
            os.remove(STATUS_FILE)
        st.cache_data.clear()
        st.rerun()

    if st.sidebar.button("🔒 Lock Admin Panel / قفل لوحة المسؤول"):
        st.session_state.admin_logged_in = False
        st.cache_data.clear()
        st.rerun()

# --- MAIN PAGE LAYOUT ---
col_title, col_refresh = st.columns([4, 1])
with col_title:
    st.title(t["title"])
with col_refresh:
    st.write("")
    if st.button(t["refresh_btn"]):
        st.cache_data.clear()
        st.success(t["refresh_success"])
        st.rerun()

st.markdown("---")

if not is_portal_open():
    st.error(t["portal_locked_msg"])
    st.stop()  

# ====================================================================
# EMPLOYEE ABSENCE PORTAL VIEW
# ====================================================================

if st.session_state.get("logged_in_user"):
    df_verify = load_excel_df()
    user_exists = False
    name_c_ver = next((c for c in df_verify.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None) if df_verify is not None else None
    
    if df_verify is not None and name_c_ver:
        v_match = df_verify[df_verify[name_c_ver].astype(str).str.strip().str.lower() == str(st.session_state.get("logged_in_user")).strip().lower()]
        if not v_match.empty:
            user_exists = True

    if not user_exists:
        st.session_state.logged_in_user = None
        st.session_state.checked_name = None
        st.rerun()

    st.success(t["welcome_banner"].format(name=st.session_state.logged_in_user))
    st.markdown(f"### 📋 {t['dashboard_title']}")
    st.info(f"**{t['name_display']}** `{str(st.session_state.get('logged_in_user')).strip()}`")

    total_absences, absence_list = get_employee_absences(df_verify, st.session_state.get("logged_in_user"))

    col_m1, col_m2 = st.columns(2)
    with col_m1:
        st.metric(label=t["total_absences_label"], value=total_absences)

    st.markdown("---")
    st.subheader(t["absence_details_header"])

    if absence_list:
        df_absences = pd.DataFrame(absence_list)
        st.markdown("""
        <style>
            [data-testid="stTable"] th, 
            [data-testid="stTable"] td {
                text-align: center !important;
                justify-content: center !important;
            }
        </style>
        """, unsafe_allow_html=True)
        st.table(df_absences)
    else:
        st.success(t["no_absences_msg"])

    st.markdown("---")
    if st.button(t["logout_btn"]):
        st.session_state.logged_in_user = None
        st.session_state.checked_name = None
        st.rerun()

else:
    st.write(t["subtitle"])
    try:
        df = load_excel_df()
        if df is None:
            st.error(t["error_read"].format(error="Could not load data."))
        else:
            name_c_main = next((c for c in df.columns if any(k in c.lower() for k in ["الاسم", "اسم", "name"])), None)
            
            if not name_c_main:
                st.error("❌ Could not locate Name column in the uploaded Excel file.")
            else:
                if st.session_state.get("checked_name") is None:
                    employee_name_input = st.text_input(t["input_label"], key="employee_name_field")
                    submit_name = st.button(t["check_id_btn"])

                    if submit_name:
                        if not employee_name_input.strip():
                            st.warning(t["empty_input"])
                        else:
                            clean_input_name = employee_name_input.strip()
                            matched = df[df[name_c_main].astype(str).str.strip().str.lower() == clean_input_name.lower()]
                            if not matched.empty:
                                # Standardize to exact case from database
                                exact_name = matched.iloc[0][name_c_main]
                                st.session_state.checked_name = exact_name
                                st.rerun()
                            else:
                                st.error(t["error_name"])
                else:
                    employee_name_input = st.session_state.checked_name
                    df_current = load_excel_df()
                    matched = df_current[df_current[name_c_main].astype(str).str.strip().str.lower() == str(employee_name_input).strip().lower()]

                    if not matched.empty:
                        emp_name = matched.iloc[0][name_c_main]
                        current_pass = str(matched.iloc[0].get("Password", "")).strip()

                        st.info(f"👤 **{emp_name}**")

                        if st.button(t["back_btn"]):
                            st.session_state.checked_name = None
                            st.rerun()

                        if current_pass == "" or current_pass.lower() == "nan":
                            st.info("✨ First time here? Please create a secure password for your attendance account.")
                            new_pass = st.text_input(t["new_password_label"], type="password", key="new_pass_field")
                            confirm_pass = st.text_input(t["confirm_password_label"], type="password", key="new_pass_field_confirm")
                            submit_register = st.button(t["register_btn"])

                            if submit_register:
                                if not new_pass or not confirm_pass:
                                    st.warning(t["empty_input"])
                                elif new_pass != confirm_pass:
                                    st.error(t["pass_mismatch"])
                                else:
                                    existing_passes = df_current["Password"].astype(str).str.strip().tolist()
                                    if new_pass.strip() in existing_passes:
                                        st.error(t["pass_taken"])
                                    else:
                                        df_current.loc[df_current[name_c_main].astype(str).str.strip().str.lower() == str(employee_name_input).strip().lower(), "Password"] = new_pass.strip()
                                        save_excel_safely(df_current)
                                        st.session_state.logged_in_user = emp_name
                                        st.session_state.checked_name = None
                                        st.success(t["register_success"])
                                        st.rerun()
                        else:
                            password_input = st.text_input(t["password_input_label"], type="password", key="password_input_field")
                            submit_login = st.button(t["login_btn"])
                            
                            if submit_login:
                                if not password_input:
                                    st.warning(t["empty_input"])
                                elif password_input.strip() == current_pass:
                                    st.session_state.logged_in_user = emp_name
                                    st.session_state.checked_name = None
                                    st.rerun()
                                else:
                                    st.error(t["error_login"])

    except Exception as e:
        st.error(t["error_read"].format(error=e))
