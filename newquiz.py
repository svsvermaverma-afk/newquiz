import streamlit as st
import sqlite3
import pandas as pd
import time
import io
import os
import re
import base64
from datetime import datetime, timedelta, timezone
import streamlit.components.v1 as components

# PDF Generation Libraries (Merit List)
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="ABIC Renukoot - Academic & Quiz Portal",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .school-header {
        text-align: center;
        padding: 14px 10px;
        margin-bottom: 20px;
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        color: white;
        border-radius: 12px;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    }
    .school-header h1 {
        margin: 0;
        font-size: 1.9rem !important;
        font-weight: 800;
        color: #ffffff !important;
    }
    .school-header h3 {
        margin: 5px 0;
        font-size: 1.15rem !important;
        color: #f1f5f9 !important;
    }
    .school-header p {
        margin: 4px 0 0 0;
        font-size: 0.95rem;
        color: #e2e8f0;
    }

    @media only screen and (max-width: 768px) {
        .block-container {
            padding-top: 1rem !important;
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
        }
        .school-header h1 { font-size: 1.3rem !important; }
        .school-header h3 { font-size: 0.95rem !important; }
        .stButton>button {
            width: 100% !important;
            padding: 12px 16px !important;
            font-size: 15px !important;
        }
        [data-testid="column"] {
            width: 100% !important;
            flex: 1 1 100% !important;
            min-width: 100% !important;
        }
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="school-header">
    <h1>ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT</h1>
    <h3>⚡ Physics Subject Exam, Academic & Portfolio Portal ⚡</h3>
    <p>Mentor & In-charge: <b>Shashank Verma, TGT (Physics)</b></p>
</div>
""", unsafe_allow_html=True)

DB_FILE = "master_quiz_system_prod_v17.db"
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "Admin@2026"

STUDENTS_FILE = "students.xlsx"
Q11_FILE = "questions_11.xlsx"
Q12_FILE = "questions_12.xlsx"

IST = timezone(timedelta(hours=5, minutes=30))

def get_ist_now():
    return datetime.now(timezone.utc).astimezone(IST)

def clean_val(val):
    if pd.isna(val) or val is None:
        return ""
    val_str = str(val).strip()
    if val_str.lower() in ["nan", "none", "nat", "<na>", "null"]:
        return ""
    if re.match(r'^-?\d+\.0+$', val_str):
        val_str = val_str.split('.')[0]
    return re.sub(r'\s+', ' ', val_str)

def clean_sr_no(sr_val):
    return clean_val(sr_val)

def safe_b64_decode(data_str):
    if not data_str or len(data_str) < 50:
        return None
    try:
        return base64.b64decode(data_str)
    except Exception:
        return None

# Smart Answer Matcher
def is_answer_correct(selected, correct, opt_a, opt_b, opt_c, opt_d):
    if not selected:
        return False
    s = clean_val(selected).lower()
    c = clean_val(correct).lower()
    a = clean_val(opt_a).lower()
    b = clean_val(opt_b).lower()
    c_opt = clean_val(opt_c).lower()
    d = clean_val(opt_d).lower()
    
    if s == c:
        return True
    
    mapping = {
        'a': a, 'option a': a, '(a)': a, '1': a,
        'b': b, 'option b': b, '(b)': b, '2': b,
        'c': c_opt, 'option c': c_opt, '(c)': c_opt, '3': c_opt,
        'd': d, 'option d': d, '(d)': d, '4': d
    }
    if c in mapping and s == mapping[c]:
        return True
    return False

# ==========================================
# 2. DATABASE MANAGEMENT
# ==========================================
def get_db():
    conn = sqlite3.connect(DB_FILE, timeout=30.0, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    c = conn.cursor()
    
    c.execute('''
        CREATE TABLE IF NOT EXISTS master_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_name TEXT NOT NULL,
            sr_no TEXT NOT NULL,
            normalized_name TEXT UNIQUE NOT NULL,
            roll_no TEXT DEFAULT '',
            target_class TEXT DEFAULT 'Class 12',
            father_name TEXT DEFAULT '',
            mother_name TEXT DEFAULT '',
            dob TEXT DEFAULT '',
            mob_no TEXT DEFAULT '',
            address TEXT DEFAULT '',
            attendance_pct TEXT DEFAULT '82.5',
            attendance_present TEXT DEFAULT '72',
            attendance_total TEXT DEFAULT '87',
            test_total TEXT DEFAULT '-',
            test_pct TEXT DEFAULT '-',
            short_term_goal TEXT DEFAULT '',
            long_term_goal TEXT DEFAULT '',
            photo_b64 TEXT DEFAULT '',
            photo_url TEXT DEFAULT ''
        )
    ''')

    try:
        c.execute("ALTER TABLE master_students ADD COLUMN target_class TEXT DEFAULT 'Class 12'")
    except sqlite3.OperationalError:
        pass

    c.execute('''
        CREATE TABLE IF NOT EXISTS quizzes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            target_class TEXT NOT NULL,
            topic TEXT NOT NULL,
            quiz_title TEXT UNIQUE NOT NULL,
            duration_minutes INTEGER DEFAULT 15,
            start_datetime TEXT NOT NULL,
            end_datetime TEXT NOT NULL,
            is_active INTEGER DEFAULT 1
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS questions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quiz_id INTEGER NOT NULL,
            question TEXT NOT NULL,
            option_a TEXT NOT NULL,
            option_b TEXT NOT NULL,
            option_c TEXT NOT NULL,
            option_d TEXT NOT NULL,
            correct_option TEXT NOT NULL
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS submissions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quiz_id INTEGER NOT NULL,
            student_name TEXT NOT NULL,
            sr_no TEXT NOT NULL,
            score INTEGER NOT NULL,
            total_questions INTEGER NOT NULL,
            tab_switches INTEGER DEFAULT 0,
            status TEXT DEFAULT 'Completed',
            submitted_at TEXT NOT NULL,
            UNIQUE(quiz_id, student_name)
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS student_responses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quiz_id INTEGER NOT NULL,
            student_name TEXT NOT NULL,
            sr_no TEXT NOT NULL,
            question_id INTEGER NOT NULL,
            question_text TEXT NOT NULL,
            selected_option TEXT,
            correct_option TEXT NOT NULL,
            is_correct INTEGER NOT NULL,
            recorded_at TEXT NOT NULL
        )
    ''')

    c.execute('''
        CREATE TABLE IF NOT EXISTS quiz_attempts (
            quiz_id INTEGER NOT NULL,
            normalized_name TEXT NOT NULL,
            start_epoch REAL NOT NULL,
            PRIMARY KEY(quiz_id, normalized_name)
        )
    ''')

    c.execute("SELECT COUNT(*) FROM quizzes")
    if c.fetchone()[0] == 0:
        now_time = get_ist_now() - timedelta(hours=1)
        s_date = now_time.strftime("%Y-%m-%d %H:%M")
        e_date = (now_time + timedelta(days=30)).strftime("%Y-%m-%d %H:%M")
        c.execute('''
            INSERT INTO quizzes (target_class, topic, quiz_title, duration_minutes, start_datetime, end_datetime, is_active)
            VALUES (?, ?, ?, ?, ?, ?, 1)
        ''', ("Class 11", "Laws of Motion & Work Energy", "Class 11 - Physics Exam", 15, s_date, e_date))
        c.execute('''
            INSERT INTO quizzes (target_class, topic, quiz_title, duration_minutes, start_datetime, end_datetime, is_active)
            VALUES (?, ?, ?, ?, ?, ?, 1)
        ''', ("Class 12", "Electrostatics & Magnetism", "Class 12 - Physics Exam", 20, s_date, e_date))

    for s_path in [STUDENTS_FILE, "students.csv"]:
        if os.path.exists(s_path):
            try:
                s_df = pd.read_csv(s_path) if s_path.endswith(".csv") else pd.read_excel(s_path)
                s_df.columns = [str(col).strip().lower().replace(" ", "_") for col in s_df.columns]
                n_col = next((col for col in s_df.columns if col in ["name", "student_name", "student", "studentname"]), s_df.columns[0])
                sr_col = next((col for col in s_df.columns if col in ["sr_no", "srno", "sr", "roll_no", "rollno", "id", "password"]), s_df.columns[1] if len(s_df.columns) > 1 else s_df.columns[0])
                r_col = next((col for col in s_df.columns if "roll" in col), None)

                for _, r in s_df.iterrows():
                    st_nm = clean_val(r[n_col])
                    st_sr = clean_sr_no(r[sr_col])
                    st_roll = clean_val(r[r_col]) if r_col else ""
                    if st_nm and st_sr:
                        c.execute('''
                            INSERT INTO master_students (student_name, sr_no, normalized_name, roll_no)
                            VALUES (?, ?, ?, ?)
                            ON CONFLICT(normalized_name) DO UPDATE SET 
                                student_name=excluded.student_name, 
                                sr_no=excluded.sr_no,
                                roll_no=CASE WHEN excluded.roll_no != '' THEN excluded.roll_no ELSE master_students.roll_no END
                        ''', (st_nm, st_sr, st_nm.lower(), st_roll))
            except Exception:
                pass

    conn.commit()
    conn.close()

init_db()

def get_all_quizzes():
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM quizzes", conn)
    conn.close()
    return df

def get_questions_by_quiz(quiz_id):
    conn = get_db()
    df = pd.read_sql_query("SELECT * FROM questions WHERE quiz_id = ?", conn, params=(quiz_id,))
    conn.close()
    return df

def get_or_set_attempt_start(quiz_id, student_norm_name):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT start_epoch FROM quiz_attempts WHERE quiz_id = ? AND normalized_name = ?", (quiz_id, student_norm_name))
    row = c.fetchone()
    if row:
        start_epoch = row["start_epoch"]
    else:
        start_epoch = time.time()
        c.execute("INSERT OR REPLACE INTO quiz_attempts (quiz_id, normalized_name, start_epoch) VALUES (?, ?, ?)", (quiz_id, student_norm_name, start_epoch))
        conn.commit()
    conn.close()
    return start_epoch

# ==========================================
# 3. PDF MERIT LIST GENERATOR (Admin Only)
# ==========================================
def generate_merit_pdf(subs_df, quiz_info):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle('SchoolTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=18, leading=22, alignment=1, textColor=colors.HexColor("#1e3c72"))
    subtitle_style = ParagraphStyle('SubTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=13, leading=16, alignment=1, textColor=colors.HexColor("#333333"))
    meta_style = ParagraphStyle('Meta', parent=styles['Normal'], fontName='Helvetica', fontSize=10, leading=14, alignment=1, textColor=colors.HexColor("#555555"))
    cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=11, alignment=1)
    cell_bold = ParagraphStyle('CellBold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=11, alignment=1)
    
    elements = []
    elements.append(Paragraph("ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT", title_style))
    elements.append(Paragraph("Official Examination Merit List Report", subtitle_style))
    elements.append(Paragraph(f"<b>Exam:</b> {quiz_info.get('quiz_title', 'Exam')} | <b>Class:</b> {quiz_info.get('target_class', '')} | <b>Topic:</b> {quiz_info.get('topic', '')}", meta_style))
    elements.append(Paragraph(f"Mentor: <b>Shashank Verma, TGT (Physics)</b> | Generated on: {get_ist_now().strftime('%d-%b-%Y %I:%M %p')}", meta_style))
    elements.append(Spacer(1, 15))
    
    table_data = [
        [Paragraph("<b>Rank</b>", cell_bold), Paragraph("<b>Student Name</b>", cell_bold), Paragraph("<b>SR No</b>", cell_bold), Paragraph("<b>Score</b>", cell_bold), Paragraph("<b>Percentage</b>", cell_bold), Paragraph("<b>Switches</b>", cell_bold), Paragraph("<b>Submitted At</b>", cell_bold)]
    ]
    
    for idx, row in subs_df.iterrows():
        rank = idx + 1
        pct = (row['score'] / row['total_questions'] * 100) if row['total_questions'] > 0 else 0
        rank_str = f"🥇 Rank {rank}" if rank == 1 else (f"🥈 Rank {rank}" if rank == 2 else (f"🥉 Rank {rank}" if rank == 3 else f"{rank}"))
        table_data.append([
            Paragraph(rank_str, cell_bold if rank <= 3 else cell_style),
            Paragraph(str(row['student_name']), cell_style),
            Paragraph(str(row['sr_no']), cell_style),
            Paragraph(f"{row['score']} / {row['total_questions']}", cell_bold),
            Paragraph(f"{pct:.1f}%", cell_style),
            Paragraph(str(row['tab_switches']), cell_style),
            Paragraph(str(row['submitted_at']), cell_style)
        ])
    
    col_widths = [65, 130, 65, 65, 60, 50, 100]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1e3c72")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#dcdcdc")),
    ]
    
    for r_idx in range(1, len(table_data)):
        if r_idx == 1:
            t_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor("#fff9db")))
        elif r_idx == 2:
            t_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor("#f1f3f5")))
        elif r_idx == 3:
            t_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor("#fff4e6")))
        elif r_idx % 2 == 0:
            t_style.append(('BACKGROUND', (0, r_idx), (-1, r_idx), colors.HexColor("#f8f9fa")))
            
    t.setStyle(TableStyle(t_style))
    elements.append(t)
    doc.build(elements)
    pdf_val = buffer.getvalue()
    buffer.close()
    return pdf_val

# Anti-Cheating & Live Timer Component
def inject_live_timer_and_security(remaining_seconds, quiz_id, student_name):
    timer_js = f"""
    <style>
        #sticky-timer-box {{
            position: fixed; 
            top: 50px; 
            right: 20px; 
            background: #ff4b4b; 
            color: #ffffff; 
            padding: 10px 18px; 
            border-radius: 8px; 
            font-family: monospace; 
            font-size: 18px; 
            font-weight: bold; 
            z-index: 999999;
            box-shadow: 0 4px 12px rgba(0,0,0,0.25);
            border: 2px solid white;
        }}
        @media only screen and (max-width: 600px) {{
            #sticky-timer-box {{
                top: 40px;
                right: 8px;
                padding: 6px 12px;
                font-size: 14px;
            }}
        }}
    </style>
    <div id="sticky-timer-box">
        ⏳ <span id="timer-display">Loading...</span> | ⚠️ <span id="switch-count">0</span>
    </div>

    <script>
    let timeLeft = {int(remaining_seconds)};
    let display = document.getElementById('timer-display');
    let switchCountElem = document.getElementById('switch-count');
    let tabSwitches = sessionStorage.getItem('tab_switches_{quiz_id}_{student_name}') || 0;
    switchCountElem.innerHTML = tabSwitches;

    function triggerAutoSubmit() {{
        let buttons = window.parent.document.querySelectorAll('button');
        buttons.forEach(btn => {{
            if (btn.innerText.includes("Submit Final Answers")) {{ btn.click(); }}
        }});
    }}

    function updateTimer() {{
        if (timeLeft <= 0) {{
            display.innerHTML = "TIME UP!";
            triggerAutoSubmit();
            return;
        }}
        let mins = Math.floor(timeLeft / 60);
        let secs = timeLeft % 60;
        display.innerHTML = (mins < 10 ? "0" : "") + mins + ":" + (secs < 10 ? "0" : "") + secs;
        timeLeft--;
    }}

    updateTimer();
    setInterval(updateTimer, 1000);

    window.addEventListener('blur', function() {{
        tabSwitches++;
        sessionStorage.setItem('tab_switches_{quiz_id}_{student_name}', tabSwitches);
        switchCountElem.innerHTML = tabSwitches;
        alert('⚠️ WARNING (' + tabSwitches + '/3): Tab switch detect hua hai! Bar-bar tab badalne par test auto-submit ho jayega.');
        if (tabSwitches >= 3) {{
            alert('❌ Maximum limit reached. Test auto-submit ho raha hai.');
            triggerAutoSubmit();
        }}
    }});

    document.addEventListener('contextmenu', function(e) {{ e.preventDefault(); }});
    document.addEventListener('copy', function(e) {{ e.preventDefault(); }});
    document.addEventListener('cut', function(e) {{ e.preventDefault(); }});
    document.addEventListener('paste', function(e) {{ e.preventDefault(); }});
    </script>
    """
    components.html(timer_js, height=65)

# ==========================================
# 4. SIDEBAR NAVIGATION
# ==========================================
st.sidebar.title("🧭 Navigation")
selected_portal = st.sidebar.radio("Select Portal:", ["🎓 Student Examination & Academic Portal", "⚙️ Teacher / Admin Control Center"])
st.sidebar.divider()

# ==========================================
# 5. ADMIN CONTROL PANEL (TEACHER ONLY)
# ==========================================
if selected_portal == "⚙️ Teacher / Admin Control Center":
    if "admin_authenticated" not in st.session_state:
        st.session_state.admin_authenticated = False

    if not st.session_state.admin_authenticated:
        st.title("🔐 Teacher Login Portal")
        st.markdown("Authorized access only for Subject Teacher.")
        col1, _ = st.columns([1.2, 1])
        with col1:
            with st.form("admin_login_form"):
                in_user = st.text_input("Username:")
                in_pass = st.text_input("Password:", type="password")
                btn_login = st.form_submit_button("Sign In", type="primary")
                if btn_login:
                    if in_user.strip() == ADMIN_USERNAME and in_pass.strip() == ADMIN_PASSWORD:
                        st.session_state.admin_authenticated = True
                        st.success("Admin Login Successful!")
                        time.sleep(0.5)
                        st.rerun()
                    else:
                        st.error("Galat Username ya Password! Access Denied.")
        st.stop()

    st.sidebar.success(f"👑 Logged in as: `{ADMIN_USERNAME}`")
    if st.sidebar.button("Log Out Admin"):
        st.session_state.admin_authenticated = False
        st.rerun()

    st.title("⚙️ Teacher & Examination Control Center")
    st.info(f"🕒 Current Indian Standard Time (IST): **{get_ist_now().strftime('%Y-%m-%d %I:%M %p')}**")

    quizzes_df = get_all_quizzes()
    admin_tab = st.selectbox("Select Management Section:", [
        "📚 Create & Manage Quizzes (Class 11 & 12 Control)", 
        "👥 Master Student Directory, Attendance & Tests", 
        "📝 Question Bank (Excel/Manual)",
        "📊 Overall Merit List & Results (PDF Download)", 
        "📷 Student Photo Upload (Roll No Wise)",
        "💾 Full Database Backup & Restore"
    ])

    st.divider()
    conn = get_db()

    # SECTION 1: QUIZZES
    if admin_tab == "📚 Create & Manage Quizzes (Class 11 & 12 Control)":
        st.subheader("Quiz Schedule & Live Controls")
        with st.expander("➕ Create New Quiz"):
            with st.form("new_quiz_form"):
                c_cls1, c_cls2 = st.columns(2)
                target_class_choice = c_cls1.selectbox("Select Class:", ["Class 11", "Class 12", "Class 9", "Class 10", "Other"])
                topic_name = c_cls2.text_input("Topic / Chapter Name:", value="Kinematics")
                q_title = st.text_input("Quiz Title:", value=f"{target_class_choice} - {topic_name}")
                q_dur = st.number_input("Duration (Minutes):", min_value=1, max_value=300, value=15)
                
                c_d1, c_d2 = st.columns(2)
                cur_ist = get_ist_now()
                start_date = c_d1.date_input("Start Date (IST):", value=cur_ist.date())
                start_time = c_d1.time_input("Start Time (IST):", value=(cur_ist - timedelta(minutes=10)).time())
                end_date = c_d2.date_input("End Date (IST):", value=(cur_ist + timedelta(days=7)).date())
                end_time = c_d2.time_input("End Time (IST):", value=cur_ist.time())
                
                if st.form_submit_button("Create Quiz"):
                    start_str = f"{start_date} {start_time.strftime('%H:%M')}"
                    end_str = f"{end_date} {end_time.strftime('%H:%M')}"
                    c_title = clean_val(q_title)
                    c_top = clean_val(topic_name)
                    if c_title:
                        try:
                            conn.execute('''
                                INSERT INTO quizzes (target_class, topic, quiz_title, duration_minutes, start_datetime, end_datetime, is_active)
                                VALUES (?, ?, ?, ?, ?, ?, 0)
                            ''', (target_class_choice, c_top, c_title, q_dur, start_str, end_str))
                            conn.commit()
                            st.success(f"Quiz '{c_title}' ban gaya!")
                            time.sleep(1)
                            st.rerun()
                        except sqlite3.IntegrityError:
                            st.error("A quiz with this title already exists.")

        st.markdown("---")
        if not quizzes_df.empty:
            for _, r in quizzes_df.iterrows():
                with st.container():
                    st.markdown(f"### 📝 **{r['quiz_title']}**")
                    cls_val = r['target_class']
                    st.markdown(f"🏷️ **Class:** `{cls_val}` | 📖 **Topic:** `{r['topic']}` | ⏱️ `{r['duration_minutes']} Mins` | **Status:** `{'🟢 LIVE' if r['is_active'] == 1 else '⚪ Disabled'}`")
                    
                    col_b1, col_b2, col_b3 = st.columns([1.5, 1.5, 1])
                    if col_b1.button(f"Toggle Active ({r['quiz_title']})", key=f"tog_{r['id']}"):
                        new_status = 0 if r['is_active'] == 1 else 1
                        if new_status == 1:
                            conn.execute("UPDATE quizzes SET is_active = 0 WHERE target_class = ?", (cls_val,))
                        conn.execute("UPDATE quizzes SET is_active = ? WHERE id = ?", (new_status, r['id']))
                        conn.commit()
                        st.rerun()
                    
                    if col_b2.button(f"⚡ Start NOW (Instant Sole LIVE)", key=f"now_{r['id']}"):
                        now_start = (get_ist_now() - timedelta(hours=1)).strftime("%Y-%m-%d %H:%M")
                        now_end = (get_ist_now() + timedelta(days=10)).strftime("%Y-%m-%d %H:%M")
                        # Class-wise rule: Only 1 live per class
                        conn.execute("UPDATE quizzes SET is_active = 0 WHERE target_class = ?", (cls_val,))
                        conn.execute("UPDATE quizzes SET start_datetime = ?, end_datetime = ?, is_active = 1 WHERE id = ?", (now_start, now_end, r['id']))
                        conn.commit()
                        st.success(f"{cls_val} ke liye sirf yeh Quiz LIVE kar diya gaya hai!")
                        time.sleep(1)
                        st.rerun()
                    
                    if col_b3.button(f"🗑️ Delete Quiz", key=f"del_quiz_{r['id']}", type="secondary"):
                        conn.execute("DELETE FROM quizzes WHERE id = ?", (r['id'],))
                        conn.commit()
                        st.warning("Quiz delete ho gaya.")
                        time.sleep(1)
                        st.rerun()
                    st.divider()

    # SECTION 2: MASTER STUDENTS, ATTENDANCE & TEST MARKS (Admin View)
    elif admin_tab == "👥 Master Student Directory, Attendance & Tests":
        st.subheader("👥 Master Student Directory, Attendance & Monthly Tests")
        stu_df = pd.read_sql_query("SELECT id, roll_no, student_name, target_class, sr_no, attendance_present, attendance_total, attendance_pct, test_total, test_pct FROM master_students ORDER BY CAST(roll_no AS INTEGER) ASC", conn)
        
        st.dataframe(stu_df, use_container_width=True)
        
        st.markdown("#### ✏️ Update Individual Student Details:")
        if not stu_df.empty:
            with st.form("admin_edit_student"):
                sel_stu_roll = st.selectbox("Select Student:", stu_df['roll_no'].tolist(), format_func=lambda x: f"Roll {x} - {stu_df[stu_df['roll_no']==x]['student_name'].values[0] if len(stu_df[stu_df['roll_no']==x])>0 else x}")
                
                c_c1, c_c2, c_c3 = st.columns(3)
                ed_pres = c_c1.text_input("Present Days:", value="72")
                ed_tot = c_c2.text_input("Total Days:", value="87")
                ed_pct = c_c3.text_input("Attendance %:", value="82.7")
                
                c_c4, c_c5 = st.columns(2)
                ed_test_tot = c_c4.text_input("Monthly Test Total (/100):", value="85")
                ed_test_pct = c_c5.text_input("Monthly Test %:", value="85.0%")
                
                ed_st_goal = st.text_input("Short Term Goal:", value="Board exam me 90%+ marks.")
                ed_lt_goal = st.text_input("Long Term Goal:", value="Engineering / Higher Studies")
                
                if st.form_submit_button("Save Student Details"):
                    conn.execute('''
                        UPDATE master_students 
                        SET attendance_present=?, attendance_total=?, attendance_pct=?,
                            test_total=?, test_pct=?, short_term_goal=?, long_term_goal=?
                        WHERE roll_no=?
                    ''', (ed_pres, ed_tot, ed_pct, ed_test_tot, ed_test_pct, ed_st_goal, ed_lt_goal, sel_stu_roll))
                    conn.commit()
                    st.success("Student details updated!")
                    time.sleep(1)
                    st.rerun()

    # SECTION 3: QUESTION BANK
    elif admin_tab == "📝 Question Bank (Excel/Manual)":
        st.subheader("Manage Question Bank")
        if not quizzes_df.empty:
            quiz_options = {f"[{r['target_class']}] {r['quiz_title']}": r['id'] for _, r in quizzes_df.iterrows()}
            sel_q_label = st.selectbox("Select Quiz:", list(quiz_options.keys()))
            sel_q_id = quiz_options[sel_q_label]
            
            uploaded_q = st.file_uploader("Upload Questions (.xlsx / .csv):", type=["xlsx", "csv"])
            if uploaded_q and st.button("Import Questions"):
                df_q = pd.read_csv(uploaded_q) if uploaded_q.name.endswith(".csv") else pd.read_excel(uploaded_q)
                df_q.columns = [str(col).strip().lower().replace(" ", "_") for col in df_q.columns]
                cnt = 0
                for _, r in df_q.iterrows():
                    conn.execute('''
                        INSERT INTO questions (quiz_id, question, option_a, option_b, option_c, option_d, correct_option)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (sel_q_id, str(r["question"]).strip(), str(r["option_a"]).strip(), str(r["option_b"]).strip(), str(r["option_c"]).strip(), str(r["option_d"]).strip(), str(r["correct_option"]).strip()))
                    cnt += 1
                conn.commit()
                st.success(f"{cnt} questions imported!")
                st.rerun()

    # SECTION 4: OVERALL MERIT LIST & RESULTS (PDF Download)
    elif admin_tab == "📊 Overall Merit List & Results (PDF Download)":
        st.subheader("Official Merit List (Topper to Lower)")
        if not quizzes_df.empty:
            quiz_options = {f"[{r['target_class']}] {r['quiz_title']}": r['id'] for _, r in quizzes_df.iterrows()}
            sel_q_label = st.selectbox("Select Quiz:", list(quiz_options.keys()))
            sel_q_id = quiz_options[sel_q_label]
            
            q_info = dict(conn.execute("SELECT * FROM quizzes WHERE id = ?", (sel_q_id,)).fetchone())
            
            # Recalculate Previous Scores
            try:
                resp_rows = conn.execute('''
                    SELECT r.id, r.selected_option, r.correct_option, q.option_a, q.option_b, q.option_c, q.option_d, r.sr_no
                    FROM student_responses r JOIN questions q ON r.question_id = q.id WHERE r.quiz_id = ?
                ''', (sel_q_id,)).fetchall()
                for rr in resp_rows:
                    is_c = 1 if is_answer_correct(rr['selected_option'], rr['correct_option'], rr['option_a'], rr['option_b'], rr['option_c'], rr['option_d']) else 0
                    conn.execute("UPDATE student_responses SET is_correct = ? WHERE id = ?", (is_c, rr['id']))
                conn.execute('''
                    UPDATE submissions SET score = (
                        SELECT COALESCE(SUM(is_correct), 0) FROM student_responses 
                        WHERE student_responses.quiz_id = submissions.quiz_id AND student_responses.sr_no = submissions.sr_no
                    ) WHERE quiz_id = ?
                ''', (sel_q_id,))
                conn.commit()
            except Exception:
                pass
            
            subs_df = pd.read_sql_query("SELECT student_name, sr_no, score, total_questions, tab_switches, status, submitted_at FROM submissions WHERE quiz_id = ? ORDER BY score DESC, submitted_at ASC", conn, params=(sel_q_id,))
            
            if subs_df.empty:
                st.info("Abhi is exam ke liye koi submission nahi hai.")
            else:
                subs_disp = subs_df.copy()
                subs_disp.insert(0, "Rank", range(1, len(subs_disp) + 1))
                st.dataframe(subs_disp, use_container_width=True)
                
                c_p1, c_p2 = st.columns(2)
                with c_p1:
                    pdf_bytes = generate_merit_pdf(subs_df, q_info)
                    st.download_button(
                        label="📄 Download Official Merit List (PDF)",
                        data=pdf_bytes,
                        file_name=f"Merit_List_{q_info.get('quiz_title','Exam').replace(' ', '_')}.pdf",
                        mime="application/pdf",
                        type="primary"
                    )
                with c_p2:
                    st.download_button("📥 Download Results (CSV)", subs_disp.to_csv(index=False).encode('utf-8'), f"results_{sel_q_id}.csv", "text/csv")

    # SECTION 5: STUDENT PHOTO MANAGEMENT
    elif admin_tab == "📷 Student Photo Upload (Roll No Wise)":
        st.subheader("Student Photo Management (Bulk / Single)")
        st.info("💡 Photo ka filename Roll Number rakhein (e.g. `1.jpg`, `15.png`) taaki wo apne aap match ho jaye.")
        
        b_files = st.file_uploader("Select Multiple Student Photos:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
        if b_files and st.button("Match & Save Photos"):
            matched = 0
            for bf in b_files:
                num_match = re.search(r'\d+', bf.name)
                if num_match:
                    r_num = str(int(num_match.group(0)))
                    encoded = base64.b64encode(bf.read()).decode("utf-8")
                    conn.execute("UPDATE master_students SET photo_b64 = ? WHERE roll_no = ?", (encoded, r_num))
                    matched += 1
            conn.commit()
            st.success(f"{matched} vidyarthiyon ki photos successfully update ho gayi!")

    # SECTION 6: BACKUP
    elif admin_tab == "💾 Full Database Backup & Restore":
        st.subheader("💾 Full Database Backup")
        stu_exp = pd.read_sql_query("SELECT * FROM master_students", conn)
        q_exp = pd.read_sql_query("SELECT * FROM quizzes", conn)
        ques_exp = pd.read_sql_query("SELECT * FROM questions", conn)
        subs_exp = pd.read_sql_query("SELECT * FROM submissions", conn)
        resp_exp = pd.read_sql_query("SELECT * FROM student_responses", conn)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            stu_exp.to_excel(writer, sheet_name='Master_Students', index=False)
            q_exp.to_excel(writer, sheet_name='Quizzes', index=False)
            ques_exp.to_excel(writer, sheet_name='Questions', index=False)
            subs_exp.to_excel(writer, sheet_name='Submissions', index=False)
            resp_exp.to_excel(writer, sheet_name='Responses', index=False)
            
        st.download_button("📥 Download Full Backup (.xlsx)", output.getvalue(), f"Full_Database_Backup_{get_ist_now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    conn.close()

# ==========================================
# 6. STUDENT INDIVIDUAL PORTAL (RESTRICTED TO SELF)
# ==========================================
else:
    if "student_name" not in st.session_state:
        st.session_state.student_name = None
    if "student_sr" not in st.session_state:
        st.session_state.student_sr = None
    if "selected_class" not in st.session_state:
        st.session_state.selected_class = "Class 11"

    quizzes_df = get_all_quizzes()
    conn = get_db()

    # Student Login Screen
    if not st.session_state.student_name:
        st.subheader("🎓 Student Login Portal")
        st.markdown("Pehle apni **Class** chunein, phir apna **Registered Name** aur Password me apna **SR No** darj karein.")
        
        c1, _ = st.columns([1.2, 1])
        with c1:
            sel_cls = st.selectbox("Select Your Class:", ["Class 11", "Class 12"])
            with st.form("stu_login_form"):
                in_name = st.text_input("Student Name (Registered):", placeholder="Enter your full name")
                in_pwd = st.text_input("Password (Your SR No):", type="password")
                btn_stu_login = st.form_submit_button("Enter Portal", type="primary")
                
                if btn_stu_login:
                    c_name = clean_val(in_name)
                    c_pwd = clean_sr_no(in_pwd)
                    norm_name = c_name.lower()
                    
                    s_data = conn.execute("SELECT * FROM master_students WHERE normalized_name = ?", (norm_name,)).fetchone()
                    if not c_name or not c_pwd:
                        st.error("Kripya Naam aur Password dono darj karein.")
                    elif not s_data:
                        st.error(f"❌ Student Name '{c_name}' registered list me nahi mila! Kripya spelling check karein.")
                    elif clean_sr_no(s_data['sr_no']) != c_pwd:
                        st.error("Galat Password! (Password aapka SR Number hai).")
                    else:
                        st.session_state.student_name = s_data['student_name']
                        st.session_state.student_sr = clean_sr_no(s_data['sr_no'])
                        st.session_state.selected_class = sel_cls
                        st.rerun()
        conn.close()
        st.stop()

    # Logged In Student Session
    student_name = st.session_state.student_name
    student_sr = st.session_state.student_sr
    target_class = st.session_state.selected_class
    
    st.sidebar.markdown(f"**Candidate:** `{student_name}`")
    st.sidebar.markdown(f"**SR No:** `{student_sr}`")
    st.sidebar.markdown(f"**Class:** `{target_class}`")
    if st.sidebar.button("Log Out"):
        st.session_state.student_name = None
        st.session_state.student_sr = None
        st.rerun()

    # Strict Privacy: Fetch ONLY this student's data
    s_rec = conn.execute("SELECT * FROM master_students WHERE normalized_name = ?", (student_name.lower(),)).fetchone()
    s_dict = dict(s_rec) if s_rec else {}

    # Student Tabs
    stu_tabs = st.tabs([
        "📝 Live Exam", 
        "📊 My Exam Scorecard & Answer Key", 
        "📅 My Attendance & Tests", 
        "🎴 My UP Board Portfolio Card"
    ])

    # -------------------------------------------------------------
    # TAB 1: LIVE EXAM (Sole Active Quiz for Target Class)
    # -------------------------------------------------------------
    with stu_tabs[0]:
        class_live_quiz = quizzes_df[(quizzes_df['target_class'] == target_class) & (quizzes_df['is_active'] == 1)]
        
        if class_live_quiz.empty:
            st.info(f"🛑 {target_class} ke liye abhi koi exam LIVE nahi hai. Teacher dwara start karne par yahan display hoga.")
        else:
            q_row = class_live_quiz.iloc[0]
            quiz_id = q_row['id']
            quiz_title_val = q_row['quiz_title']
            quiz_topic_val = q_row['topic']
            quiz_dur_val = int(q_row['duration_minutes'])
            
            st.markdown(f"### 📝 {quiz_title_val}")
            st.markdown(f"📖 **Topic:** `{quiz_topic_val}` | ⏱️ **Duration:** `{quiz_dur_val} Minutes`")

            sub_check = conn.execute("SELECT * FROM submissions WHERE quiz_id = ? AND LOWER(student_name) = ?", (quiz_id, student_name.lower())).fetchone()
            
            if sub_check:
                st.success(f"✅ {student_name}, aapka yeh exam pehle hi successfully submit ho chuka hai!")
                st.metric("Aapka Score", f"{sub_check['score']} / {sub_check['total_questions']}")
                st.info("👉 Question-wise scorecard dekhne ke liye upar **'My Exam Scorecard & Answer Key'** tab dekhein.")
            else:
                questions_df = get_questions_by_quiz(quiz_id)
                if questions_df.empty:
                    st.info("Is quiz me abhi questions upload nahi hue hain.")
                else:
                    norm_name = student_name.lower()
                    attempt_row = conn.execute("SELECT start_epoch FROM quiz_attempts WHERE quiz_id = ? AND normalized_name = ?", (quiz_id, norm_name)).fetchone()
                    
                    if not attempt_row:
                        st.markdown(f"""
                        **📌 Exam Rules:**
                        1. **'Start Exam Now'** par click karte hi **{quiz_dur_val} minute** ka countdown timer shuru hoga.
                        2. Page refresh ya close karne par bhi timer background me chalta rahega.
                        3. Tab switch karne par warning aayegi aur 3 tab switches par auto-submit ho jayega.
                        """)
                        if st.button("🚀 Start Exam Now", type="primary"):
                            get_or_set_attempt_start(quiz_id, norm_name)
                            st.rerun()
                    else:
                        attempt_start = attempt_row["start_epoch"]
                        elapsed = time.time() - attempt_start
                        remaining = (quiz_dur_val * 60) - elapsed

                        if remaining <= 0:
                            sub_time = get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
                            conn.execute('''
                                INSERT OR REPLACE INTO submissions (quiz_id, student_name, sr_no, score, total_questions, tab_switches, status, submitted_at)
                                VALUES (?, ?, ?, 0, ?, 0, 'Auto-Submitted (Time Up)', ?)
                            ''', (quiz_id, student_name, student_sr, len(questions_df), sub_time))
                            conn.commit()
                            st.error("⏰ Time Up! Samay samapt ho gaya hai.")
                            st.rerun()

                        inject_live_timer_and_security(remaining, quiz_id, student_name)

                        with st.form("exam_form"):
                            answers = {}
                            for idx, row in questions_df.iterrows():
                                st.markdown(f"**Q{idx+1}. {row['question']}**")
                                opts = [row['option_a'], row['option_b'], row['option_c'], row['option_d']]
                                answers[row['id']] = st.radio("Option:", opts, key=f"q_{row['id']}", index=None)
                                st.markdown("---")
                                
                            if st.form_submit_button("Submit Final Answers", type="primary"):
                                sub_time = get_ist_now().strftime("%Y-%m-%d %H:%M:%S")
                                score = 0
                                for _, row in questions_df.iterrows():
                                    qid_num = row['id']
                                    sel_opt = answers.get(qid_num)
                                    correct_opt = row['correct_option']
                                    is_corr = 1 if is_answer_correct(sel_opt, correct_opt, row['option_a'], row['option_b'], row['option_c'], row['option_d']) else 0
                                    if is_corr:
                                        score += 1
                                    conn.execute('''
                                        INSERT INTO student_responses (quiz_id, student_name, sr_no, question_id, question_text, selected_option, correct_option, is_correct, recorded_at)
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                                    ''', (quiz_id, student_name, student_sr, qid_num, row['question'], sel_opt if sel_opt else "Unattempted", correct_opt, is_corr, sub_time))
                                
                                conn.execute('''
                                    INSERT OR REPLACE INTO submissions (quiz_id, student_name, sr_no, score, total_questions, tab_switches, status, submitted_at)
                                    VALUES (?, ?, ?, score, ?, 0, 'Completed', ?)
                                ''', (quiz_id, student_name, student_sr, score, len(questions_df), sub_time))
                                conn.commit()
                                st.balloons()
                                st.success(f"🎉 Exam Successfully Submitted! Score: {score}/{len(questions_df)}")
                                time.sleep(1)
                                st.rerun()

    # -------------------------------------------------------------
    # TAB 2: MY EXAM SCORECARD & DETAILED ANSWER KEY (Strictly Self)
    # -------------------------------------------------------------
    with stu_tabs[1]:
        st.subheader("📋 My Physics Exam Responses & Detailed Answer Key")
        sub_history = pd.read_sql_query("SELECT * FROM submissions WHERE LOWER(student_name) = ? ORDER BY id DESC", conn, params=(student_name.lower(),))
        
        if sub_history.empty:
            st.info("Aapne abhi tak koi test submit nahi kiya hai.")
        else:
            for _, sh in sub_history.iterrows():
                q_meta = conn.execute("SELECT quiz_title, topic FROM quizzes WHERE id = ?", (sh['quiz_id'],)).fetchone()
                q_tit = q_meta['quiz_title'] if q_meta else "Exam"
                
                with st.expander(f"📝 {q_tit} — Score: {sh['score']}/{sh['total_questions']}", expanded=True):
                    c_m1, c_m2, c_m3 = st.columns(3)
                    pct = (sh['score'] / sh['total_questions'] * 100) if sh['total_questions'] > 0 else 0
                    c_m1.metric("Final Score", f"{sh['score']} / {sh['total_questions']}")
                    c_m2.metric("Percentage", f"{pct:.1f}%")
                    c_m3.metric("Tab Switches Recorded", f"{sh['tab_switches']} times")

                    st_res = pd.read_sql_query("SELECT question_text, selected_option, correct_option, is_correct FROM student_responses WHERE quiz_id = ? AND LOWER(student_name) = ?", conn, params=(sh['quiz_id'], student_name.lower()))
                    for idx, r_row in st_res.iterrows():
                        is_r = (r_row['is_correct'] == 1)
                        b_col = "#28a745" if is_r else "#dc3545"
                        st.markdown(f"""
                        <div style="border-left: 5px solid {b_col}; padding: 10px 14px; margin-bottom: 10px; background: #f8fafc; border-radius: 6px;">
                            <b>Q{idx+1}. {r_row['question_text']}</b><br>
                            <span>Aapka Chuna Hua Option: <b>{r_row['selected_option']}</b> ({'✅ Sahi' if is_r else '❌ Galat'})</span><br>
                            <span style="color: #1e7e34; font-weight: bold;">Sahi Option: {r_row['correct_option']}</span>
                        </div>
                        """, unsafe_allow_html=True)

    # -------------------------------------------------------------
    # TAB 3: MY ATTENDANCE & UNIT TESTS (Strictly Self)
    # -------------------------------------------------------------
    with stu_tabs[2]:
        st.subheader("📅 Official School Attendance & Academic Test Record")
        
        col_at1, col_at2 = st.columns(2)
        with col_at1:
            st.markdown("#### 📊 उपस्थिति विवरण (Attendance)")
            st.metric("Total Working Days", f"{s_dict.get('attendance_total', '87')} Days")
            st.metric("Present Days", f"{s_dict.get('attendance_present', '72')} Days")
            st.metric("Annual Attendance %", f"{s_dict.get('attendance_pct', '82.5')}%")
            
        with col_at2:
            st.markdown("#### 📝 मासिक टेस्ट रिकॉर्ड (Monthly Test)")
            st.metric("Total Score (/100)", f"{s_dict.get('test_total', '-')}")
            st.metric("Test Percentage (%)", f"{s_dict.get('test_pct', '-')}")
            
        st.markdown("---")
        st.markdown(f"**📌 Short Term Goal:** {s_dict.get('short_term_goal', 'कक्षा 12वीं में 90%+ अंक अर्जित करना।')}")
        st.markdown(f"**🎯 Long Term Goal:** {s_dict.get('long_term_goal', 'उच्च शिक्षा एवं प्रतियोगी परीक्षाओं में सफलता प्राप्त करना।')}")

    # -------------------------------------------------------------
    # TAB 4: MY UP BOARD PORTFOLIO CARD (View & Download)
    # -------------------------------------------------------------
    with stu_tabs[3]:
        st.subheader("🎴 2-Page UP Board Student Portfolio Card")
        
        s_photo = s_dict.get("photo_b64", "")
        if safe_b64_decode(s_photo):
            photo_html = f'<img src="data:image/jpeg;base64,{s_photo}" style="width: 95px; height: 115px; object-fit: cover; border-radius: 6px; border: 2px solid #1E3A8A;"/>'
        else:
            photo_html = '<div style="font-size: 38px;">🎓</div><div style="font-size: 11px; color: #94A3B8;">फोटो प्रतीक्षित</div>'

        port_html = f"""<!DOCTYPE html>
        <html>
        <head><meta charset="utf-8">
        <style>
            body {{ font-family: Arial, sans-serif; background: #f8fafc; padding: 10px; color: #1e293b; }}
            .page {{ max-width: 800px; margin: 0 auto; background: white; border: 2px solid #1E3A8A; border-radius: 8px; padding: 20px; }}
        </style>
        </head>
        <body>
            <div class="page">
                <div style="text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 10px; margin-bottom: 15px;">
                    <h2 style="margin: 0; color: #1E3A8A;">ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT</h2>
                    <h4 style="margin: 4px 0 0 0; color: #059669;">विद्यार्थी पोर्टफोलियो एवं सतत आंतरिक मूल्यांकन रिकॉर्ड</h4>
                    <div style="font-size: 12px; color: #475569;">सत्र: 2026 - 2027 | कक्षा: {target_class}</div>
                </div>
                <div style="display: flex; gap: 15px;">
                    <div style="flex: 3; font-size: 13px; line-height: 1.8;">
                        <b>नाम:</b> {s_dict.get('student_name')}<br>
                        <b>अनुक्रमांक (Roll No):</b> {s_dict.get('roll_no')}<br>
                        <b>S.R. No:</b> {s_dict.get('sr_no')}<br>
                        <b>उपस्थिति (Attendance):</b> {s_dict.get('attendance_present', '72')}/{s_dict.get('attendance_total', '87')} दिन ({s_dict.get('attendance_pct', '82.5')}%)<br>
                        <b>मासिक टेस्ट प्राप्तांक:</b> {s_dict.get('test_total', '-')}/100 ({s_dict.get('test_pct', '-')})<br>
                        <b>अल्पकालिक लक्ष्य:</b> {s_dict.get('short_term_goal', 'सत्र में उत्कृष्ट अंक अर्जित करना')}<br>
                        <b>दीर्घकालिक लक्ष्य:</b> {s_dict.get('long_term_goal', 'उच्च शिक्षा एवं प्रतियोगी परीक्षाओं में सफलता')}
                    </div>
                    <div style="flex: 1; text-align: center; border: 1px dashed #CBD5E1; padding: 8px; border-radius: 6px;">
                        {photo_html}
                    </div>
                </div>
            </div>
        </body></html>"""
        
        st.download_button(
            label="📥 Download My Official Portfolio Card (.html)",
            data=port_html,
            file_name=f"Portfolio_{s_dict.get('student_name')}.html",
            mime="text/html"
        )
        st.components.v1.html(port_html, height=450, scrolling=True)

    conn.close()
