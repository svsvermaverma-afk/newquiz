import streamlit as st
import sqlite3
import pandas as pd
import time
import io
import os
import glob
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
    page_title="ABIC Renukoot - Comprehensive Portal",
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
        font-size: 1.85rem !important;
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
        .school-header h1 { font-size: 1.25rem !important; }
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

SCHOOL_NAME_HEADER = "ADITYA BIRLA INTERMEDIATE COLLEGE, RENUKOOT, SONEBHADRA (UP)"

st.markdown(f"""
<div class="school-header">
    <h1>{SCHOOL_NAME_HEADER}</h1>
    <h3>⚡ Physics Exam, Academic Analytics & UP Board Portfolio Portal ⚡</h3>
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

def convert_gdrive_link(url):
    if not url or not isinstance(url, str):
        return ""
    url = url.strip()
    match1 = re.search(r'/d/([a-zA-Z0-9_-]+)', url)
    if match1:
        return f"https://lh3.googleusercontent.com/d/{match1.group(1)}"
    match2 = re.search(r'id=([a-zA-Z0-9_-]+)', url)
    if match2:
        return f"https://lh3.googleusercontent.com/d/{match2.group(1)}"
    return url

# 14 Official School Activities (UP Board Calendar 2026-27)
DEFAULT_ACTIVITIES = [
    {"sno": 1, "date": "27.08.2026", "name": "Tata Building India School Essay Competition", "cat": "साहित्यिक (निबंध)", "desc": "2047 तक भारत को विश्व का सबसे विकसित देश बनाने के लिए मैं यह पांच कार्य करूंगा/करूंगी", "incharge": "श्री विकास कुमार चक्रवर्ती / कक्षा अध्यापक"},
    {"sno": 2, "date": "27.08.2026", "name": "रंगोली प्रतियोगिता", "cat": "कला एवं संस्कृति", "desc": "रंगोली निर्माण (समूह गतिविधि - प्रति समूह 4 विद्यार्थी)", "incharge": "श्रीमती साधना भरद्वाज"},
    {"sno": 3, "date": "27.08.2026", "name": "मेहंदी प्रतियोगिता", "cat": "कला एवं संस्कृति", "desc": "मेहंदी आलेखन (रचनात्मकता, मौलिकता व बारीकी)", "incharge": "श्रीमती पूजा सिंह"},
    {"sno": 4, "date": "20.08.2026", "name": "राखी निर्माण प्रतियोगिता", "cat": "क्राफ्ट एवं रचनात्मक कौशल", "desc": "आकर्षक व सुंदर राखी निर्माण (राखी प्रदर्शनी हेतु)", "incharge": "श्री शशिकांत सर / श्री विकास कुमार चक्रवर्ती"},
    {"sno": 5, "date": "13.08.2026", "name": "चित्रकला प्रतियोगिता", "cat": "दृश्य कला (Drawing)", "desc": "सरदार वल्लभभाई पटेल के जीवन एवं आदर्शों पर आधारित चित्रकला", "incharge": "डॉ. संतोष कुमार तिवारी"},
    {"sno": 6, "date": "06.08.2026", "name": "निबंध प्रतियोगिता", "cat": "साहित्यिक (निबंध)", "desc": "सरदार वल्लभभाई पटेल की 150वीं जयंती पर उनके जीवन, आदर्श व मूल्यों पर निबंध", "incharge": "डॉ. बबलू कुमार भट्ट"},
    {"sno": 7, "date": "31.07.2026", "name": "बाल संसद (Student Council)", "cat": "नेतृत्व कौशल (Leadership)", "desc": "बाल संसद पदाधिकारियों का शपथ ग्रहण समारोह", "incharge": "विद्यालय प्रशासन / हिंडालको प्रबंधन"},
    {"sno": 8, "date": "30.07.2026", "name": "कक्षा सज्जा एवं शैक्षणिक चार्ट प्रतियोगिता", "cat": "रचनात्मक एवं शैक्षणिक कौशल", "desc": "कक्षा कक्ष सौंदर्यीकरण एवं शिक्षण-अधिगम चार्ट निर्माण", "incharge": "कक्षा अध्यापक / श्री विकास कुमार चक्रवर्ती"},
    {"sno": 9, "date": "23.07.2026", "name": "Elocution (भाषण प्रतियोगिता)", "cat": "साहित्यिक (मौखिक अभिव्यक्ति)", "desc": "विषय: अनुशासन का महत्व, प्रिय कवि, आतंकवाद, स्वतंत्रता दिवस, बेरोजगारी", "incharge": "श्री शशिकांत मौर्या"},
    {"sno": 10, "date": "16.07.2026", "name": "Story Telling (कहानी लेखन)", "cat": "साहित्यिक (रचनात्मक लेखन)", "desc": "विषय: 'The Power of Honesty'", "incharge": "श्री वशिष्ठ राकेश कुमार"},
    {"sno": 11, "date": "09.07.2026", "name": "IEP पोस्टर प्रतियोगिता", "cat": "कला एवं पर्यावरण जागरूकता", "desc": "विषय: पर्यावरण संरक्षण / सड़क सुरक्षा (चार्ट पेपर पोस्टर)", "incharge": "श्री विकास कुमार चक्रवर्ती"},
    {"sno": 12, "date": "02.07.2026", "name": "ABG Group Orchestra प्रतियोगिता", "cat": "प्रदर्शन कला (संगीत)", "desc": "वाद्य यंत्र / संगीत प्रदर्शन (ऑर्केस्ट्रा)", "incharge": "श्रीमती ज्योति मिश्रा"},
    {"sno": 13, "date": "02.07.2026", "name": "लेख प्रतियोगिता (Article Writing)", "cat": "सामाजिक जागरूकता / वैचारिक लेखन", "desc": "विषय: 'जनगणना का महत्व तथा आवश्यकता'", "incharge": "कक्षा अध्यापक / एक्टिविटी प्रभारी"},
    {"sno": 14, "date": "14.05.2026", "name": "Creative Story Writing Competition", "cat": "साहित्यिक (अंग्रेजी लेखन)", "desc": "English Story Writing (Thinking & Writing Skills)", "incharge": "श्री अशोक द्विवेदी"}
]

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
    
    # 1. Master Students Table (Integrated Profile)
    c.execute('''
        CREATE TABLE IF NOT EXISTS master_students (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no TEXT,
            student_name TEXT NOT NULL,
            student_name_hindi TEXT DEFAULT '',
            sr_no TEXT NOT NULL,
            normalized_name TEXT UNIQUE NOT NULL,
            target_class TEXT DEFAULT 'Class 12',
            roll_no_10th TEXT DEFAULT '',
            pen_no TEXT DEFAULT '',
            dob TEXT DEFAULT '',
            father_name TEXT DEFAULT '',
            father_name_hindi TEXT DEFAULT '',
            mother_name TEXT DEFAULT '',
            mother_name_hindi TEXT DEFAULT '',
            gender TEXT DEFAULT '',
            category TEXT DEFAULT '',
            mob_no TEXT DEFAULT '',
            email_id TEXT DEFAULT '',
            address TEXT DEFAULT '',
            occupation TEXT DEFAULT '-',
            ecode TEXT DEFAULT '-',
            dept TEXT DEFAULT '-',
            caste TEXT DEFAULT '-',
            religion TEXT DEFAULT '-',
            attendance_pct TEXT DEFAULT '82.5',
            attendance_present TEXT DEFAULT '72',
            attendance_total TEXT DEFAULT '87',
            test_hindi TEXT DEFAULT '',
            test_eng TEXT DEFAULT '',
            test_maths TEXT DEFAULT '',
            test_phy TEXT DEFAULT '',
            test_che TEXT DEFAULT '',
            test_total TEXT DEFAULT '',
            test_pct TEXT DEFAULT '',
            short_term_goal TEXT DEFAULT '',
            long_term_goal TEXT DEFAULT '',
            academic_goals TEXT DEFAULT '',
            strengths_weaknesses TEXT DEFAULT '',
            photo_b64 TEXT DEFAULT '',
            photo_url TEXT DEFAULT ''
        )
    ''')

    # Safe Schema Migrations
    c.execute("PRAGMA table_info(master_students)")
    cols = [info[1] for info in c.fetchall()]
    new_cols = [
        ("target_class", "TEXT DEFAULT 'Class 12'"),
        ("student_name_hindi", "TEXT DEFAULT ''"),
        ("roll_no_10th", "TEXT DEFAULT ''"),
        ("pen_no", "TEXT DEFAULT ''"),
        ("father_name_hindi", "TEXT DEFAULT ''"),
        ("mother_name", "TEXT DEFAULT ''"),
        ("mother_name_hindi", "TEXT DEFAULT ''"),
        ("gender", "TEXT DEFAULT ''"),
        ("category", "TEXT DEFAULT ''"),
        ("mob_no", "TEXT DEFAULT ''"),
        ("email_id", "TEXT DEFAULT ''"),
        ("address", "TEXT DEFAULT ''"),
        ("occupation", "TEXT DEFAULT '-'"),
        ("ecode", "TEXT DEFAULT '-'"),
        ("dept", "TEXT DEFAULT '-'"),
        ("caste", "TEXT DEFAULT '-'"),
        ("religion", "TEXT DEFAULT '-'"),
        ("attendance_present", "TEXT DEFAULT '72'"),
        ("attendance_total", "TEXT DEFAULT '87'"),
        ("attendance_pct", "TEXT DEFAULT '82.5'"),
        ("test_hindi", "TEXT DEFAULT ''"),
        ("test_eng", "TEXT DEFAULT ''"),
        ("test_maths", "TEXT DEFAULT ''"),
        ("test_phy", "TEXT DEFAULT ''"),
        ("test_che", "TEXT DEFAULT ''"),
        ("test_total", "TEXT DEFAULT ''"),
        ("test_pct", "TEXT DEFAULT ''"),
        ("short_term_goal", "TEXT DEFAULT ''"),
        ("long_term_goal", "TEXT DEFAULT ''"),
        ("academic_goals", "TEXT DEFAULT ''"),
        ("strengths_weaknesses", "TEXT DEFAULT ''"),
        ("photo_b64", "TEXT DEFAULT ''"),
        ("photo_url", "TEXT DEFAULT ''")
    ]
    for col_name, col_type in new_cols:
        if col_name not in cols:
            c.execute(f"ALTER TABLE master_students ADD COLUMN {col_name} {col_type}")

    # 2. Quizzes Table
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

    # 3. Questions Table
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

    # 4. Submissions Table
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

    # 5. Question Responses Table
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

    # 6. Persistent Attempt Timers
    c.execute('''
        CREATE TABLE IF NOT EXISTS quiz_attempts (
            quiz_id INTEGER NOT NULL,
            normalized_name TEXT NOT NULL,
            start_epoch REAL NOT NULL,
            PRIMARY KEY(quiz_id, normalized_name)
        )
    ''')

    # 7. Portfolio Activities Entries
    c.execute('''
        CREATE TABLE IF NOT EXISTS portfolio_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            roll_no TEXT,
            activity_name TEXT NOT NULL,
            category TEXT,
            activity_date TEXT,
            student_description TEXT,
            student_reflection TEXT,
            evidence_link TEXT,
            marks_awarded INTEGER DEFAULT 5,
            teacher_remarks TEXT DEFAULT 'उत्कृष्ट सहभागिता',
            submitted_on TEXT,
            UNIQUE(roll_no, activity_name) ON CONFLICT REPLACE
        )
    ''')

    # Default Quizzes
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

    # Auto load initial students if available
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

# DB Helpers
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
# 3. 2-PAGE UP BOARD PORTFOLIO CARD HTML
# ==========================================
def generate_upboard_card(student, entries_df):
    s_photo = student.get("photo_b64", "")
    p_url = student.get("photo_url", "")
    
    if safe_b64_decode(s_photo):
        photo_html = f'<img src="data:image/jpeg;base64,{s_photo}" style="width: 95px; height: 115px; object-fit: cover; border-radius: 6px; border: 2px solid #1E3A8A;"/>'
    elif p_url:
        photo_html = f'<img src="{p_url}" style="width: 95px; height: 115px; object-fit: cover; border-radius: 6px; border: 2px solid #1E3A8A;" onerror="this.style.display=\'none\';"/>'
    else:
        photo_html = '<div style="font-size: 42px;">🎓</div><div style="font-size: 11px; color: #94A3B8;">फोटो प्रतीक्षित</div>'

    activities_rows = ""
    if entries_df.empty:
        for act in DEFAULT_ACTIVITIES[:5]:
            activities_rows += f"""
            <tr style="border-bottom: 1px solid #E2E8F0; font-size: 11.5px;">
                <td style="padding: 6px; text-align: center;">{act['date']}</td>
                <td style="padding: 6px; font-weight: 600; color: #1E3A8A;">{act['name']}<br><span style="font-weight: normal; color: #64748B; font-size: 10.5px;">{act['desc']}</span></td>
                <td style="padding: 6px; text-align: center;">{act['cat']}</td>
                <td style="padding: 6px; color: #334155;">सक्रिय प्रतिभागिता एवं उत्तम प्रदर्शन</td>
                <td style="padding: 6px; text-align: center; font-weight: bold; color: #059669;">5/5</td>
            </tr>
            """
    else:
        for _, itm in entries_df.iterrows():
            reflection = itm['student_reflection'] if clean_val(itm['student_reflection']) else "सक्रिय सहभागिता एवं व्यावहारिक अनुभव।"
            desc = itm['student_description'] if clean_val(itm['student_description']) else "गतिविधि में योगदान"
            marks = itm['marks_awarded'] if itm['marks_awarded'] else 5
            link_badge = f'<br><a href="{itm["evidence_link"]}" target="_blank" style="font-size:11px; color:#2563EB;">🔗 फोटो लिंक</a>' if itm.get("evidence_link") else ''

            activities_rows += f"""
            <tr style="border-bottom: 1px solid #E2E8F0; font-size: 11.5px;">
                <td style="padding: 6px; text-align: center;">{itm['activity_date']}</td>
                <td style="padding: 6px; font-weight: 600; color: #1E3A8A;">{itm['activity_name']}<br><span style="font-weight: normal; color: #475569; font-size: 10.5px;">{desc}</span></td>
                <td style="padding: 6px; text-align: center;">{itm['category']}</td>
                <td style="padding: 6px; color: #0284C7; font-style: italic;">{reflection}{link_badge}</td>
                <td style="padding: 6px; text-align: center; font-weight: bold; color: #059669;">{marks}/5</td>
            </tr>
            """

    today_str = datetime.now().strftime('%d-%m-%Y')
    hindi_name = f"({student.get('student_name_hindi')})" if student.get('student_name_hindi') else ""
    short_term = student.get('short_term_goal', '').strip()
    long_term = student.get('long_term_goal', '').strip()
    general_goals = student.get('academic_goals', '').strip()

    raw_pct = student.get('attendance_pct', '')
    raw_pres = student.get('attendance_present', '')
    raw_tot = student.get('attendance_total', '87')

    try:
        pct_float = float(str(raw_pct).replace('%', '').strip())
        display_pct = f"{pct_float:.1f}%"
    except Exception:
        display_pct = f"{raw_pct}%" if raw_pct else "82.5%"

    pct_num_match = re.findall(r'\d+\.?\d*', display_pct)
    pct_val = float(pct_num_match[0]) if pct_num_match else 80.0
    status_label = "✅ संतोषजनक (>=75%)" if pct_val >= 75.0 else "⚠️ ध्यान देने योग्य (<75%)"
    status_color = "#059669" if pct_val >= 75.0 else "#DC2626"

    m_hin = student.get('test_hindi', '')
    m_eng = student.get('test_eng', '')
    m_mat = student.get('test_maths', '')
    m_phy = student.get('test_phy', '')
    m_che = student.get('test_che', '')
    m_tot = student.get('test_total', '')
    m_pct = student.get('test_pct', '')

    test_display_tot = f"{float(m_tot):.0f}" if m_tot and re.match(r'^\d+(\.\d+)?$', str(m_tot)) else (m_tot if m_tot else "-")
    test_display_pct = f"{float(m_pct):.1f}%" if m_pct and re.match(r'^\d+(\.\d+)?$', str(m_pct)) else (f"{m_pct}%" if m_pct else "-")

    if not short_term and not long_term:
        vision_html = f"""
        <div style="background: #F8FAFC; border-left: 4px solid #3B82F6; padding: 10px 14px; border-radius: 4px; font-size: 13px; color: #334155; line-height: 1.5;">
            {general_goals if general_goals else "सत्र 2026-27 में बोर्ड परीक्षा में उत्कृष्ट अंक अर्जित करना तथा नियमित अध्ययन करना।"}
        </div>
        """
    else:
        st_text = short_term if short_term else "कक्षा 12वीं में 90%+ अंक अर्जित करना तथा विषयों में प्रवीणता प्राप्त करना।"
        lt_text = long_term if long_term else "उच्च शिक्षा एवं प्रतियोगी परीक्षाओं में सफलता प्राप्त करना।"
        vision_html = f"""
        <div style="display: flex; gap: 12px; margin-top: 5px;">
            <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #3B82F6; padding: 8px 12px; border-radius: 4px; font-size: 12.5px; color: #1e293b;">
                <strong style="color: #1E3A8A;">📌 अल्पकालिक लक्ष्य (Short-Term Goal 2026-27):</strong><br>{st_text}
            </div>
            <div style="flex: 1; background: #F8FAFC; border-left: 4px solid #059669; padding: 8px 12px; border-radius: 4px; font-size: 12.5px; color: #1e293b;">
                <strong style="color: #059669;">🎯 दीर्घकालिक लक्ष्य (Long-Term Goal - Career):</strong><br>{lt_text}
            </div>
        </div>
        """

    sw = student.get('strengths_weaknesses') if student.get('strengths_weaknesses') else "ताकत: परिश्रम व अनुशासन | सुधार क्षेत्र: समय प्रबंधन।"

    return f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>Portfolio - {student.get('student_name')}</title>
    <style>
        body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f8fafc; padding: 15px; color: #1e293b; }}
        .page {{ max-width: 850px; margin: 0 auto 25px auto; background: #ffffff; border: 2px solid #1E3A8A; border-radius: 10px; padding: 25px; box-shadow: 0 4px 10px rgba(0,0,0,0.06); }}
        @media print {{ body {{ background: none; padding: 0; }} .page {{ box-shadow: none; margin: 0; border: 2px solid #000; page-break-after: always; }} }}
    </style>
</head>
<body>
    <!-- PAGE 1 -->
    <div class="page">
        <div style="text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 12px; margin-bottom: 18px;">
            <h2 style="margin: 0; color: #1E3A8A; font-size: 20px; text-transform: uppercase; letter-spacing: 1px;">{SCHOOL_NAME_HEADER}</h2>
            <h3 style="margin: 4px 0 0 0; color: #059669; font-size: 16px;">छात्र पोर्टफोलियो एवं सतत आंतरिक मूल्यांकन रिकॉर्ड</h3>
            <div style="font-size: 13px; color: #475569; margin-top: 4px;">सत्र: 2026 - 2027 | कक्षा: {student.get('target_class', 'Class 12')}</div>
            <div style="display: inline-block; background: #1E3A8A; color: white; padding: 3px 14px; border-radius: 12px; font-size: 11px; margin-top: 6px; font-weight: 600;">भाग 1 : व्यक्तिगत विवरण एवं स्व-मूल्यांकन</div>
        </div>

        <div style="display: flex; gap: 15px; margin-bottom: 15px;">
            <table style="width: 72%; border-collapse: collapse; font-size: 13px;">
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold; width: 35%;">छात्र/छात्रा का नाम:</td><td style="padding: 6px; color: #1E3A8A; font-weight: bold; font-size: 14px;">{student.get('student_name')} {hindi_name}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">अनुक्रमांक (Roll No.):</td><td style="padding: 6px; font-weight: bold;">{student.get('roll_no')}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold;">S.R. No. / PEN:</td><td style="padding: 6px;">{student.get('sr_no')} / {student.get('pen_no')}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">पिता का नाम:</td><td style="padding: 6px;">{student.get('father_name')}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold;">माता का नाम:</td><td style="padding: 6px;">{student.get('mother_name')}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">जन्म तिथि (D.O.B.):</td><td style="padding: 6px;">{student.get('dob')}</td></tr>
                <tr style="background: #F1F5F9;"><td style="padding: 6px; font-weight: bold;">संपर्क सूत्र (Mobile):</td><td style="padding: 6px;">{student.get('mob_no')}</td></tr>
                <tr><td style="padding: 6px; font-weight: bold;">निवास पता:</td><td style="padding: 6px;">{student.get('address')}</td></tr>
            </table>
            <div style="width: 28%; border: 2px dashed #94A3B8; border-radius: 8px; display: flex; flex-direction: column; align-items: center; justify-content: center; background: #F8FAFC; padding: 10px; text-align: center;">
                {photo_html}
                <div style="font-weight: bold; font-size: 13px; color: #1E3A8A; margin-top: 6px;">{student.get('student_name')}</div>
                <div style="font-size: 11px; color: #64748B;">कक्षा: {student.get('target_class', 'Class 12')}</div>
                <div style="font-size: 10px; color: #059669; margin-top: 4px; border: 1px solid #059669; padding: 2px 6px; border-radius: 8px;">सत्यापित विद्यार्थी</div>
            </div>
        </div>

        <div style="background: #F1F5F9; border: 1px solid #CBD5E1; border-radius: 6px; padding: 10px 14px; margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; align-items: center;">
                <div style="font-size: 13px; font-weight: bold; color: #1E3A8A;">📊 सत्र 2026-27 उपस्थिति विवरण (Official Attendance Record):</div>
                <div style="font-size: 12.5px; font-weight: bold; color: {status_color};">{status_label}</div>
            </div>
            <div style="display: flex; gap: 20px; margin-top: 6px; font-size: 12.5px; color: #334155;">
                <div><strong>कुल कार्य दिवस:</strong> {raw_tot}</div>
                <div><strong>उपस्थित दिवस:</strong> {raw_pres if raw_pres else 'N/A'}</div>
                <div><strong>वार्षिक उपस्थिति %:</strong> <span style="font-weight: bold; color: {status_color}; font-size: 13.5px;">{display_pct}</span></div>
            </div>
        </div>

        <div style="margin-top: 10px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 14px; margin-bottom: 6px;">🎯 शैक्षणिक लक्ष्य एवं संकल्प (Academic Vision & Career Goals):</div>
            {vision_html}
        </div>

        <div style="margin-top: 15px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 14px; margin-bottom: 6px;">💡 क्षमताएं एवं सुधार क्षेत्र (Self-Reflection):</div>
            <div style="background: #F8FAFC; border-left: 4px solid #10B981; padding: 10px 14px; border-radius: 4px; font-size: 13px; color: #334155; line-height: 1.5;">{sw}</div>
        </div>
    </div>

    <!-- PAGE 2 -->
    <div class="page">
        <div style="text-align: center; border-bottom: 2px solid #E2E8F0; padding-bottom: 10px; margin-bottom: 12px;">
            <h2 style="margin: 0; color: #1E3A8A; font-size: 18px; text-transform: uppercase;">{SCHOOL_NAME_HEADER}</h2>
            <h3 style="margin: 3px 0 0 0; color: #059669; font-size: 15px;">मासिक परीक्षा मूल्यांकन एवं सह-पाठ्यचर्या गतिविधि प्रपत्र</h3>
            <div style="display: inline-block; background: #059669; color: white; padding: 2px 14px; border-radius: 12px; font-size: 11px; margin-top: 4px; font-weight: 600;">भाग 2 : मासिक परीक्षा परिणाम, गतिविधियां व रूब्रिक्स</div>
        </div>

        <div style="margin-bottom: 14px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 13px; margin-bottom: 5px;">📝 मासिक यूनिट टेस्ट मूल्यांकन (Monthly Unit Test Record - Max: 100):</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 11.5px; border: 1px solid #CBD5E1; text-align: center;">
                <thead>
                    <tr style="background: #1E3A8A; color: white;">
                        <th style="padding: 6px;">हिन्दी (20)</th>
                        <th style="padding: 6px;">अंग्रेजी (20)</th>
                        <th style="padding: 6px;">गणित (20)</th>
                        <th style="padding: 6px;">भौतिक (20)</th>
                        <th style="padding: 6px;">रसायन (20)</th>
                        <th style="padding: 6px; background: #0F172A;">कुल प्राप्तांक (100)</th>
                        <th style="padding: 6px; background: #059669;">प्रतिशत (%)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr style="background: #F8FAFC; font-weight: bold; color: #1E293B;">
                        <td style="padding: 6px; border: 1px solid #CBD5E1;">{m_hin if m_hin else '-'}</td>
                        <td style="padding: 6px; border: 1px solid #CBD5E1;">{m_eng if m_eng else '-'}</td>
                        <td style="padding: 6px; border: 1px solid #CBD5E1;">{m_mat if m_mat else '-'}</td>
                        <td style="padding: 6px; border: 1px solid #CBD5E1;">{m_phy if m_phy else '-'}</td>
                        <td style="padding: 6px; border: 1px solid #CBD5E1;">{m_che if m_che else '-'}</td>
                        <td style="padding: 6px; border: 1px solid #CBD5E1; color: #1E3A8A; font-size: 12.5px;">{test_display_tot}</td>
                        <td style="padding: 6px; border: 1px solid #CBD5E1; color: #059669; font-size: 12.5px;">{test_display_pct}</td>
                    </tr>
                </tbody>
            </table>
        </div>

        <div style="margin-bottom: 12px;">
            <div style="color: #1E3A8A; font-weight: bold; font-size: 13px; margin-bottom: 5px;">📋 प्रमुख सह-पाठ्यचर्या गतिविधियां एवं प्रतियोगिताएं:</div>
            <table style="width: 100%; border-collapse: collapse; font-size: 11.5px; border: 1px solid #CBD5E1;">
                <thead>
                    <tr style="background: #334155; color: white; text-align: left;">
                        <th style="padding: 6px; width: 12%; text-align: center;">तिथि</th>
                        <th style="padding: 6px; width: 38%;">गतिविधि / प्रतियोगिता</th>
                        <th style="padding: 6px; width: 18%; text-align: center;">श्रेणी</th>
                        <th style="padding: 6px; width: 22%;">सीख / प्रस्तुति</th>
                        <th style="padding: 6px; width: 10%; text-align: center;">अंक</th>
                    </tr>
                </thead>
                <tbody>
                    {activities_rows}
                </tbody>
            </table>
        </div>

        <div style="border: 1px solid #CBD5E1; border-radius: 6px; padding: 10px; background: #F8FAFC; margin-top: 15px;">
            <div style="margin: 0 0 6px 0; color: #1E3A8A; font-weight: bold; font-size: 12.5px;">📝 आंतरिक मूल्यांकन रूब्रिक्स (UP Board Marking Criteria - पूर्णांक: 20)</div>
            <div style="display: flex; gap: 8px; font-size: 11.5px; text-align: center;">
                <div style="flex: 1; background: white; padding: 5px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>1. नियमितता</strong><br>(5 M)</div>
                <div style="flex: 1; background: white; padding: 5px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>2. मौलिकता</strong><br>(5 M)</div>
                <div style="flex: 1; background: white; padding: 5px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>3. रचनात्मकता</strong><br>(5 M)</div>
                <div style="flex: 1; background: white; padding: 5px; border: 1px solid #CBD5E1; border-radius: 4px;"><strong>4. आचरण</strong><br>(5 M)</div>
            </div>

            <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-top: 25px; padding-top: 8px; border-top: 1px dashed #94A3B8; font-size: 11.5px;">
                <div><strong>विद्यार्थी के हस्ताक्षर:</strong> _____________________<br><span style="color:#64748B;">दिनांक: {today_str}</span></div>
                <div style="text-align: right;"><strong>कक्षा अध्यापक / प्रभारी हस्ताक्षर:</strong> _____________________<br><span style="color:#64748B;">कक्षा अध्यापक</span></div>
            </div>
        </div>
    </div>
</body>
</html>"""

# ==========================================
# 4. SIDEBAR NAVIGATION
# ==========================================
st.sidebar.title("🧭 Portal Navigation")
selected_portal = st.sidebar.radio("Select Access:", ["🎓 Student Individual Portal", "⚙️ Teacher / Admin Control Center"])
st.sidebar.divider()

# ==========================================
# 5. ADMIN CONTROL CENTER (Teacher Only)
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
        "📥 Google Form Sync (All-in-One Responses)",
        "📝 Question Bank (Excel/Manual)",
        "📊 Overall Merit List & Results (PDF Download)", 
        "📷 Student Photo Upload (Roll No Wise)",
        "📋 14 Official Activities Calendar",
        "💾 Full Database Backup & Restore"
    ])

    st.divider()
    conn = get_db()

    # SECTION 1: QUIZZES (Class-Wise Single Live Quiz Rule)
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

    # SECTION 2: MASTER STUDENTS, ATTENDANCE & TEST MARKS
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

    # SECTION 3: GOOGLE FORM SYNC
    elif admin_tab == "📥 Google Form Sync (All-in-One Responses)":
        st.subheader("📥 Google Form Responses File (.xlsx / .csv)")
        uploaded_form = st.file_uploader("Upload Responses File:", type=["xlsx", "csv"])
        if uploaded_form and st.button("Sync Responses & Goals", type="primary"):
            try:
                df_form = pd.read_csv(uploaded_form, dtype=str) if uploaded_form.name.endswith('.csv') else pd.read_excel(uploaded_form, dtype=str)
                cols = list(df_form.columns)
                roll_col = next((col for col in cols if "roll" in col.lower() or "अनुक्रमांक" in col), None)
                st_col = next((col for col in cols if "अल्पकालिक" in col or "short-term" in col.lower()), None)
                lt_col = next((col for col in cols if "दीर्घकालिक" in col or "long-term" in col.lower()), None)

                if not roll_col:
                    st.error("Roll Number column nahi mila!")
                else:
                    goals_cnt = 0
                    acts_cnt = 0
                    for _, r in df_form.iterrows():
                        r_no = clean_val(r.get(roll_col, ""))
                        if not r_no:
                            continue
                        st_val = clean_val(r.get(st_col, "")) if st_col else ""
                        lt_val = clean_val(r.get(lt_col, "")) if lt_col else ""

                        if st_val or lt_val:
                            conn.execute("""
                                UPDATE master_students
                                SET short_term_goal = CASE WHEN ? != '' THEN ? ELSE short_term_goal END,
                                    long_term_goal  = CASE WHEN ? != '' THEN ? ELSE long_term_goal END
                                WHERE roll_no = ?
                            """, (st_val, st_val, lt_val, lt_val, r_no))
                            goals_cnt += 1

                        for act in DEFAULT_ACTIVITIES:
                            act_num = str(act["sno"])
                            desc_c = next((cn for cn in cols if f"[{act_num}." in cn and ("description" in cn.lower() or "कार्य किया" in cn)), None)
                            refl_c = next((cn for cn in cols if f"[{act_num}." in cn and ("reflection" in cn.lower() or "सीखा" in cn)), None)
                            link_c = next((cn for cn in cols if f"[{act_num}." in cn and ("link" in cn.lower() or "photo" in cn.lower() or "drive" in cn.lower())), None)

                            d_val = clean_val(r.get(desc_c, "")) if desc_c else ""
                            rf_val = clean_val(r.get(refl_c, "")) if refl_c else ""
                            lk_val = clean_val(r.get(link_c, "")) if link_c else ""

                            if d_val or rf_val or lk_val:
                                d_img = convert_gdrive_link(lk_val)
                                today_now = datetime.now().strftime("%d-%m-%Y")
                                conn.execute("""
                                    INSERT INTO portfolio_entries (
                                        roll_no, activity_name, category, activity_date,
                                        student_description, student_reflection, evidence_link,
                                        marks_awarded, submitted_on
                                    ) VALUES (?, ?, ?, ?, ?, ?, ?, 5, ?)
                                """, (r_no, act["name"], act["cat"], act["date"], d_val, rf_val, d_img, today_now))
                                acts_cnt += 1

                    conn.commit()
                    st.success(f"Successfully synced: {goals_cnt} goals and {acts_cnt} activity records!")
                    st.rerun()
            except Exception as e:
                st.error(f"Sync failed: {e}")

    # SECTION 4: QUESTION BANK
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
                st.success(f"{cnt} questions imported successfully!")
                st.rerun()

    # SECTION 5: MERIT LIST (PDF)
    elif admin_tab == "📊 Overall Merit List & Results (PDF Download)":
        st.subheader("Official Merit List (Topper to Lower)")
        if not quizzes_df.empty:
            quiz_options = {f"[{r['target_class']}] {r['quiz_title']}": r['id'] for _, r in quizzes_df.iterrows()}
            sel_q_label = st.selectbox("Select Quiz:", list(quiz_options.keys()))
            sel_q_id = quiz_options[sel_q_label]
            
            q_info = dict(conn.execute("SELECT * FROM quizzes WHERE id = ?", (sel_q_id,)).fetchone())
            
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

    # SECTION 6: PHOTO MANAGEMENT
    elif admin_tab == "📷 Student Photo Upload (Roll No Wise)":
        st.subheader("Student Photo Management (Bulk / Single)")
        st.info("💡 Photo ka filename Roll Number rakhein (e.g. `1.jpg`, `15.png`).")
        
        b_files = st.file_uploader("Select Multiple Photos:", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
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

    # SECTION 7: CALENDAR
    elif admin_tab == "📋 14 Official Activities Calendar":
        st.subheader("📋 कक्षा 12-B आधिकारिक गतिविधि एवं प्रतियोगिता कैलेंडर (UP Board 2026-27)")
        df_acts = pd.DataFrame(DEFAULT_ACTIVITIES)
        df_acts.columns = ["क्र. सं.", "तिथि", "प्रतियोगिता / गतिविधि का नाम", "श्रेणी / प्रकार", "विषय / विवरण", "प्रभारी / मूल्यांकनकर्ता"]
        st.dataframe(df_acts, use_container_width=True)

    # SECTION 8: BACKUP
    elif admin_tab == "💾 Full Database Backup & Restore":
        st.subheader("💾 Full Database Backup")
        stu_exp = pd.read_sql_query("SELECT * FROM master_students", conn)
        q_exp = pd.read_sql_query("SELECT * FROM quizzes", conn)
        ques_exp = pd.read_sql_query("SELECT * FROM questions", conn)
        subs_exp = pd.read_sql_query("SELECT * FROM submissions", conn)
        resp_exp = pd.read_sql_query("SELECT * FROM student_responses", conn)
        port_exp = pd.read_sql_query("SELECT * FROM portfolio_entries", conn)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            stu_exp.to_excel(writer, sheet_name='Master_Students', index=False)
            q_exp.to_excel(writer, sheet_name='Quizzes', index=False)
            ques_exp.to_excel(writer, sheet_name='Questions', index=False)
            subs_exp.to_excel(writer, sheet_name='Submissions', index=False)
            resp_exp.to_excel(writer, sheet_name='Responses', index=False)
            port_exp.to_excel(writer, sheet_name='Portfolio_Entries', index=False)
            
        st.download_button("📥 Download Full Backup (.xlsx)", output.getvalue(), f"Full_Database_Backup_{get_ist_now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    conn.close()

# ==========================================
# 6. STUDENT INDIVIDUAL PORTAL (Strictly Self)
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

    # Student Login Form
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
    s_roll = s_dict.get('roll_no', '')

    # Fetch individual portfolio activities
    entries_df = pd.read_sql_query("SELECT * FROM portfolio_entries WHERE roll_no = ? ORDER BY id ASC", conn, params=(s_roll,)) if s_roll else pd.DataFrame()

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
    # TAB 2: MY EXAM SCORECARD & ANSWER KEY (Self Only)
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
    # TAB 3: MY ATTENDANCE & UNIT TESTS (Self Only)
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
    # TAB 4: MY UP BOARD PORTFOLIO CARD (View & Print Self)
    # -------------------------------------------------------------
    with stu_tabs[3]:
        st.subheader("🎴 2-Page UP Board Student Portfolio Card")
        
        portfolio_html = generate_upboard_card(s_dict, entries_df)
        
        st.download_button(
            label="📥 Download My Official Portfolio Card (.html)",
            data=portfolio_html,
            file_name=f"Portfolio_{s_dict.get('student_name')}.html",
            mime="text/html",
            type="primary"
        )
        st.caption("💡 Downloaded HTML file ko kisi bhi browser me open karke 'Ctrl + P' dabayein aur Save as PDF karein.")
        st.components.v1.html(portfolio_html, height=1150, scrolling=True)

    conn.close()
