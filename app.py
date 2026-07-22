import streamlit as st
import google.generativeai as genai
import datetime
import io
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from supabase import create_client, Client
from PIL import Image

st.set_page_config(page_title="Nyaya Assist AI", page_icon="⚖️", layout="wide")

TEMPLATES_FOLDER = "master_templates"
if not os.path.exists(TEMPLATES_FOLDER):
    os.makedirs(TEMPLATES_FOLDER)

# 🔑 Fetch Secrets (Hidden from UI)
default_api_key = st.secrets.get("GEMINI_API_KEY", "")
supabase_url = st.secrets.get("SUPABASE_URL", "")
supabase_key = st.secrets.get("SUPABASE_KEY", "")

# 🗄️ Supabase Initialization
supabase: Client = None
if supabase_url and supabase_key:
    try:
        supabase = create_client(supabase_url, supabase_key)
    except Exception as e:
        st.error(f"Supabase connection error: {e}")

# 🔐 PREMIUM MODERN CSS STYLING
st.markdown("""
<style>
    .stApp {
        background-color: #050508 !important;
        color: #f3f4f6 !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #08090d !important;
        border-right: 1px solid #161821 !important;
        padding-top: 10px;
    }
    
    .stTextInput > div > div > input, .stSelectbox > div > div, .stTextArea > div > div > textarea {
        background-color: #0e1017 !important;
        color: #ffffff !important;
        border: 1px solid #1f222e !important;
        border-radius: 12px !important;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #7c3aed, #6366f1) !important;
        color: white !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 8px 16px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px rgba(124, 58, 237, 0.3) !important;
        transition: all 0.2s ease-in-out;
    }
    
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(124, 58, 237, 0.5) !important;
    }
    
    .stDownloadButton > button {
        background: #1e1b4b !important;
        color: #c084fc !important;
        border: 1px solid #a855f7 !important;
        border-radius: 10px !important;
    }

    div.stRadio > div {
        gap: 6px !important;
    }
    div.stRadio > div[role="radiogroup"] > label {
        background-color: #0e1017 !important;
        border: 1px solid #1a1d2d !important;
        border-radius: 10px !important;
        padding: 10px 14px !important;
        color: #9ca3af !important;
        font-weight: 500 !important;
        font-size: 14px !important;
        transition: all 0.2s ease;
        box-shadow: none !important;
    }
    div.stRadio > div[role="radiogroup"] > label:hover {
        background-color: #161925 !important;
        border-color: #7c3aed !important;
        color: #ffffff !important;
    }
    div.stRadio > div[role="radiogroup"] > label[data-baseweb="radio"] input:checked + div {
        background: linear-gradient(135deg, #7c3aed, #6366f1) !important;
        color: #ffffff !important;
    }

    div[data-testid="column"]:nth-child(3) .stButton > button {
        border-radius: 50% !important;
        width: 42px !important;
        height: 42px !important;
        padding: 0 !important;
        font-weight: 700 !important;
        font-size: 14px !important;
        line-height: 42px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        margin-left: auto !important;
    }

    .pro-banner {
        background: linear-gradient(135deg, #1e1b4b, #0f172a);
        border: 1px solid #3b0764;
        border-radius: 14px;
        padding: 12px 14px;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 20px;
        box-shadow: 0 4px 14px rgba(0,0,0,0.3);
    }

    .welcome-hero {
        text-align: center;
        padding: 20px 20px 10px 20px;
    }
    .welcome-title {
        font-size: 32px;
        font-weight: 700;
        color: #ffffff;
        margin-bottom: 6px;
    }
    .welcome-subtitle {
        color: #9ca3af;
        font-size: 15px;
        margin-bottom: 25px;
    }
    
    .case-studio-banner {
        background-color: #0e1017;
        border: 1px solid #1f222e;
        border-radius: 16px;
        padding: 18px 24px;
        margin: 0 auto 25px auto;
        max-width: 800px;
    }

    .action-card {
        background-color: #0e1017;
        border: 1px solid #1f222e;
        border-radius: 12px;
        padding: 12px;
        margin-bottom: 8px;
    }

    .attachment-chip {
        background-color: #1e1b4b;
        border: 1px solid #7c3aed;
        border-radius: 10px;
        padding: 8px 12px;
        color: #c084fc;
        font-size: 13px;
        font-weight: 600;
        margin-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 🔐 USER SYSTEM STATE
if "user" not in st.session_state:
    st.session_state.user = None
if "auth_mode" not in st.session_state:
    st.session_state.auth_mode = "login"
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0
if "current_session" not in st.session_state:
    st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
if "nav_menu" not in st.session_state:
    st.session_state.nav_menu = "💬 Case Studio"

def login_user(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        st.success("✅ Login successful!")
        st.rerun()
    except Exception as e:
        st.error(f"Login failed: {e}")

def signup_user(email, password, full_name, phone):
    try:
        res = supabase.auth.sign_up({
            "email": email, 
            "password": password,
            "options": {
                "data": {
                    "full_name": full_name,
                    "phone": phone
                }
            }
        })
        st.session_state.user = res.user
        st.success("✅ Account created successfully!")
        st.rerun()
    except Exception as e:
        st.error(f"Signup failed: {e}")

# ----------------- LOGIN / SIGNUP SCREEN -----------------
if not st.session_state.user:
    st.markdown("<br>", unsafe_allow_html=True)
    
    col_auth1, col_auth2, col_auth3 = st.columns([1, 2, 1])
    with col_auth2:
        if st.session_state.auth_mode == "login":
            st.markdown("""
            <div style="text-align: center;">
                <div style="font-size: 55px;">⚖️</div>
                <h2 style="margin: 0; font-weight: 700; color: #ffffff;">Nyaya Assist <span style="color: #a855f7;">AI</span></h2>
                <p style="color: #9ca3af; font-size: 14px; margin-top: 4px;">Your AI Legal Assistant</p>
                <br>
                <h4 style="margin: 0; font-weight: 600; color: #e5e7eb;">Login to your account</h4>
            </div>
            """, unsafe_allow_html=True)
            
            email = st.text_input("Email Address", placeholder="name@example.com", key="login_email")
            password = st.text_input("Password", type="password", placeholder="••••••••", key="login_pass")
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Login  ➔", key="btn_login", use_container_width=True):
                if email and password and supabase:
                    login_user(email, password)
                else:
                    st.warning("Please enter Email & Password.")
                    
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Don't have an account? Sign up", key="switch_to_signup", use_container_width=True):
                st.session_state.auth_mode = "signup"
                st.rerun()

        else:
            st.markdown("""
            <div style="text-align: left;">
                <h2 style="margin: 0; color: #ffffff;">Create your account</h2>
            </div>
            """, unsafe_allow_html=True)
            
            full_name = st.text_input("Full Name", placeholder="Adv. Rajesh Sharma", key="signup_name")
            email = st.text_input("Email", placeholder="name@example.com", key="signup_email")
            phone = st.text_input("Phone Number (Optional)", placeholder="+91 98765 43210", key="signup_phone")
            password = st.text_input("Password", type="password", placeholder="••••••••", key="signup_pass")
            confirm_pass = st.text_input("Confirm Password", type="password", placeholder="••••••••", key="signup_cpass_unique")
            
            agree = st.checkbox("I agree to the Terms of Service and Privacy Policy")
            
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Sign Up  ➔", key="btn_signup", use_container_width=True):
                if not agree:
                    st.warning("Please agree to the Terms & Privacy Policy.")
                elif password != confirm_pass:
                    st.error("Passwords do not match!")
                elif email and password and supabase:
                    signup_user(email, password, full_name, phone)
                else:
                    st.warning("Please fill all required fields.")
                    
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Already have an account? Login", key="switch_to_login", use_container_width=True):
                st.session_state.auth_mode = "login"
                st.rerun()

    st.stop()

# ----------------- LOGGED IN APPLICATION DASHBOARD -----------------
current_user_email = st.session_state.user.email
display_user_name = current_user_email.split('@')[0].capitalize()

try:
    meta_name = st.session_state.user.user_metadata.get("full_name")
    if meta_name:
        display_user_name = meta_name.split()[0].capitalize()
except Exception:
    pass

user_initials = "".join([part[0].upper() for part in display_user_name.split()[:2]]) if display_user_name else "VT"

def save_chat_to_supabase(session_name, role, content):
    if supabase:
        try:
            supabase.table("chats").insert({
                "user_email": current_user_email,
                "session_name": session_name,
                "role": role,
                "content": content
            }).execute()
        except Exception as e:
            st.error(f"Error saving chat: {e}")

def get_supabase_sessions():
    if supabase:
        try:
            res = supabase.table("chats").select("session_name").eq("user_email", current_user_email).execute()
            sessions = list(set([row["session_name"] for row in res.data]))
            return sorted(sessions, reverse=True)
        except Exception:
            return []
    return []

def get_supabase_chat_history(session_name):
    if supabase:
        try:
            res = supabase.table("chats").select("role, content").eq("user_email", current_user_email).eq("session_name", session_name).order("id").execute()
            return res.data
        except Exception:
            return []
    return []

def delete_supabase_session(session_name):
    if supabase:
        try:
            supabase.table("chats").delete().eq("user_email", current_user_email).eq("session_name", session_name).execute()
            st.success(f"🗑️ Deleted chat: {session_name}")
            st.rerun()
        except Exception as e:
            st.error(f"Error deleting chat: {e}")

def load_all_local_templates():
    template_data = ""
    if os.path.exists(TEMPLATES_FOLDER):
        files = os.listdir(TEMPLATES_FOLDER)
        for file in sorted(files):
            file_path = os.path.join(TEMPLATES_FOLDER, file)
            if file.endswith(".txt"):
                try:
                    with open(file_path, "r", encoding="utf-8") as f:
                        template_data += f"\n\n--- MASTER TEMPLATE DATABASE: {file} ---\n" + f.read()
                except Exception:
                    pass
            elif file.endswith(".docx"):
                try:
                    doc_read = Document(file_path)
                    doc_text = '\n'.join([p.text for p in doc_read.paragraphs])
                    template_data += f"\n\n--- MASTER TEMPLATE DATABASE: {file} ---\n" + doc_text
                except Exception:
                    pass
    return template_data

def read_uploaded_file_content(uploaded_file):
    if uploaded_file is None:
        return "", None
    file_type = uploaded_file.name.split('.')[-1].lower()
    content = ""
    image_obj = None
    
    if file_type == "txt":
        content = uploaded_file.read().decode("utf-8")
    elif file_type == "docx":
        doc = Document(uploaded_file)
        content = '\n'.join([p.text for p in doc.paragraphs])
    elif file_type == "pdf":
        try:
            import pypdf
            pdf_reader = pypdf.PdfReader(uploaded_file)
            content = "\n".join([page.extract_text() for page in pdf_reader.pages if page.extract_text()])
        except Exception:
            content = f"[PDF FILE ATTACHED: {uploaded_file.name}]"
    elif file_type in ["jpg", "jpeg", "png"]:
        image_obj = Image.open(uploaded_file)
        content = f"[IMAGE FILE ATTACHED: {uploaded_file.name}]"
        
    return content, image_obj

def create_court_ready_docx(text):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(14.0)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.5)
    section.right_margin = Inches(1.0)
    
    draft_content = ""
    upper_text = text.upper()
    if "START_DRAFT" in upper_text and "END_DRAFT" in upper_text:
        start_idx = upper_text.find("START_DRAFT") + len("START_DRAFT")
        end_idx = upper_text.find("END_DRAFT")
        draft_content = text[start_idx:end_idx].strip()
    else:
        draft_content = text.strip()

    lines = draft_content.split('\n')
    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(6)
        
        if stripped.startswith("##") or stripped.startswith("**"):
            clean_line = stripped.replace("#", "").replace("**", "").strip()
            run = p.add_run(clean_line)
            run.bold = True
            run.font.size = Pt(14)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        else:
            run = p.add_run(stripped)
            run.font.size = Pt(12)
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            
    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ----------------- 🎯 PREMIUM SIDEBAR NAVIGATION -----------------
nav_options = ["💬 Case Studio", "📖 Library", "📁 My Cases", "⚙️ Settings", "📂 Chat History"]

with st.sidebar:
    st.markdown("""
    <div style="text-align: center; margin-bottom: 22px; padding-top: 10px;">
        <div style="font-size: 38px;">⚖️</div>
        <h3 style="margin:4px 0 0 0; font-size: 20px; font-weight:700; letter-spacing: -0.5px;">Nyaya Assist <span style="color:#a855f7;">AI</span></h3>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="pro-banner">
        <div>
            <div style="font-weight: 700; font-size: 13px; color: #ffffff;">👑 Upgrade to Pro</div>
            <div style="font-size: 11px; color: #9ca3af;">Unlock advanced features</div>
        </div>
        <span style="font-size: 16px; color: #c084fc;">➔</span>
    </div>
    """, unsafe_allow_html=True)

    if st.button("➕ New Case", key="btn_side_newcase", use_container_width=True):
        st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
        st.session_state.nav_menu = "💬 Case Studio"
        st.rerun()

    st.markdown("<div style='margin: 12px 0;'></div>", unsafe_allow_html=True)
    curr_idx = nav_options.index(st.session_state.nav_menu) if st.session_state.nav_menu in nav_options else 0
    menu = st.radio("Menu", nav_options, index=curr_idx, label_visibility="collapsed")
    st.session_state.nav_menu = menu

    st.markdown("<hr style='border: 0; border-top: 1px solid #1f222e; margin: 20px 0;'>", unsafe_allow_html=True)
    st.markdown("<div style='font-size: 12px; font-weight:700; color:#8e92a4; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom:12px;'>Chat History</div>", unsafe_allow_html=True)

    sidebar_sessions = get_supabase_sessions()
    if sidebar_sessions:
        for sess in sidebar_sessions[:5]:
            if st.button(f"💬 {sess[:22]}...", key=f"sbar_sess_{sess}", use_container_width=True):
                st.session_state.current_session = sess
                st.session_state.nav_menu = "💬 Case Studio"
                st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("💬 View All Chats", use_container_width=True):
        st.session_state.nav_menu = "📂 Chat History"
        st.rerun()

    st.markdown("<br>"*2, unsafe_allow_html=True)
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.user = None
        st.session_state.auth_mode = "login"
        st.rerun()

# ----------------- 👤 PROFILE POPUP MODAL FUNCTION -----------------
@st.dialog("👤 Manage Account & Subscription")
def show_profile_dialog():
    st.markdown(f"### Advocate: **{display_user_name}**")
    st.markdown(f"📧 **Email:** `{current_user_email}`")
    st.markdown("👑 **Subscription Status:** `Free Tier`")
    
    st.markdown("---")
    st.markdown("#### 🔒 Change Password")
    new_pass = st.text_input("New Password", type="password", placeholder="••••••••", key="modal_new_pass")
    if st.button("Update Password", key="modal_update_pass_btn"):
        if new_pass and supabase:
            try:
                supabase.auth.update_user({"password": new_pass})
                st.success("✅ Password updated successfully!")
            except Exception as e:
                st.error(f"Error: {e}")
        else:
            st.warning("Please enter a valid password.")
            
    st.markdown("---")
    if st.button("🚀 Upgrade to Pro Plan", key="modal_upgrade_btn", use_container_width=True):
        st.info("Pro Plan subscription gateway will be integrated soon!")

# ----------------- TOP APP BAR -----------------
col_top_h1, col_top_h2, col_top_h3 = st.columns([5, 1, 1])
with col_top_h1:
    st.markdown("### Nyaya Assist <span style='color:#a855f7;'>AI</span>", unsafe_allow_html=True)
with col_top_h2:
    st.markdown("<div style='text-align: right; padding-top: 10px; color: #9ca3af; font-size: 16px;'>🔔</div>", unsafe_allow_html=True)
with col_top_h3:
    if st.button(user_initials, key="btn_open_profile_modal"):
        show_profile_dialog()

if default_api_key:
    os.environ["GEMINI_API_KEY"] = default_api_key
    genai.configure(api_key=default_api_key)
    
    if "app_lang" not in st.session_state:
        st.session_state.app_lang = "Pure Hindi (Formal Legal)"
    if "advocate_name" not in st.session_state:
        st.session_state.advocate_name = ""
    if "default_court" not in st.session_state:
        st.session_state.default_court = ""

    system_instruction = f"""
    ROLE & OBJECTIVE:
    You are a legendary Senior Advocate & Legal Strategist of the Supreme Court of India. You possess expert legal knowledge of Indian law (BNS, BNSS, BSA, CPC, CrPC, Constitution, etc.) and Supreme Court/High Court precedents.
    Response Language Style: {st.session_state.app_lang}.
    
    MANDATORY DRAFTING DETAILS:
    - Advocate Name: {st.session_state.advocate_name if st.session_state.advocate_name else "[Advocate Name]"}
    - Court Heading: {st.session_state.default_court if st.session_state.default_court else "[Court Name]"}
    
    DEEP ANALYTICAL RESPONSE STRUCTURE:
    When user uploads case files, photos of documents, or asks complex legal questions:
    1. 🔍 **DEEP CASE ANALYSIS & LOOPHOLES:** Highlight strong facts, weaknesses, and loopholes in the opponent's case.
    2. 🎯 **STRATEGIC CROSS-EXAMINATION QUESTIONS:** Provide 5 to 7 sharp, high-impact questions to ask the opponent or witnesses in court.
    3. ⚖️ **APPLICABLE SECTIONS & PRECEDENTS:** Mention key legal sections and strategic precedents.
    4. 📜 **FORMAL COURT DRAFT:** ALWAYS generate a full court draft based on standard Indian court format.
    
    OUTPUT FORMATTING:
    Final legal draft MUST ALWAYS be strictly enclosed inside:
      START_DRAFT
      ... (Full legal court draft content) ...
      END_DRAFT
    """

    model = genai.GenerativeModel(
        model_name="gemini-3.6-flash",
        system_instruction=system_instruction,
        generation_config={"temperature": 0.0}
    )

    if menu == "📖 Library":
        st.markdown("## 📖 Legal Library & Judgment Search")
        st.markdown("यहाँ आप और आपके पापा अपनी जरूरी लीगल बुक्स, जजमेंट्स या केस फाइल्स अपलोड और सर्च कर सकते हैं।")
        
        search_query = st.text_input("🔍 Search Library (Enter book name, section, or keyword)...", placeholder="e.g. Section 498A, Bail Judgment, etc.")
        
        st.markdown("### 📤 Upload Files / Judgments to Library")
        uploaded_doc = st.file_uploader(
            "Select Judgment or Law Reference File (.pdf, .docx, .txt)", 
            type=["pdf", "docx", "txt"], 
            key="library_search_uploader"
        )
        
        if uploaded_doc:
            save_path = os.path.join(TEMPLATES_FOLDER, uploaded_doc.name)
            with open(save_path, "wb") as f:
                f.write(uploaded_doc.getbuffer())
            st.success(f"✅ Successfully saved `{uploaded_doc.name}` to library storage!")

        st.markdown("---")
        st.markdown("### 📁 Saved Documents & Judgments")
        if os.path.exists(TEMPLATES_FOLDER):
            files = os.listdir(TEMPLATES_FOLDER)
            if files:
                filtered_files = [f for f in files if search_query.lower() in f.lower()] if search_query else files
                for f in sorted(filtered_files):
                    col_l1, col_l2 = st.columns([4, 1])
                    with col_l1:
                        st.markdown(f"📄 **`{f}`**")
                    with col_l2:
                        if st.button("🗑️ Delete", key=f"del_lib_{f}"):
                            os.remove(os.path.join(TEMPLATES_FOLDER, f))
                            st.success(f"Deleted {f}")
                            st.rerun()
            else:
                st.info("No documents stored in library yet.")

    elif menu == "📁 My Cases":
        st.markdown("## 📁 My Cases & Documents")
        st.markdown("All your active court drafts and client case files in one place.")
        sessions = get_supabase_sessions()
        if sessions:
            for s in sessions:
                st.markdown(f"- 📁 **Case Session:** `{s}`")
        else:
            st.info("No active cases found.")

    elif menu == "⚙️ Settings":
        st.markdown("## ⚙️ App & Profile Settings")
        st.markdown("Configure your legal credentials for auto-filled drafting.")
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            adv_name_input = st.text_input("Advocate Full Name:", value=st.session_state.advocate_name, placeholder="e.g. Adv. Vedant Sharma")
        with col_s2:
            court_input = st.text_input("Default Court Jurisdiction:", value=st.session_state.default_court, placeholder="e.g. High Court of Delhi")
            
        selected_lang = st.selectbox("Preferred AI Language Mode", [
            "Pure Hindi (Formal Legal)",
            "Hinglish / Bilingual (Hindi & English)",
            "Pure English (Professional Legal)"
        ], index=0)
        
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💾 Save Settings"):
            st.session_state.advocate_name = adv_name_input
            st.session_state.default_court = court_input
            st.session_state.app_lang = selected_lang
            st.success("✅ Settings updated successfully!")

    elif menu == "📂 Chat History":
        st.markdown("## 📂 Saved Chat Archives")
        st.markdown("यहाँ आपकी पुरानी सभी चैट्स सुरक्षित हैं। आप किसी भी चैट को फिर से खोल सकते हैं या डिलीट कर सकते हैं।")
        
        sessions = get_supabase_sessions()
        if sessions:
            for s in sessions:
                with st.expander(f"💬 Chat Session: {s}"):
                    col_h1, col_h2 = st.columns([4, 1])
                    with col_h1:
                        if st.button(f"💬 Open & Resume Chat ({s})", key=f"open_{s}"):
                            st.session_state.current_session = s
                            st.session_state.nav_menu = "💬 Case Studio"
                            st.rerun()
                    with col_h2:
                        if st.button("🗑️ Delete Chat", key=f"del_{s}"):
                            delete_supabase_session(s)

                    st.markdown("---")
                    chats = get_supabase_chat_history(s)
                    for c in chats:
                        role_label = "👤 You" if c['role'] == "user" else "⚖️ NyayaAssist AI"
                        st.markdown(f"**{role_label}:**")
                        st.markdown(c['content'])
                        st.markdown("---")
        else:
            st.info("No saved chat history found in your account.")

    elif menu == "💬 Case Studio":
        sessions = get_supabase_sessions()
        messages = get_supabase_chat_history(st.session_state.current_session)

        # 🌟 HERO HOME SECTION WITH NEW CASE BUTTON
        if not messages:
            st.markdown(f"""
            <div class="welcome-hero">
                <div style="margin-bottom: 15px;">
                    <svg width="90" height="90" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
                        <path d="M50 15V80M50 15L25 35M50 15L75 35" stroke="url(#paint0_linear)" stroke-width="3" stroke-linecap="round"/>
                        <path d="M15 50C15 50 25 65 35 65C45 65 35 50 35 50" stroke="url(#paint1_linear)" stroke-width="3"/>
                        <path d="M65 50C65 50 75 65 85 65C95 65 85 50 85 50" stroke="url(#paint2_linear)" stroke-width="3"/>
                        <path d="M38 80H62" stroke="#8b5cf6" stroke-width="3" stroke-linecap="round"/>
                        <defs>
                            <linearGradient id="paint0_linear" x1="50" y1="15" x2="50" y2="80" gradientUnits="userSpaceOnUse">
                                <stop stop-color="#c084fc"/>
                                <stop offset="1" stop-color="#6366f1"/>
                            </linearGradient>
                            <linearGradient id="paint1_linear" x1="15" y1="50" x2="35" y2="65" gradientUnits="userSpaceOnUse">
                                <stop stop-color="#a855f7"/>
                                <stop offset="1" stop-color="#6366f1"/>
                            </linearGradient>
                            <linearGradient id="paint2_linear" x1="65" y1="50" x2="85" y2="65" gradientUnits="userSpaceOnUse">
                                <stop stop-color="#a855f7"/>
                                <stop offset="1" stop-color="#6366f1"/>
                            </linearGradient>
                        </defs>
                    </svg>
                </div>
                <div class="welcome-title">Welcome back, <span style="color: #a855f7;">{display_user_name}</span></div>
                <div class="welcome-subtitle">Your AI Legal Assistant is ready to support your legal work.</div>
            </div>
            """, unsafe_allow_html=True)

            col_b1, col_b2, col_b3 = st.columns([1, 4, 1])
            with col_b2:
                col_banner_txt, col_banner_btn = st.columns([3, 1])
                with col_banner_txt:
                    st.markdown("""
                    <div class="case-studio-banner">
                        <div style="display: flex; align-items: center; gap: 15px;">
                            <span style="font-size: 28px;">📁</span>
                            <div>
                                <div style="font-weight: 700; font-size: 16px; color: #ffffff;">Case Studio</div>
                                <div style="font-size: 13px; color: #9ca3af;">Describe your case, upload documents, and get AI guidance.</div>
                            </div>
                        </div>
                    </div>
                    """, unsafe_allow_html=True)
                with col_banner_btn:
                    if st.button("➕ New Case", key="btn_new_case_home", use_container_width=True):
                        st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
                        st.rerun()

        # 📜 CHAT MESSAGES DISPLAY AREA
        else:
            col_hdr1, col_hdr2 = st.columns([4, 1])
            with col_hdr1:
                st.markdown(f"#### 💬 Case: `{st.session_state.current_session}`")
            with col_hdr2:
                if st.button("➕ New Case", key="btn_new_case_chat", use_container_width=True):
                    st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
                    st.rerun()

            for idx, message in enumerate(messages):
                with st.chat_message(message["role"]):
                    st.markdown(message["content"])
                    
                    upper_content = message["content"].upper()
                    if message["role"] == "assistant" and "START_DRAFT" in upper_content:
                        docx_data = create_court_ready_docx(message["content"])
                        st.download_button(
                            label="📥 Download Court-Ready Word (.docx)",
                            data=docx_data,
                            file_name=f"Court_Draft_{idx}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"docx_{idx}"
                        )

        # ----------------- 🎯 CLEAN & ACCESSIBLE INPUT DECK -----------------
        st.markdown("---")
        col_deck1, col_deck2, col_deck3 = st.columns(3)

        with col_deck1:
            st.markdown("<div class='action-card'><strong style='color:#c084fc;'>📤 2. Upload Template</strong></div>", unsafe_allow_html=True)
            format_file = st.file_uploader(
                "Notice / Petition format", 
                type=["docx", "txt", "pdf"], 
                key=f"format_{st.session_state.uploader_key}"
            )

        with col_deck2:
            st.markdown("<div class='action-card'><strong style='color:#c084fc;'>📸 3. Upload Case Documents</strong></div>", unsafe_allow_html=True)
            case_file = st.file_uploader(
                "Case PDF, Word or Photos", 
                type=["docx", "txt", "pdf", "jpg", "jpeg", "png"], 
                key=f"case_{st.session_state.uploader_key}"
            )

        with col_deck3:
            st.markdown("<div class='action-card'><strong style='color:#c084fc;'>🎙️ Voice Assistant</strong></div>", unsafe_allow_html=True)
            voice_html = """
            <div style="background: #0e1017; border: 1px solid #1f222e; border-radius: 8px; padding: 6px; text-align: center;">
                <button id="recordBtn" onclick="toggleRecord()" style="background: linear-gradient(135deg, #7c3aed, #6366f1); color: white; border: none; padding: 6px 12px; border-radius: 6px; font-weight: 600; cursor: pointer; font-size: 12px; width: 100%;">
                    🎤 Speak Now
                </button>
                <div id="statusTxt" style="color: #a855f7; font-size: 10px; margin-top: 3px;"></div>
                <textarea id="speechOutput" placeholder="Voice text appears here... copy into chat input below." style="width: 100%; height: 32px; background-color: #050508; color: #ffffff; border: 1px solid #1f222e; border-radius: 4px; padding: 4px; font-size: 11px; resize: none; margin-top: 3px;"></textarea>
            </div>
            <script>
                var recognition;
                var isRecording = false;
                function toggleRecord() {
                    if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) {
                        alert("Use Google Chrome.");
                        return;
                    }
                    if (!isRecording) {
                        var SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                        recognition = new SpeechRecognition();
                        recognition.continuous = true;
                        recognition.interimResults = true;
                        recognition.lang = 'hi-IN';
                        recognition.onstart = function() {
                            isRecording = true;
                            document.getElementById('recordBtn').innerText = "🛑 Stop";
                            document.getElementById('recordBtn').style.background = "#ef4444";
                            document.getElementById('statusTxt').innerText = "Listening...";
                        };
                        recognition.onresult = function(event) {
                            var transcript = '';
                            for (var i = event.resultIndex; i < event.results.length; ++i) {
                                transcript += event.results[i][0].transcript;
                            }
                            document.getElementById('speechOutput').value = transcript;
                        };
                        recognition.onerror = function(event) {
                            document.getElementById('statusTxt').innerText = "Error: " + event.error;
                        };
                        recognition.onend = function() {
                            isRecording = false;
                            document.getElementById('recordBtn').innerText = "🎤 Speak Now";
                            document.getElementById('recordBtn').style.background = "linear-gradient(135deg, #7c3aed, #6366f1)";
                            document.getElementById('statusTxt').innerText = "Done!";
                        };
                        recognition.start();
                    } else {
                        recognition.stop();
                    }
                }
            </script>
            """
            st.components.v1.html(voice_html, height=100)

        # 🖼️ VISUAL ATTACHMENT PREVIEW CHIP
        if format_file or case_file:
            st.markdown("##### 📎 Attached Files Ready to Submit:")
            col_p1, col_p2 = st.columns(2)
            with col_p1:
                if format_file:
                    st.markdown(f'<div class="attachment-chip">📐 Template: {format_file.name}</div>', unsafe_allow_html=True)
            with col_p2:
                if case_file:
                    st.markdown(f'<div class="attachment-chip">📸 Case Doc/Image: {case_file.name}</div>', unsafe_allow_html=True)
                    ext = case_file.name.split('.')[-1].lower()
                    if ext in ["jpg", "jpeg", "png"]:
                        st.image(Image.open(case_file), width=120)

        # 💬 1. CHAT INPUT WITH STICKY SEND (↑) ARROW
        user_input = st.chat_input("💬 Ask Nyaya AI — Type your instructions here and click Send (↑)...")

        # EXECUTE CHAT WHEN USER CLICKS SEND ARROW
        if user_input or ((format_file or case_file) and user_input is not None):
            format_content, _ = read_uploaded_file_content(format_file) if format_file else ("", None)
            case_content, case_image = read_uploaded_file_content(case_file) if case_file else ("", None)

            combined_prompt = ""
            if format_content:
                combined_prompt += f"\n\n--- CUSTOM FORMAT TEMPLATE ATTACHED ---\n{format_content}\n"
            if case_content:
                combined_prompt += f"\n\n--- CASE FILES & EVIDENCE ATTACHED ---\n{case_content}\n"
                
            prompt_text = user_input if user_input else "Perform deep legal analysis on attached files and build court draft."
            combined_prompt += f"\n\nUSER INSTRUCTION: {prompt_text}"

            save_chat_to_supabase(st.session_state.current_session, "user", prompt_text)

            with st.chat_message("user"):
                st.markdown(prompt_text)
                if format_file:
                    st.info(f"📐 Custom Format Template: `{format_file.name}`")
                if case_file:
                    st.info(f"📸 Case Document / Photo: `{case_file.name}`")
                    if case_image:
                        st.image(case_image, width=300)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = ""
                try:
                    prompt_payload = [combined_prompt]
                    if case_image:
                        prompt_payload.append(case_image)
                        
                    response = model.generate_content(prompt_payload, stream=True)
                    for chunk in response:
                        full_response += chunk.text
                        message_placeholder.markdown(full_response + "▌")
                    message_placeholder.markdown(full_response)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
                
            if full_response:
                save_chat_to_supabase(st.session_state.current_session, "assistant", full_response)
                st.session_state.uploader_key += 1
                st.rerun()
else:
    st.info("👈 Server API Key missing. Please check secrets configuration.")
