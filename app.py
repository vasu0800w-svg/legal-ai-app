import streamlit as st
import google.generativeai as genai
import datetime
import io
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from supabase import create_client, Client

st.set_page_config(page_title="Nyaya Assist AI", page_icon="⚖️", layout="wide")

TEMPLATES_FOLDER = "master_templates"
if not os.path.exists(TEMPLATES_FOLDER):
    os.makedirs(TEMPLATES_FOLDER)

# 🔑 Fetch Secrets
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

# 🔐 STYLING FOR UNIFIED MODERN DARK UI ACROSS WHOLE APP
st.markdown("""
<style>
    .stApp {
        background-color: #0b0c10 !important;
        color: #f3f4f6 !important;
        font-family: 'Inter', sans-serif;
    }
    section[data-testid="stSidebar"] {
        background-color: #12141a !important;
        border-right: 1px solid #2a2d37 !important;
    }
    .stTextInput > div > div > input, .stSelectbox > div > div {
        background-color: #16181e !important;
        color: #ffffff !important;
        border: 1px solid #2a2d37 !important;
        border-radius: 10px !important;
    }
    .stTextInput > div > div > input:focus {
        border-color: #8b5cf6 !important;
        box-shadow: 0 0 8px rgba(139, 92, 246, 0.4) !important;
    }
    .stButton > button {
        background: linear-gradient(135deg, #7c3aed, #6366f1) !important;
        color: white !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 10px 20px !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 12px rgba(124, 58, 237, 0.25) !important;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(124, 58, 237, 0.4) !important;
    }
    .stDownloadButton > button {
        background: #1e1b4b !important;
        color: #c084fc !important;
        border: 1px solid #a855f7 !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
    }
    .stDownloadButton > button:hover {
        background: #a855f7 !important;
        color: white !important;
    }
    .user-profile-card {
        background-color: #16181e;
        border: 1px solid #2a2d37;
        border-radius: 12px;
        padding: 14px;
        margin-bottom: 20px;
    }
    .chat-header-card {
        background: linear-gradient(135deg, #16181e, #1f1d2b);
        border: 1px solid #2a2d37;
        border-radius: 14px;
        padding: 16px 20px;
        margin-bottom: 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }
    .streamlit-expanderHeader {
        background-color: #16181e !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        border: 1px solid #2a2d37 !important;
    }
</style>
""", unsafe_allow_html=True)

# 🔐 USER SYSTEM STATE
if "user" not in st.session_state:
    st.session_state.user = None
if "auth_mode" not in st.session_state:
    st.session_state.auth_mode = "login"

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
                <p style="color: #6b7280; font-size: 13px;">Welcome back! Please enter your details.</p>
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
                <h2 style="margin: 0; font-weight: 700; color: #ffffff;">Create your account</h2>
                <p style="color: #9ca3af; font-size: 14px; margin-top: 4px;">Join Nyaya Assist AI and simplify legal tasks with AI.</p>
            </div>
            """, unsafe_allow_html=True)
            
            full_name = st.text_input("Full Name", placeholder="Adv. Rajesh Sharma", key="signup_name")
            email = st.text_input("Email", placeholder="name@example.com", key="signup_email")
            phone = st.text_input("Phone Number (Optional)", placeholder="+91 98765 43210", key="signup_phone")
            password = st.text_input("Password", type="password", placeholder="••••••••", key="signup_pass")
            confirm_pass = st.text_input("Confirm Password", type="password", placeholder="••••••••", key="signup_cpass")
            
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

with st.sidebar:
    st.markdown("""
    <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 15px;">
        <span style="font-size: 32px;">⚖️</span>
        <div>
            <h3 style="margin:0; font-size: 18px; font-weight:700;">Nyaya Assist <span style="color:#a855f7;">AI</span></h3>
            <span style="font-size:11px; color:#9ca3af;">Legal Drafting Studio</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div class="user-profile-card">
        <div style="font-size: 11px; color: #9ca3af;">LOGGED IN ADVOCATE</div>
        <div style="font-weight: 600; font-size: 13px; color: #ffffff; word-break: break-all;">{current_user_email}</div>
    </div>
    """, unsafe_allow_html=True)

    menu = st.radio("Navigation", ["💬 Case Studio", "📂 Case History", "⚙️ Settings"], label_visibility="collapsed")
    
    st.markdown("---")
    
    with st.expander("🔑 Gemini API Settings"):
        api_key = st.text_input("API Key", value=default_api_key, type="password")

    st.markdown("<br>"*3, unsafe_allow_html=True)
    if st.button("🚪 Log Out", use_container_width=True):
        st.session_state.user = None
        st.session_state.auth_mode = "login"
        st.rerun()

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

if api_key:
    genai.configure(api_key=api_key)
    
    if "app_lang" not in st.session_state:
        st.session_state.app_lang = "Pure Hindi (Formal Legal)"
    if "advocate_name" not in st.session_state:
        st.session_state.advocate_name = ""
    if "default_court" not in st.session_state:
        st.session_state.default_court = ""

    local_templates_text = load_all_local_templates()

    system_instruction = f"""
    ROLE & OBJECTIVE:
    You are an elite Indian Senior Advocate and Master Legal Draftsman assisting a veteran attorney.
    Response Language Style: {st.session_state.app_lang}.
    
    MANDATORY DRAFTING DETAILS:
    - Advocate Name: {st.session_state.advocate_name if st.session_state.advocate_name else "[Advocate Name]"}
    - Court Heading: {st.session_state.default_court if st.session_state.default_court else "[Court Name]"}
    
    PERMANENT MASTER TEMPLATES IN DATABASE:
    {local_templates_text if local_templates_text else "Default Standard Court Layout (High Court / District Court Format)"}
    
    OUTPUT FORMATTING:
    Final court draft MUST be strictly enclosed inside:
      START_DRAFT
      ...
      END_DRAFT
    """
    
    model = genai.GenerativeModel(
        model_name="gemini-3.5-flash",
        system_instruction=system_instruction,
        generation_config={"temperature": 0.0}
    )

    if menu == "⚙️ Settings":
        st.markdown("## ⚙️ App & Profile Settings")
        st.markdown("Configure your legal credentials for auto-filled drafting.")
        
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            adv_name_input = st.text_input("Advocate Full Name:", value=st.session_state.advocate_name, placeholder="e.g. Adv. R. K. Sharma")
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

    elif menu == "📂 Case History":
        st.markdown("## 📂 Saved Case Archives")
        st.markdown("Access all your cloud-backed legal discussions and generated drafts.")
        
        sessions = get_supabase_sessions()
        if sessions:
            for s in sessions:
                with st.expander(f"📁 Case Session: {s}"):
                    chats = get_supabase_chat_history(s)
                    for c in chats:
                        role_label = "👤 You" if c['role'] == "user" else "⚖️ NyayaAssist AI"
                        st.markdown(f"**{role_label}:**")
                        st.markdown(c['content'])
                        st.markdown("---")
        else:
            st.info("No saved cases found in your account history.")

    elif menu == "💬 Case Studio":
        sessions = get_supabase_sessions()
        if "current_session" not in st.session_state:
            st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"

        st.markdown("""
        <div class="chat-header-card">
            <div>
                <h3 style="margin: 0; font-size: 20px; font-weight: 700; color: #ffffff;">💬 Precision Legal Studio</h3>
                <p style="margin: 0; font-size: 13px; color: #9ca3af;">AI-Powered Case Analysis & Court-Ready Drafting</p>
            </div>
        </div>
        """, unsafe_allow_html=True)

        col_c1, col_c2 = st.columns([3, 1])
        with col_c1:
            options = [st.session_state.current_session] + [s for s in sessions if s != st.session_state.current_session]
            current_chat_name = st.selectbox("Active Case File:", options, index=0)
            st.session_state.current_session = current_chat_name
        with col_c2:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            if st.button("➕ New Case", use_container_width=True):
                st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
                st.rerun()

        st.markdown("<br>", unsafe_allow_html=True)

        messages = get_supabase_chat_history(current_chat_name)

        if not messages:
            st.markdown("""
            <div style="text-align: center; padding: 40px; background-color: #16181e; border: 1px dashed #2a2d37; border-radius: 14px;">
                <div style="font-size: 40px;">⚖️</div>
                <h4 style="margin: 10px 0 5px 0; color: #ffffff;">New Case File Active</h4>
                <p style="color: #6b7280; font-size: 14px;">Enter case facts, upload evidence, or ask legal questions below to generate formal court drafts.</p>
            </div>
            """, unsafe_allow_html=True)

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

        st.markdown("<br>", unsafe_allow_html=True)

        user_input = st.chat_input("केस के तथ्य, धाराएं या ड्राफ्टिंग विवरण लिखें...")

        if user_input:
            prompt_parts = [user_input]
            save_chat_to_supabase(current_chat_name, "user", user_input)

            with st.chat_message("user"):
                st.markdown(user_input)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = ""
                try:
                    response = model.generate_content(prompt_parts, stream=True)
                    for chunk in response:
                        full_response += chunk.text
                        message_placeholder.markdown(full_response + "▌")
                    message_placeholder.markdown(full_response)
                except Exception as e:
                    st.error(f"Error: {str(e)}")
                
            if full_response:
                save_chat_to_supabase(current_chat_name, "assistant", full_response)
                st.rerun()
else:
    st.info("👈 Please enter your Gemini API Key in the sidebar settings.")
