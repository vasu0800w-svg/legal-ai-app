import streamlit as st
import google.generativeai as genai
import datetime
import io
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from supabase import create_client, Client

st.set_page_config(page_title="Nyaya Assist AI", page_icon="⚖️", layout="centered")

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

# 🔐 STYLING FOR MODERN DARK UI
st.markdown("""
<style>
    .stApp {
        background-color: #0b0c10;
        color: #ffffff;
    }
    .stTextInput > div > div > input {
        background-color: #16181e !important;
        color: #ffffff !important;
        border: 1px solid #2a2d37 !important;
        border-radius: 12px !important;
        padding: 12px 15px !important;
    }
    .stTextInput > div > div > input:focus {
        border-color: #8b5cf6 !important;
        box-shadow: 0 0 8px rgba(139, 92, 246, 0.4) !important;
    }
    .stButton > button {
        background: linear-gradient(135deg, #7c3aed, #6366f1) !important;
        color: white !important;
        border-radius: 12px !important;
        border: none !important;
        padding: 12px 24px !important;
        font-weight: 600 !important;
        font-size: 16px !important;
        width: 100% !important;
        box-shadow: 0 4px 12px rgba(124, 58, 237, 0.3) !important;
    }
    .stButton > button:hover {
        background: linear-gradient(135deg, #6d28d9, #4f46e5) !important;
    }
    .divider {
        display: flex;
        align-items: center;
        text-align: center;
        color: #6b7280;
        margin: 20px 0;
    }
    .divider::before, .divider::after {
        content: '';
        flex: 1;
        border-bottom: 1px solid #2a2d37;
    }
    .divider:not(:empty)::before { margin-right: .25em; }
    .divider:not(:empty)::after { margin-left: .25em; }
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

def get_google_auth_url():
    if supabase:
        try:
            # Using official Supabase SDK function for OAuth URL
            res = supabase.auth.sign_in_with_oauth({
                "provider": "google",
                "options": {
                    "redirect_to": "https://nyaya-ai-studio.streamlit.app"
                }
            })
            return res.url
        except Exception as e:
            return None
    return None

# ----------------- LOGIN / SIGNUP SCREEN -----------------
if not st.session_state.user:
    st.markdown("<br>", unsafe_allow_html=True)
    google_url = get_google_auth_url()
    
    if st.session_state.auth_mode == "login":
        st.markdown("""
        <div style="text-align: center;">
            <div style="font-size: 50px;">⚖️</div>
            <h2 style="margin: 0; font-weight: 700; color: #ffffff;">Nyaya Assist <span style="color: #a855f7;">AI</span></h2>
            <p style="color: #9ca3af; font-size: 14px; margin-top: 4px;">Your AI Legal Assistant</p>
            <br>
            <h4 style="margin: 0; font-weight: 600;">Login to your account</h4>
            <p style="color: #6b7280; font-size: 13px;">Welcome back! Please enter your details.</p>
        </div>
        """, unsafe_allow_html=True)
        
        email = st.text_input("Email Address", placeholder="name@example.com", key="login_email")
        password = st.text_input("Password", type="password", placeholder="••••••••", key="login_pass")
        
        if st.button("Login  ➔", key="btn_login"):
            if email and password and supabase:
                login_user(email, password)
            else:
                st.warning("Please enter Email & Password.")
                
        st.markdown("<div class='divider'>or continue with</div>", unsafe_allow_html=True)
        
        if google_url:
            st.markdown(f'''
            <a href="{google_url}" target="_self" style="text-decoration: none;">
                <div style="background-color: #16181e; border: 1px solid #2a2d37; border-radius: 12px; padding: 12px; text-align: center; color: white; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 10px;">
                    🌐 Continue with Google
                </div>
            </a>
            ''', unsafe_allow_html=True)
        else:
            st.error("Google Auth unavailable.")
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
        with col_s2:
            if st.button("Don't have an account? Sign up", key="switch_to_signup"):
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
        
        if st.button("Sign Up  ➔", key="btn_signup"):
            if not agree:
                st.warning("Please agree to the Terms & Privacy Policy.")
            elif password != confirm_pass:
                st.error("Passwords do not match!")
            elif email and password and supabase:
                signup_user(email, password, full_name, phone)
            else:
                st.warning("Please fill all required fields.")
                
        st.markdown("<div class='divider'>or continue with</div>", unsafe_allow_html=True)
        
        if google_url:
            st.markdown(f'''
            <a href="{google_url}" target="_self" style="text-decoration: none;">
                <div style="background-color: #16181e; border: 1px solid #2a2d37; border-radius: 12px; padding: 12px; text-align: center; color: white; font-weight: 600; cursor: pointer; display: flex; align-items: center; justify-content: center; gap: 10px;">
                    🌐 Continue with Google
                </div>
            </a>
            ''', unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        col_s1, col_s2, col_s3 = st.columns([1, 2, 1])
        with col_s2:
            if st.button("Already have an account? Login", key="switch_to_login"):
                st.session_state.auth_mode = "login"
                st.rerun()

    st.stop()

# ----------------- LOGGED IN APPLICATION AREA -----------------
current_user_email = st.session_state.user.email

st.sidebar.title("⚖️ Nyaya Assist AI")
st.sidebar.write(f"👤 **User:** `{current_user_email}`")
if st.sidebar.button("🚪 Logout"):
    st.session_state.user = None
    st.session_state.auth_mode = "login"
    st.rerun()

menu = st.sidebar.radio("Navigation", ["💬 Case Studio", "📂 Case History", "⚙️ Settings"])
api_key = st.sidebar.text_input("Enter Gemini API Key", value=default_api_key, type="password")

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
        st.header("⚙️ Application Settings")
        adv_name_input = st.text_input("Advocate Name:", value=st.session_state.advocate_name)
        court_input = st.text_input("Default Court Name:", value=st.session_state.default_court)
        selected_lang = st.selectbox("Preferred AI Language", [
            "Pure Hindi (Formal Legal)",
            "Hinglish / Bilingual (Hindi & English)",
            "Pure English (Professional Legal)"
        ], index=0)
        
        if st.button("💾 Save Settings"):
            st.session_state.advocate_name = adv_name_input
            st.session_state.default_court = court_input
            st.session_state.app_lang = selected_lang
            st.success("Settings updated!")

    elif menu == "📂 Case History":
        st.header("📂 Your Cloud Saved Case History")
        sessions = get_supabase_sessions()
        if sessions:
            for s in sessions:
                with st.expander(f"📁 Case: {s}"):
                    chats = get_supabase_chat_history(s)
                    for c in chats:
                        st.markdown(f"**{c['role'].capitalize()}:** {c['content']}")
        else:
            st.info("No saved cases found.")

    elif menu == "💬 Case Studio":
        st.title("⚖️ Nyaya Assist AI")
        sessions = get_supabase_sessions()
        if "current_session" not in st.session_state:
            st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"

        col1, col2 = st.columns([3, 1])
        with col1:
            options = [st.session_state.current_session] + [s for s in sessions if s != st.session_state.current_session]
            current_chat_name = st.selectbox("Active Case Session", options, index=0)
            st.session_state.current_session = current_chat_name
        with col2:
            if st.button("➕ New Case"):
                st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
                st.rerun()

        st.markdown("---")
        messages = get_supabase_chat_history(current_chat_name)

        if not messages:
            st.info("✨ **New Case Initialized.** Enter evidence or ask questions below to start drafting.")

        for idx, message in enumerate(messages):
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                upper_content = message["content"].upper()
                if message["role"] == "assistant" and "START_DRAFT" in upper_content:
                    docx_data = create_court_ready_docx(message["content"])
                    st.download_button(
                        label="📥 Download Court Word Doc",
                        data=docx_data,
                        file_name=f"Court_Draft_{idx}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"docx_{idx}"
                    )

        st.markdown("---")
        user_input = st.chat_input("यहाँ केस की बात लिखें...")

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
    st.info("👈 Gemini API Key missing.")
