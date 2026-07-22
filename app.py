import streamlit as st
import google.generativeai as genai
import streamlit.components.v1 as components
import datetime
import io
import os
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from supabase import create_client, Client

st.set_page_config(page_title="NyayaAI Studio Pro", page_icon="⚖️", layout="centered")

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

# 🔐 LOGIN & USER SYSTEM
if "user" not in st.session_state:
    st.session_state.user = None

def login_user(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        st.success("✅ Login successful!")
        st.rerun()
    except Exception as e:
        st.error(f"Login failed: {e}")

def signup_user(email, password):
    try:
        res = supabase.auth.sign_up({"email": email, "password": password})
        st.session_state.user = res.user
        st.success("✅ Account created successfully!")
        st.rerun()
    except Exception as e:
        st.error(f"Signup failed: {e}")

if not st.session_state.user:
    st.title("⚖️ NyayaAI Studio Pro")
    st.subheader("Login / Signup to Access Legal AI")
    
    auth_mode = st.radio("Choose Action", ["Login", "Sign Up"])
    email = st.text_input("Email Address")
    password = st.text_input("Password", type="password")
    
    if auth_mode == "Login":
        if st.button("🔐 Log In"):
            if email and password and supabase:
                login_user(email, password)
            else:
                st.warning("Please enter Email & Password.")
    else:
        if st.button("📝 Sign Up"):
            if email and password and supabase:
                signup_user(email, password)
            else:
                st.warning("Please enter Email & Password.")
    st.stop()

# ----------------- LOGGED IN APPLICATION AREA -----------------
current_user_email = st.session_state.user.email

st.sidebar.title("⚖️ NyayaAI Pro")
st.sidebar.write(f"👤 **User:** `{current_user_email}`")
if st.sidebar.button("🚪 Logout"):
    st.session_state.user = None
    st.rerun()

menu = st.sidebar.radio("Navigation", ["💬 Case Studio", "📂 Case History", "⚙️ Settings"])

api_key = st.sidebar.text_input("Enter Gemini API Key", value=default_api_key, type="password")

# Database Helper Functions
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
        
        st.subheader("👤 Profile Information")
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
        st.title("⚖️ NyayaAI: Precision Legal Studio")
        
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
