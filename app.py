import streamlit as st
from google import genai
from google.genai import types
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

# 🔑 Secrets
default_api_key = st.secrets.get("GEMINI_API_KEY", "")
supabase_url = st.secrets.get("SUPABASE_URL", "")
supabase_key = st.secrets.get("SUPABASE_KEY", "")

# 🗄️ Supabase
supabase: Client = None
if supabase_url and supabase_key:
    try:
        supabase = create_client(supabase_url, supabase_key)
    except Exception as e:
        st.error(f"Supabase connection error: {e}")

# 🔐 UI Styling
st.markdown("""
<style>
    .stApp { background-color: #0b0c10 !important; color: #f3f4f6 !important; font-family: 'Inter', sans-serif; }
    section[data-testid="stSidebar"] { background-color: #12141a !important; border-right: 1px solid #2a2d37 !important; }
    .stTextInput > div > div > input, .stSelectbox > div > div, .stTextArea > div > div > textarea {
        background-color: #16181e !important; color: #ffffff !important; border: 1px solid #2a2d37 !important; border-radius: 10px !important;
    }
    .stButton > button {
        background: linear-gradient(135deg, #7c3aed, #6366f1) !important; color: white !important; border-radius: 10px !important; border: none !important; padding: 10px 20px !important; font-weight: 600 !important;
    }
    .stDownloadButton > button { background: #1e1b4b !important; color: #c084fc !important; border: 1px solid #a855f7 !important; border-radius: 10px !important; }
    .user-profile-card { background-color: #16181e; border: 1px solid #2a2d37; border-radius: 12px; padding: 14px; margin-bottom: 20px; }
    .chat-header-card { background: linear-gradient(135deg, #16181e, #1f1d2b); border: 1px solid #2a2d37; border-radius: 14px; padding: 16px 20px; margin-bottom: 20px; }
</style>
""", unsafe_allow_html=True)

if "user" not in st.session_state: st.session_state.user = None
if "auth_mode" not in st.session_state: st.session_state.auth_mode = "login"

def login_user(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        st.success("✅ Login successful!")
        st.rerun()
    except Exception as e: st.error(f"Login failed: {e}")

def signup_user(email, password, full_name, phone):
    try:
        res = supabase.auth.sign_up({"email": email, "password": password, "options": {"data": {"full_name": full_name, "phone": phone}}})
        st.session_state.user = res.user
        st.success("✅ Account created successfully!")
        st.rerun()
    except Exception as e: st.error(f"Signup failed: {e}")

if not st.session_state.user:
    st.markdown("<br>", unsafe_allow_html=True)
    col_auth1, col_auth2, col_auth3 = st.columns([1, 2, 1])
    with col_auth2:
        if st.session_state.auth_mode == "login":
            st.markdown("""<div style="text-align: center;"><div style="font-size: 55px;">⚖️</div><h2 style="margin: 0; font-weight: 700; color: #ffffff;">Nyaya Assist <span style="color: #a855f7;">AI</span></h2><p style="color: #9ca3af; font-size: 14px;">Your AI Legal Assistant</p><br><h4 style="margin: 0; color: #e5e7eb;">Login to your account</h4></div>""", unsafe_allow_html=True)
            email = st.text_input("Email Address", placeholder="name@example.com", key="login_email")
            password = st.text_input("Password", type="password", placeholder="••••••••", key="login_pass")
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Login  ➔", key="btn_login", use_container_width=True):
                if email and password and supabase: login_user(email, password)
                else: st.warning("Please enter Email & Password.")
            if st.button("Don't have an account? Sign up", key="switch_to_signup", use_container_width=True):
                st.session_state.auth_mode = "signup"
                st.rerun()
        else:
            st.markdown("""<div style="text-align: left;"><h2 style="margin: 0; color: #ffffff;">Create your account</h2></div>""", unsafe_allow_html=True)
            full_name = st.text_input("Full Name", placeholder="Adv. Rajesh Sharma", key="signup_name")
            email = st.text_input("Email", placeholder="name@example.com", key="signup_email")
            phone = st.text_input("Phone Number", placeholder="+91 98765 43210", key="signup_phone")
            password = st.text_input("Password", type="password", placeholder="••••••••", key="signup_pass")
            confirm_pass = st.text_input("Confirm Password", type="password", placeholder="••••••••", key="signup_cpass_unique")
            agree = st.checkbox("I agree to Terms & Privacy Policy")
            if st.button("Sign Up  ➔", key="btn_signup", use_container_width=True):
                if not agree: st.warning("Please agree to terms.")
                elif password != confirm_pass: st.error("Passwords do not match!")
                elif email and password and supabase: signup_user(email, password, full_name, phone)
                else: st.warning("Fill required fields.")
            if st.button("Already have an account? Login", key="switch_to_login", use_container_width=True):
                st.session_state.auth_mode = "login"
                st.rerun()
    st.stop()

current_user_email = st.session_state.user.email

with st.sidebar:
    st.markdown("""<div style="display: flex; align-items: center; gap: 10px; margin-bottom: 15px;"><span style="font-size: 32px;">⚖️</span><div><h3 style="margin:0; font-size: 18px; font-weight:700;">Nyaya Assist <span style="color:#a855f7;">AI</span></h3></div></div>""", unsafe_allow_html=True)
    st.markdown(f"""<div class="user-profile-card"><div style="font-size: 11px; color: #9ca3af;">LOGGED IN ADVOCATE</div><div style="font-weight: 600; font-size: 13px; color: #ffffff; word-break: break-all;">{current_user_email}</div></div>""", unsafe_allow_html=True)
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
        try: supabase.table("chats").insert({"user_email": current_user_email, "session_name": session_name, "role": role, "content": content}).execute()
        except Exception as e: st.error(f"Error saving chat: {e}")

def get_supabase_sessions():
    if supabase:
        try:
            res = supabase.table("chats").select("session_name").eq("user_email", current_user_email).execute()
            return sorted(list(set([row["session_name"] for row in res.data])), reverse=True)
        except Exception: return []
    return []

def get_supabase_chat_history(session_name):
    if supabase:
        try:
            res = supabase.table("chats").select("role, content").eq("user_email", current_user_email).eq("session_name", session_name).order("id").execute()
            return res.data
        except Exception: return []
    return []

def read_uploaded_file_content(uploaded_file):
    if uploaded_file is None: return "", None
    file_type = uploaded_file.name.split('.')[-1].lower()
    content, image_obj = "", None
    if file_type == "txt": content = uploaded_file.read().decode("utf-8")
    elif file_type == "docx":
        doc = Document(uploaded_file)
        content = '\n'.join([p.text for p in doc.paragraphs])
    elif file_type in ["jpg", "jpeg", "png"]:
        image_obj = Image.open(uploaded_file)
        content = f"[IMAGE FILE ATTACHED: {uploaded_file.name}]"
    return content, image_obj

def create_court_ready_docx(text):
    doc = Document()
    section = doc.sections[0]
    section.page_width, section.page_height = Inches(8.5), Inches(14.0)
    section.top_margin, section.bottom_margin = Inches(1.0), Inches(1.0)
    section.left_margin, section.right_margin = Inches(1.5), Inches(1.0)
    draft_content = ""
    upper_text = text.upper()
    if "START_DRAFT" in upper_text and "END_DRAFT" in upper_text:
        start_idx = upper_text.find("START_DRAFT") + len("START_DRAFT")
        end_idx = upper_text.find("END_DRAFT")
        draft_content = text[start_idx:end_idx].strip()
    else: draft_content = text.strip()

    for line in draft_content.split('\n'):
        stripped = line.strip()
        if not stripped: continue
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing, p.paragraph_format.space_after = 1.5, Pt(6)
        if stripped.startswith("##") or stripped.startswith("**"):
            run = p.add_run(stripped.replace("#", "").replace("**", "").strip())
            run.bold, run.font.size = True, Pt(14)
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
    # 🌟 NEW OFFICIAL GOOGLE GENAI CLIENT
    ai_client = genai.Client(api_key=api_key)
    
    if "app_lang" not in st.session_state: st.session_state.app_lang = "Pure Hindi (Formal Legal)"
    if "advocate_name" not in st.session_state: st.session_state.advocate_name = ""
    if "default_court" not in st.session_state: st.session_state.default_court = ""

    system_instruction = f"""
    You are a Senior Advocate & Legal Strategist of the Supreme Court of India.
    Language: {st.session_state.app_lang}.
    Advocate: {st.session_state.advocate_name if st.session_state.advocate_name else "[Advocate Name]"}
    Court: {st.session_state.default_court if st.session_state.default_court else "[Court Name]"}
    
    Always perform deep analysis, suggest loopholes, cross-questions, legal sections, and a full court draft inside:
    START_DRAFT
    ... draft ...
    END_DRAFT
    """

    if menu == "⚙️ Settings":
        st.markdown("## ⚙️ Settings")
        adv_name_input = st.text_input("Advocate Full Name:", value=st.session_state.advocate_name)
        court_input = st.text_input("Default Court Jurisdiction:", value=st.session_state.default_court)
        selected_lang = st.selectbox("Preferred AI Language Mode", ["Pure Hindi (Formal Legal)", "Hinglish / Bilingual (Hindi & English)", "Pure English (Professional Legal)"], index=0)
        if st.button("💾 Save Settings"):
            st.session_state.advocate_name, st.session_state.default_court, st.session_state.app_lang = adv_name_input, court_input, selected_lang
            st.success("✅ Settings updated!")

    elif menu == "📂 Case History":
        st.markdown("## 📂 Saved Case Archives")
        sessions = get_supabase_sessions()
        if sessions:
            for s in sessions:
                with st.expander(f"📁 Case Session: {s}"):
                    for c in get_supabase_chat_history(s):
                        st.markdown(f"**{'👤 You' if c['role']=='user' else '⚖️ NyayaAssist AI'}:**\n{c['content']}\n---")
        else: st.info("No saved cases found.")

    elif menu == "💬 Case Studio":
        sessions = get_supabase_sessions()
        if "current_session" not in st.session_state: st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
        st.markdown("""<div class="chat-header-card"><h3 style="margin: 0; color: #ffffff;">💬 Deep Legal Analysis Studio</h3></div>""", unsafe_allow_html=True)

        col_c1, col_c2 = st.columns([3, 1])
        with col_c1:
            st.session_state.current_session = st.selectbox("Active Case File:", [st.session_state.current_session] + [s for s in sessions if s != st.session_state.current_session], index=0)
        with col_c2:
            if st.button("➕ New Case", use_container_width=True):
                st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
                st.rerun()

        col_f1, col_f2 = st.columns(2)
        with col_f1: format_file = st.file_uploader("Upload Format (.docx, .txt)", type=["docx", "txt"], key="format_file")
        with col_f2: case_file = st.file_uploader("Upload Case Docs or Photos", type=["docx", "txt", "jpg", "jpeg", "png"], key="case_file")

        messages = get_supabase_chat_history(st.session_state.current_session)
        for idx, message in enumerate(messages):
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                if message["role"] == "assistant" and "START_DRAFT" in message["content"].upper():
                    st.download_button("📥 Download Court-Ready Word (.docx)", data=create_court_ready_docx(message["content"]), file_name=f"Court_Draft_{idx}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"docx_{idx}")

        # 🎙️ Voice Assistant Box
        st.markdown("##### 🎙️ Voice Typing Assistant (बोलकर टाइप करें):")
        voice_html = """
        <div style="background-color: #16181e; padding: 15px; border-radius: 12px; border: 1px solid #2a2d37;">
            <div style="display: flex; gap: 10px; align-items: center; margin-bottom: 10px;">
                <button id="recordBtn" onclick="toggleRecord()" style="background: linear-gradient(135deg, #7c3aed, #6366f1); color: white; border: none; padding: 10px 18px; border-radius: 8px; font-weight: 600; cursor: pointer;">🎤 स्टार्ट वॉइस रिकॉर्डिंग (Hindi/English)</button>
                <span id="statusTxt" style="color: #a855f7; font-size: 13px; font-weight: 500;"></span>
            </div>
            <textarea id="speechOutput" placeholder="जो आप बोलेंगे वो यहाँ लाइव टाइप होगा... फिर इसे नीचे दिए गए बॉक्स में कॉपी करके भेजें।" style="width: 100%; height: 80px; background-color: #0b0c10; color: #ffffff; border: 1px solid #2a2d37; border-radius: 8px; padding: 10px; font-size: 14px; resize: none;"></textarea>
        </div>
        <script>
            var recognition; var isRecording = false;
            function toggleRecord() {
                if (!('webkitSpeechRecognition' in window) && !('SpeechRecognition' in window)) { alert("ब्राउज़र वॉइस सपोर्ट नहीं करता। Chrome इस्तेमाल करें।"); return; }
                if (!isRecording) {
                    var SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
                    recognition = new SpeechRecognition(); recognition.continuous = true; recognition.interimResults = true; recognition.lang = 'hi-IN';
                    recognition.onstart = function() { isRecording = true; document.getElementById('recordBtn').innerText = "🛑 रिकॉर्डिंग रोकें"; document.getElementById('recordBtn').style.background = "#ef4444"; document.getElementById('statusTxt').innerText = "🎙️ बोलिए..."; };
                    recognition.onresult = function(event) {
                        var transcript = ''; for (var i = event.resultIndex; i < event.results.length; ++i) { transcript += event.results[i][0].transcript; }
                        document.getElementById('speechOutput').value = transcript;
                    };
                    recognition.onend = function() { isRecording = false; document.getElementById('recordBtn').innerText = "🎤 स्टार्ट वॉइस रिकॉर्डिंग"; document.getElementById('recordBtn').style.background = "linear-gradient(135deg, #7c3aed, #6366f1)"; document.getElementById('statusTxt').innerText = "✅ कॉपी करें!"; };
                    recognition.start();
                } else { recognition.stop(); }
            }
        </script>
        """
        st.components.v1.html(voice_html, height=160)

        user_input = st.chat_input("यहाँ केस के तथ्य, निर्देश या वॉइस से बोला हुआ टेक्स्ट पेस्ट करें...")
        if user_input or format_file or case_file:
            format_content, _ = read_uploaded_file_content(format_file) if format_file else ("", None)
            case_content, case_image = read_uploaded_file_content(case_file) if case_file else ("", None)

            combined_prompt = ""
            if format_content: combined_prompt += f"\n\n--- FORMAT TEMPLATE ---\n{format_content}\n"
            if case_content: combined_prompt += f"\n\n--- CASE FILES ---\n{case_content}\n"
            combined_prompt += f"\n\nUSER INSTRUCTION: {user_input if user_input else 'Analyze files and generate draft.'}"

            save_chat_to_supabase(st.session_state.current_session, "user", user_input if user_input else "Uploaded Files for Analysis")

            with st.chat_message("user"):
                if user_input: st.markdown(user_input)
                if format_file: st.info(f"📐 Format: `{format_file.name}`")
                if case_file:
                    st.info(f"📸 Case File: `{case_file.name}`")
                    if case_image: st.image(case_image, width=300)

            with st.chat_message("assistant"):
                message_placeholder = st.empty()
                full_response = ""
                try:
                    contents_payload = [combined_prompt]
                    if case_image: contents_payload.append(case_image)
                    
                    # 🌟 New SDK Endpoint Call (Prevents 404 permanently)
                    response = ai_client.models.generate_content_stream(
                        model='gemini-2.5-flash',
                        contents=contents_payload,
                        config=types.GenerateContentConfig(system_instruction=system_instruction, temperature=0.0)
                    )
                    for chunk in response:
                        full_response += chunk.text
                        message_placeholder.markdown(full_response + "▌")
                    message_placeholder.markdown(full_response)
                except Exception as e: st.error(f"Error: {str(e)}")

            if full_response:
                save_chat_to_supabase(st.session_state.current_session, "assistant", full_response)
                st.rerun()
else:
    st.info("👈 Please enter your Gemini API Key in the sidebar settings.")
