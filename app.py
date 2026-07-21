import streamlit as st
import google.generativeai as genai
import streamlit.components.v1 as components
import sqlite3
import datetime
import urllib.parse
import io
import os
import json
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

st.set_page_config(page_title="NyayaAI Studio", page_icon="⚖️", layout="centered")

TEMPLATES_FOLDER = "master_templates"
if not os.path.exists(TEMPLATES_FOLDER):
    os.makedirs(TEMPLATES_FOLDER)

# 🗄️ Database Setup
def init_db():
    conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS chats 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, session_name TEXT, role TEXT, content TEXT, timestamp TEXT)''')
    conn.commit()
    conn.close()

init_db()

def get_sessions():
    conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("SELECT DISTINCT session_name FROM chats ORDER BY id DESC")
    sessions = [row[0] for row in c.fetchall()]
    conn.close()
    return sessions

def load_chat_history(session_name):
    conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("SELECT role, content FROM chats WHERE session_name = ? ORDER BY id ASC", (session_name,))
    rows = c.fetchall()
    conn.close()
    return [{"role": row[0], "content": row[1]} for row in rows]

def save_message(session_name, role, content):
    conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
    c = conn.cursor()
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    c.execute("INSERT INTO chats (session_name, role, content, timestamp) VALUES (?, ?, ?, ?)", (session_name, role, content, timestamp))
    conn.commit()
    conn.close()

def search_chats(keyword):
    conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("SELECT DISTINCT session_name, content FROM chats WHERE content LIKE ?", ('%' + keyword + '%',))
    results = c.fetchall()
    conn.close()
    return results

def export_all_data():
    conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
    c = conn.cursor()
    c.execute("SELECT session_name, role, content, timestamp FROM chats ORDER BY id ASC")
    rows = c.fetchall()
    conn.close()
    data = []
    for r in rows:
        data.append({"session_name": r[0], "role": r[1], "content": r[2], "timestamp": r[3]})
    return json.dumps(data, ensure_ascii=False, indent=2)

def import_all_data(json_str):
    try:
        data = json.loads(json_str)
        conn = sqlite3.connect("nyaya_chats.db", check_same_thread=False)
        c = conn.cursor()
        for item in data:
            c.execute("INSERT INTO chats (session_name, role, content, timestamp) VALUES (?, ?, ?, ?)",
                      (item["session_name"], item["role"], item["content"], item.get("timestamp", "")))
        conn.commit()
        conn.close()
        return True
    except Exception:
        return False

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

st.sidebar.title("⚖️ NyayaAI Control Panel")
menu = st.sidebar.radio("Navigation", ["💬 Case Studio", "📂 Case History & Search", "⚙️ Settings"])

default_api_key = ""
if "GEMINI_API_KEY" in st.secrets:
    default_api_key = st.secrets["GEMINI_API_KEY"]

api_key = st.sidebar.text_input("Enter Gemini API Key", value=default_api_key, type="password")

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
    
    CRITICAL DRAFTING RULES:
    1. MULTI-TEMPLATE SELECTION: Match requested doc with Database Master Templates.
    2. ACCURACY & FLEXIBILITY: Write detailed numbered paragraphs based on case facts.
    
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
        st.header("⚙️ Application Settings & Format Database")
        
        st.subheader("👤 Profile & Court Information")
        adv_name_input = st.text_input("Advocate Name:", value=st.session_state.advocate_name)
        court_input = st.text_input("Default Court Name:", value=st.session_state.default_court)
        
        selected_lang = st.selectbox("Preferred AI Language & Style", [
            "Pure Hindi (Formal Legal)",
            "Hinglish / Bilingual (Hindi & English)",
            "Pure English (Professional Legal)"
        ], index=0)
        
        if st.button("💾 Save Settings"):
            st.session_state.advocate_name = adv_name_input
            st.session_state.default_court = court_input
            st.session_state.app_lang = selected_lang
            st.success("Settings updated!")
            
        st.markdown("---")
        st.subheader("📁 Bulk Upload Court Formats to AI Database")
        master_files = st.file_uploader("Upload Master Legal Templates (.docx / .txt)", type=["docx", "txt"], accept_multiple_files=True, key="settings_bulk_uploader")
        
        if st.button("💾 Save Templates to AI Database"):
            if master_files:
                for mf in master_files:
                    save_path = os.path.join(TEMPLATES_FOLDER, mf.name)
                    with open(save_path, "wb") as f:
                        f.write(mf.read())
                st.success("✅ Formats saved to database!")
                st.rerun()

        st.markdown("---")
        if os.path.exists(TEMPLATES_FOLDER):
            files = os.listdir(TEMPLATES_FOLDER)
            if files:
                for f in files:
                    col_f1, col_f2 = st.columns([4, 1])
                    with col_f1:
                        st.code(f"📄 {f}")
                    with col_f2:
                        if st.button("❌ Delete", key=f"del_{f}"):
                            os.remove(os.path.join(TEMPLATES_FOLDER, f))
                            st.rerun()

    elif menu == "📂 Case History & Search":
        st.header("📂 Past Case History & Backup Studio")
        
        col_bk1, col_bk2 = st.columns(2)
        with col_bk1:
            st.subheader("📥 Export / Backup History")
            backup_data = export_all_data()
            st.download_button(
                label="⬇️ Download History Backup (.json)",
                data=backup_data,
                file_name=f"NyayaAI_Backup_{datetime.datetime.now().strftime('%Y%m%d')}.json",
                mime="application/json"
            )
        
        with col_bk2:
            st.subheader("📤 Import Backup File")
            uploaded_backup = st.file_uploader("Upload Backup (.json)", type=["json"], key="history_importer")
            if uploaded_backup is not None:
                if st.button("🔄 Restore History"):
                    content = uploaded_backup.read().decode("utf-8")
                    if import_all_data(content):
                        st.success("✅ History Restored Successfully!")
                        st.rerun()
                    else:
                        st.error("❌ Failed to restore history.")

        st.markdown("---")
        search_query = st.text_input("🔍 Search past cases:")
        if search_query:
            matches = search_chats(search_query)
            for sess, content in matches:
                with st.expander(f"📁 Case: {sess}"):
                    st.write(content[:300] + "...")

    elif menu == "💬 Case Studio":
        st.title("⚖️ NyayaAI: Precision Legal Studio")
        
        sessions = get_sessions()
        if "current_session" not in st.session_state:
            st.session_state.current_session = sessions[0] if sessions else f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"

        col1, col2 = st.columns([3, 1])
        with col1:
            current_chat_name = st.selectbox("Active Case Session", sessions if sessions else [st.session_state.current_session], index=0)
            st.session_state.current_session = current_chat_name
        with col2:
            if st.button("➕ New Case"):
                new_name = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
                save_message(new_name, "assistant", "New case session initialized.")
                st.session_state.current_session = new_name
                st.rerun()

        st.markdown("---")

        messages = load_chat_history(current_chat_name)

        for idx, message in enumerate(messages):
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                
                upper_content = message["content"].upper()
                if message["role"] == "assistant" and "START_DRAFT" in upper_content:
                    col_a, col_b = st.columns([1, 1])
                    
                    with col_a:
                        tts_html = f"""
                        <button id="tts-{idx}" style="padding: 5px 10px; font-size: 11px; border-radius: 4px; border: none; background-color: #4A90E2; color: white; cursor: pointer;">🔊 Listen</button>
                        <script>
                            document.getElementById('tts-{idx}').onclick = () => {{
                                window.speechSynthesis.cancel();
                                const u = new SpeechSynthesisUtterance(`{message["content"].replace('`', '\\`').replace('$', '\\$')}`);
                                u.lang = 'hi-IN'; window.speechSynthesis.speak(u);
                            }};
                        </script>
                        """
                        components.html(tts_html, height=30)
                    
                    with col_b:
                        docx_data = create_court_ready_docx(message["content"])
                        st.download_button(
                            label="📥 Download Court Word Doc",
                            data=docx_data,
                            file_name=f"Court_Draft_{idx}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"docx_{idx}"
                        )

        st.markdown("---")

        if "uploader_key" not in st.session_state:
            st.session_state.uploader_key = 0

        col_opt1, col_opt2 = st.columns(2)
        
        with col_opt1:
            st.markdown("### 📁 1. Case Evidence / Facts")
            evidence_files = st.file_uploader("Upload Evidence", type=["pdf", "txt", "png", "jpg", "jpeg", "mp3", "wav", "m4a"], accept_multiple_files=True, key=f"evidence_uploader_{st.session_state.uploader_key}")

        with col_opt2:
            st.markdown("### 📜 2. Quick New Format")
            format_file = st.file_uploader("Upload Format", type=["docx", "txt"], accept_multiple_files=False, key=f"format_uploader_{st.session_state.uploader_key}")

        st.write("🎙️ **Voice Dictation:**")
        voice_code = """
        <button id="speech-btn" style="padding: 6px 12px; background-color: #28a745; color: white; border: none; border-radius: 6px; font-weight: bold; cursor: pointer; width: 120px;">
            🎙️ Speak
        </button>
        <script>
            const btn = document.getElementById('speech-btn');
            let rec = null; let isListening = false;
            if ('webkitSpeechRecognition' in window || 'SpeechRecognition' in window) {
                const SR = window.SpeechRecognition || window.webkitSpeechRecognition;
                rec = new SR(); rec.continuous = false; rec.interimResults = false; rec.lang = 'hi-IN';
                
                rec.onstart = () => { isListening = true; btn.innerText = "🛑 Stop"; btn.style.backgroundColor = "#dc3545"; };
                rec.onresult = (e) => {
                    const text = e.results[0][0].transcript;
                    const inputEl = window.parent.document.querySelector('textarea[data-testid="stChatInputTextArea"]');
                    if(inputEl) {
                        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, "value").set;
                        nativeInputValueSetter.call(inputEl, (inputEl.value + " " + text).trim());
                        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
                    }
                };
                rec.onend = () => { isListening = false; btn.innerText = "🎙️ Speak"; btn.style.backgroundColor = "#28a745"; };
            }
            btn.onclick = () => { if(!isListening) { rec.start(); } else { rec.stop(); } };
        </script>
        """
        components.html(voice_code, height=40)

        user_input = st.chat_input("यहाँ केस की बात लिखें...")

        if user_input:
            if format_file is not None:
                save_path = os.path.join(TEMPLATES_FOLDER, format_file.name)
                with open(save_path, "wb") as f:
                    f.write(format_file.read())
                st.toast("✅ Format Saved!", icon="📜")

            prompt_parts = []
            
            if evidence_files:
                for ef in evidence_files:
                    b_data = ef.read()
                    if ef.type == "text/plain":
                        prompt_parts.append(f"NEW CASE EVIDENCE ({ef.name}):\n{b_data.decode('utf-8')}")
                    else:
                        prompt_parts.append({"mime_type": ef.type, "data": b_data})

            prompt_parts.append(user_input)

            save_message(current_chat_name, "user", user_input)
            st.session_state.uploader_key += 1

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
                save_message(current_chat_name, "assistant", full_response)
                st.rerun()
else:
    st.info("👈 Please enter your Gemini API Key in the sidebar or configure GEMINI_API_KEY in Streamlit Secrets.")
