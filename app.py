import streamlit as st
import google.generativeai as genai
import datetime
import os
from supabase import create_client, Client

st.set_page_config(page_title="Nyaya Assist AI", page_icon="⚖️", layout="wide", initial_sidebar_state="collapsed")

default_api_key = st.secrets.get("GEMINI_API_KEY", "")
supabase_url = st.secrets.get("SUPABASE_URL", "")
supabase_key = st.secrets.get("SUPABASE_KEY", "")

supabase = None
if supabase_url and supabase_key:
    try:
        supabase = create_client(supabase_url, supabase_key)
    except Exception as e:
        st.error(f"Supabase connection error: {e}")

# 🎨 CHATGPT EXACT SIDEBAR COMPONENT STYLING
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .stDeployButton {display:none;}
    .stApp { background-color: #050508 !important; color: #f3f4f6 !important; font-family: 'Inter', sans-serif; }
    
    .stTextInput > div > div > input, .stSelectbox > div > div, .stTextArea > div > div > textarea { 
        background-color: #0e1017 !important; color: #ffffff !important; border: 1px solid #1f222e !important; border-radius: 12px !important; 
    }
    .stButton > button { 
        background: linear-gradient(135deg, #7c3aed, #6366f1) !important; color: white !important; border-radius: 10px !important; border: none !important; padding: 8px 16px !important; font-weight: 600 !important; width: 100%; 
    }
</style>
""", unsafe_allow_html=True)

if "user" not in st.session_state: st.session_state.user = None
if "auth_mode" not in st.session_state: st.session_state.auth_mode = "login"
if "current_session" not in st.session_state: st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
if "sidebar_open" not in st.session_state: st.session_state.sidebar_open = False

params = st.query_params
if "logged_in_user" in params and not st.session_state.user:
    saved_email = params["logged_in_user"]
    class PersistentUser:
        def __init__(self, email):
            self.email = email
            self.user_metadata = {"full_name": email.split('@')[0].capitalize()}
    st.session_state.user = PersistentUser(saved_email)

def login_user(email, password):
    try:
        res = supabase.auth.sign_in_with_password({"email": email, "password": password})
        st.session_state.user = res.user
        st.query_params["logged_in_user"] = email
        st.success("Login successful!")
        st.rerun()
    except Exception as e:
        st.error(f"Login failed: {e}")

if not st.session_state.user:
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 4, 1])
    with col2:
        st.markdown("<h2 style='text-align:center; color:white;'>⚖️ Nyaya Assist AI</h2>", unsafe_allow_html=True)
        email = st.text_input("Email", key="l_email")
        password = st.text_input("Password", type="password", key="l_pass")
        if st.button("Login", use_container_width=True): login_user(email, password)
    st.stop()

current_user_email = st.session_state.user.email
display_user_name = current_user_email.split('@')[0].capitalize()
user_initials = "".join([part[0].upper() for part in display_user_name.split()[:2]])

@st.dialog("👤 Account & Subscription")
def show_profile_modal():
    st.markdown(f"**Email:** `{current_user_email}`")
    st.markdown("**Plan:** `Free Tier` 👑")
    if st.button("Logout", use_container_width=True):
        if "logged_in_user" in st.query_params: del st.query_params["logged_in_user"]
        st.session_state.user = None; st.rerun()

# 🍔 TOP BAR WITH TOGGLE HAMBURGER BUTTON
col1, col2, col3 = st.columns([1, 6, 1])
with col1:
    if st.button("☰", key="toggle_btn"):
        st.session_state.sidebar_open = not st.session_state.sidebar_open
        st.rerun()
with col2:
    st.markdown("### Nyaya Assist <span style='color:#a855f7;'>AI</span>", unsafe_allow_html=True)
with col3: 
    if st.button(user_initials, key="prof_btn"): show_profile_modal()

# 📂 CHATGPT SIDEBAR DRAWER (VIA STREAMLIT NATIVE SIDEBAR OVERRIDE)
if st.session_state.sidebar_open:
    with st.sidebar:
        st.markdown("""
        <div style="background: linear-gradient(135deg, #1e1b4b, #0f172a); border: 1px solid #3b0764; border-radius: 12px; padding: 10px; margin-bottom: 15px;">
            <div style="font-weight: 700; font-size: 12px; color: #fff;">👑 Upgrade to Pro</div>
            <div style="font-size: 10px; color: #9ca3af;">Unlock advanced features</div>
        </div>
        """, unsafe_allow_html=True)
        
        if st.button("➕ New Case", use_container_width=True):
            st.session_state.current_session = f"Case_{datetime.datetime.now().strftime('%d%b_%H%M')}"
            st.rerun()
            
        st.markdown("---")
        st.markdown("<span style='font-size:11px; color:#8e92a4; font-weight:700;'>TODAY</span>", unsafe_allow_html=True)
        st.markdown("<div style='font-size:13px; color:#d1d5db; padding:6px 0; cursor:pointer;'>💬 Understanding Section 498A</div>", unsafe_allow_html=True)
        st.markdown("<div style='font-size:13px; color:#d1d5db; padding:6px 0; cursor:pointer;'>💬 Anticipatory Bail Grounds</div>", unsafe_allow_html=True)
        
        st.markdown("<span style='font-size:11px; color:#8e92a4; font-weight:700; margin-top:10px; display:block;'>YESTERDAY</span>", unsafe_allow_html=True)
        st.markdown("<div style='font-size:13px; color:#d1d5db; padding:6px 0; cursor:pointer;'>💬 Cheque Bounce Notice Format</div>", unsafe_allow_html=True)

        st.markdown("---")
        if st.button("📂 View All Chats", use_container_width=True): pass
        if st.button("🚪 Logout", use_container_width=True):
            if "logged_in_user" in st.query_params: del st.query_params["logged_in_user"]
            st.session_state.user = None; st.rerun()

if default_api_key:
    genai.configure(api_key=default_api_key)
    model = genai.GenerativeModel(model_name="gemini-3.6-flash", generation_config={"temperature": 0.0})

    st.markdown("### Case Studio")
    user_input = st.chat_input("Ask Nyaya AI...")
    if user_input:
        with st.chat_message("user"): st.markdown(user_input)
        with st.chat_message("assistant"):
            resp = model.generate_content(user_input).text
            st.markdown(resp)
