

import streamlit as st
import json
import time
from datetime import datetime
from pathlib import Path
from report import WeeklyReportGenerator

# --- PAGE CONFIG ---
st.set_page_config(
    page_title="AoD Weekly Newsletter",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# --- CUSTOM CSS ---
st.markdown("""
<style>
    .main > div {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    .hero-section {
        background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
        color: white;
        padding: 2rem;
        border-radius: 8px;
        text-align: center;
        margin-bottom: 2rem;
    }
    
    .hero-title {
        font-size: 2.2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .hero-subtitle {
        font-size: 1.1rem;
        opacity: 0.9;
    }
    
    .step-card {
        background: #f8f9fa;
        border: 1px solid #e1e8ed;
        border-radius: 6px;
        padding: 1.5rem;
        margin-bottom: 1rem;
    }
    
    .step-number {
        background: #2c3e50;
        color: white;
        border-radius: 50%;
        width: 28px;
        height: 28px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font-weight: 600;
        font-size: 0.9rem;
        margin-right: 12px;
    }
    
    .success-box {
        background: #d4edda;
        color: #155724;
        border: 1px solid #c3e6cb;
        padding: 1rem;
        border-radius: 6px;
        margin: 1rem 0;
        text-align: center;
    }
    
    .error-box {
        background: #f8d7da;
        color: #721c24;
        border: 1px solid #f5c6cb;
        padding: 1rem;
        border-radius: 6px;
        margin: 1rem 0;
        text-align: center;
    }
    
    .upload-section {
        background: white;
        border: 2px dashed #2c3e50;
        border-radius: 8px;
        padding: 2rem;
        text-align: center;
        margin: 2rem 0;
    }
    
    .info-section {
        background: #f8f9fa;
        border-left: 4px solid #2c3e50;
        padding: 1rem;
        margin: 1rem 0;
        border-radius: 4px;
    }
    
    .instructions-section {
        background: #f8f9fa;
        border: 1px solid #e1e8ed;
        border-radius: 6px;
        padding: 1.5rem;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# --- HERO SECTION ---
st.markdown("""
<div class="hero-section">
    <div class="hero-title">AoD Weekly Newsletter</div>
    <div class="hero-subtitle">Generate comprehensive weekly reports</div>
</div>
""", unsafe_allow_html=True)

# --- INSTRUCTIONS SECTION ---
st.markdown("## Getting Started")

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="step-card">
        <div style="display: flex; align-items: center; margin-bottom: 1rem;">
            <div class="step-number">1</div>
            <strong>Login to Canvas</strong>
        </div>
        <p>Open Canvas in a new tab and log in to your account.</p>
    </div>
    """, unsafe_allow_html=True)
    
    # st.markdown('[Open Canvas Login](https://canvas.artofdrawers.com)')

with col2:
    st.markdown("""
    <div class="step-card">
        <div style="display: flex; align-items: center; margin-bottom: 1rem;">
            <div class="step-number">2</div>
            <strong>Upload Cookies</strong>
        </div>
        <p>Return to this tab and upload your cookies.json file below.</p>
    </div>
    """, unsafe_allow_html=True)

# --- COOKIE INSTRUCTIONS EXPANDER ---
with st.expander("How to get your cookies.json file", expanded=False):
    st.markdown("**Follow these steps to generate your cookies.json file:**")
    
    steps = [
        "Open [Canvas in Chrome](https://canvas.artofdrawers.com)",
        "Install the [Cookie-Editor Extension](https://chrome.google.com/webstore/detail/cookie-editor/hlkenndedbemlkljdomclgjgkkdggpac)",
        "In Canvas, click on the Cookie-Editor Extension",
        "Keep only the **PHPSESSID** and **username** cookies",
        "Click the **Export** button at the top of the extension",
        "Visit [JSON Editor Online](https://jsoneditoronline.org/)",
        "Paste the cookies into the 'New Document 1' text box on the left",
        "Click the **save** button, then 'Save to Disk'",
        "Give it a name and click **Save**",
        "You now have your cookies.json file!"
    ]
    
    for i, step in enumerate(steps, 1):
        st.write(f"{i}. {step}")

# --- UPLOAD SECTION ---
st.markdown("## Upload Your Cookies File")

uploaded_cookie = st.file_uploader(
    "Choose your cookies.json file",
    type="json",
    help="Upload your cookies.json file from Canvas"
)

# --- HELPER FUNCTIONS ---
def get_expiration_date(cookie_bytes):
    try:
        cookies = json.load(cookie_bytes)
        expiration_timestamps = [
            int(cookie.get("expirationDate"))
            for cookie in cookies if "expirationDate" in cookie
        ]
        if not expiration_timestamps:
            return None
        latest_exp = max(expiration_timestamps)
        return datetime.fromtimestamp(latest_exp)
    except Exception:
        return None

# --- COOKIE VALIDATION ---
valid_cookies = False
cookie_exp = None

if uploaded_cookie:
    cookie_exp = get_expiration_date(uploaded_cookie)
    uploaded_cookie.seek(0)

    if cookie_exp:
        if cookie_exp < datetime.now():
            st.markdown(f"""
            <div class="error-box">
                <div style="font-weight: 600; margin-bottom: 0.5rem;">
                    Cookie Expired
                </div>
                <div>
                    Your cookies expired on: <strong>{cookie_exp.strftime('%Y-%m-%d %H:%M:%S')}</strong>
                </div>
                <div style="margin-top: 0.5rem;">
                    Please follow the instructions below to get a new cookies.json file.
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            st.markdown("### Get Fresh Cookies")
            st.markdown("""
            <div class="instructions-section">
                <strong>Follow these steps to generate a new cookies.json file:</strong>
            </div>
            """, unsafe_allow_html=True)
            
            steps = [
                "Open [Canvas in Chrome](https://canvas.artofdrawers.com)",
                "Install the [Cookie-Editor Extension](https://chrome.google.com/webstore/detail/cookie-editor/hlkenndedbemlkljdomclgjgkkdggpac)",
                "In Canvas, click on the Cookie-Editor Extension",
                "Keep only the **PHPSESSID** and **username** cookies",
                "Click the **Export** button at the top of the extension",
                "Visit [JSON Editor Online](https://jsoneditoronline.org/)",
                "Paste the cookies into the 'New Document 1' text box on the left",
                "Click the **save** button, then 'Save to Disk'",
                "Give it a name and click **Save**",
                "You now have your new cookies.json file!"
            ]
            
            for i, step in enumerate(steps, 1):
                st.write(f"{i}. {step}")
            
        else:
            st.markdown(f"""
            <div class="success-box">
                <div style="font-weight: 600; margin-bottom: 0.5rem;">
                    Cookies Valid
                </div>
                <div>
                    Expires on: <strong>{cookie_exp.strftime('%Y-%m-%d %H:%M:%S')}</strong>
                </div>
            </div>
            """, unsafe_allow_html=True)
            valid_cookies = True
    else:
        st.markdown("""
        <div class="error-box">
            <div style="font-weight: 600; margin-bottom: 0.5rem;">
                Invalid Cookie File
            </div>
            <div>
                Could not read the cookie file. Please check the format and try again.
            </div>
        </div>
        """, unsafe_allow_html=True)

# --- GENERATE REPORT SECTION ---
if valid_cookies:
    st.markdown("## Generate Your Report")
    
    st.markdown("""
    <div class="info-section">
        <strong>Ready to generate!</strong> Your report will include data from the last 30 days.
    </div>
    """, unsafe_allow_html=True)
    
    if st.button("Generate Weekly PDF Report", type="primary"):
        with st.status("Generating your report. This may take up to 2 minutes. Please wait...", expanded=True) as status:
            gen = WeeklyReportGenerator()
            pdf_path = gen.generate_report(days_back=30)
            status.update(label="Report complete! Your PDF is ready.", state="complete")
        
        # Download section
        st.markdown("### Download Your Report")
        
        with open(pdf_path, "rb") as f:
            st.download_button(
                label="Download PDF Report",
                data=f,
                file_name=pdf_path.name,
                mime="application/pdf",
                help="Click to download your weekly newsletter report",
                type="primary"
            )

# --- FOOTER ---
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 0.9rem; margin-top: 2rem;">
    <p>AoD Weekly Newsletter Report Generator</p>
</div>
""", unsafe_allow_html=True)