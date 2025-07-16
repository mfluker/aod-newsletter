import streamlit as st
import json
import time
from datetime import datetime
from pathlib import Path
from report import WeeklyReportGenerator

# --- HEADER SECTION ---
st.set_page_config(page_title="AoD Weekly Newsletter", layout="centered")
st.title("📬 AoD Weekly Newsletter Report")

# st.markdown("""
# Before continuing:
# 1. **Log in to Canvas** in a separate tab  
# 2. Then return here to upload your `canvas_cookies.json`  
# """)
# st.markdown("[🔗 Open Canvas Login](https://canvas.artofdrawers.com)")

st.markdown("""
Before continuing:

1. **[Log in to Canvas](https://canvas.artofdrawers.com)** in a separate tab  
2. Then return here to upload your `canvas_cookies.json`  
""")


# --- COOKIE UPLOAD ---
uploaded_cookie = st.file_uploader("Step 1: Upload your `canvas_cookies.json`", type="json")

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

# Hold result
valid_cookies = False
cookie_exp = None

if uploaded_cookie:
    cookie_exp = get_expiration_date(uploaded_cookie)
    uploaded_cookie.seek(0)  # rewind for later use

    if cookie_exp:
        st.success(f"Cookie expires on: `{cookie_exp.strftime('%Y-%m-%d %H:%M:%S')}`")
        if cookie_exp < datetime.now():
            st.error("❌ Your cookies have expired.")
            st.markdown("#### 🔁 How to get fresh cookies:")
            st.markdown("""
1. Open [Canvas in Chrome](https://canvas.artofdrawers.com)
2. Install the [Cookie-Editor Extension](https://chrome.google.com/webstore/detail/cookie-editor/hlkenndedbemlkljdomclgjgkkdggpac)
3. Click "Export" in the extension, and save it as `canvas_cookies.json`, then upload here.
""")
        else:
            valid_cookies = True
    else:
        st.error("❌ Could not read cookie file. Please check the format.")

# --- GENERATE REPORT BUTTON ---
if valid_cookies:
    if st.button("📄 Generate Weekly PDF Report"):
        with st.status("⏳ Running report (may take ~ 2 minutes)...", expanded=True) as status:
            # Save cookies to correct path
            cookie_path = Path("canvas_cookies.json")
            cookie_path.write_bytes(uploaded_cookie.read())

            st.write("🔌 Connecting to Canvas...")
            time.sleep(1)
            gen = WeeklyReportGenerator()

            st.write("📦 Fetching data from all sources...")
            # Simulated progress bar
            progress = st.progress(0)
            for i in range(1):
                time.sleep(1)
                progress.progress((i + 1) / 90.0)

            st.write("🧾 Building PDF report...")
            pdf_path = gen.generate_report(days_back=30)

            status.update(label="✅ Report complete!", state="complete")

        with open(pdf_path, "rb") as f:
            st.download_button(
                label="📥 Download PDF",
                data=f,
                file_name=pdf_path.name,
                mime="application/pdf"
            )
