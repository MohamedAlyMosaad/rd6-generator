import streamlit as st

st.set_page_config(
    page_title="RD Generator | SOCOTEC Arabia",
    page_icon="📄",
    layout="centered"
)

st.markdown("""
<style>
    [data-testid="stSidebar"] { display: none; }
    .main-header {
        background: linear-gradient(135deg, #0072BB, #005a94);
        padding: 28px 32px;
        border-radius: 10px;
        margin-bottom: 28px;
        box-shadow: 0 4px 16px rgba(0,114,187,0.25);
    }
    .main-header h1 { color: white; margin: 0; font-size: 2rem; letter-spacing: -0.5px; }
    .main-header p  { color: rgba(255,255,255,0.82); margin: 6px 0 0 0; font-size: 0.95rem; }
    .report-card {
        border: 2px solid #e0eaf3;
        border-radius: 10px;
        padding: 28px 24px;
        text-align: center;
        background: #f8fbff;
        transition: border-color 0.2s;
        margin-bottom: 8px;
    }
    .report-card:hover { border-color: #0072BB; }
    .report-card .icon { font-size: 2.8rem; margin-bottom: 6px; }
    .report-card h3 { color: #0072BB; margin: 8px 0 4px 0; font-size: 1.2rem; }
    .report-card p  { color: #555; font-size: 0.88rem; margin: 0; }
    .badge {
        display: inline-block;
        background: #e8f4fd;
        color: #0072BB;
        border-radius: 12px;
        padding: 3px 10px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-top: 8px;
    }
    .footer { text-align: center; color: #999; font-size: 0.8rem; margin-top: 40px; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📄 RD Generator</h1>
    <p>SOCOTEC Arabia — Technical Inspection Report Generator</p>
</div>
""", unsafe_allow_html=True)

st.markdown("### Select Report Type")
st.markdown("Both generators produce a ready-to-sign Word document (.docx) using the official SOCOTEC template.")

col1, col2 = st.columns(2, gap="large")

with col1:
    st.markdown("""
    <div class="report-card">
        <div class="icon">📋</div>
        <h3>RD6 Generator</h3>
        <p>Final Completion of Works<br>Certificate Report</p>
        <div class="badge">Completion Report</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Open RD6 Generator →", use_container_width=True, key="rd6"):
        st.switch_page("pages/1_RD6_Generator.py")

with col2:
    st.markdown("""
    <div class="report-card">
        <div class="icon">🏗️</div>
        <h3>RD3 Generator</h3>
        <p>Waterproofness Final<br>Assessment Report</p>
        <div class="badge">Waterproofing Report</div>
    </div>
    """, unsafe_allow_html=True)
    if st.button("Open RD3 Generator →", use_container_width=True, key="rd3"):
        st.switch_page("pages/2_RD3_Generator.py")

st.markdown("""
<div class="footer">
    SOCOTEC Arabia – KSA &nbsp;|&nbsp; Internal Use Only
</div>
""", unsafe_allow_html=True)

