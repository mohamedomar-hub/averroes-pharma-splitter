# -*- coding: utf-8 -*-
"""
Averroes Pharma File Splitter & Dashboard — Streamlit Cloud Compatible
"""
import streamlit as st
import pandas as pd
from io import BytesIO
from zipfile import ZipFile
import re
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl import load_workbook, Workbook
# Dashboard & Reporting
import matplotlib.pyplot as plt
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
import plotly.express as px
import plotly.graph_objects as go
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image
import numpy as np

# Optional opencv
try:
    import cv2
    CV2_AVAILABLE = True
except Exception:
    CV2_AVAILABLE = False

# Session state
if 'clear_counter' not in st.session_state:
    st.session_state.clear_counter = 0
if 'sidebar_open' not in st.session_state:
    st.session_state.sidebar_open = False
if 'last_action' not in st.session_state:
    st.session_state.last_action = None
if 'show_toast' not in st.session_state:
    st.session_state.show_toast = False

# Page config
st.set_page_config(
    page_title="Averroes Pharma File Splitter & Dashboard",
    page_icon="💊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Load external CSS
with open("style.css", "r", encoding="utf-8") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# Load external JS
with open("script.js", "r", encoding="utf-8") as f:
    st.markdown(f"<script>{f.read()}</script>", unsafe_allow_html=True)

# Sidebar HTML
sidebar_html = """
<div id="sidebar-custom" class="sidebar-custom">
    <h3>💊 Averroes Pharma</h3>
    <div class="small">Toolbox: File Splitter - Dashboard</div>
    <hr style="border-color: rgba(255,215,0,0.06); margin:8px 0;">
    <a href="javascript:void(0)" onclick="navigateTo('home')">🏠 &nbsp; Home</a>
    <a href="javascript:void(0)" onclick="navigateTo('split')">📂 &nbsp; Split & Merge</a>
    <a href="javascript:void(0)" onclick="navigateTo('images')">📷 &nbsp; Image to PDF</a>
    <a href="javascript:void(0)" onclick="navigateTo('dashboard')">📊 &nbsp; Auto Dashboard</a>
    <a href="javascript:void(0)" onclick="navigateTo('info-section')">ℹ️ &nbsp; Info</a>
    <hr style="border-color: rgba(255,215,0,0.06); margin:8px 0;">
    <div style="margin-top:12px; color:#e6eef8; font-size:13px;">
        <div>Made by <strong>Mohamed Abd ELGhany</strong></div>
        <div style="margin-top:8px;"><a href="javascript:void(0)" onclick="openWhatsAppJS()" style="color:#FFD700; text-decoration:none;">Contact Support</a></div>
    </div>
</div>
"""

# Render UI controls
st.markdown("""
<button id="sidebarToggle" onclick="toggleSidebarJS()">☰</button>
""" + sidebar_html + """
<div id="global-toast" class="toast" role="status" aria-live="polite"></div>
<button id="backToTop" onclick="backToTopJS()">⬆️</button>
""", unsafe_allow_html=True)

# =============== Helper Functions ===============
# (same as your original code – no change needed)
# ... [paste all your helper functions here: _safe_name, _find_col, build_pdf, etc.] ...

# =============== Page Content ===============
# Home
st.markdown("<div id='home'></div>", unsafe_allow_html=True)
col1, col2 = st.columns([3,1])
with col1:
    st.markdown("""
    <div class="hero" role="banner">
        <div style="flex:1;">
            <div class="title">💊 Averroes Pharma File Splitter & Dashboard</div>
            <div class="subtitle">Split, Merge, convert images to PDF and auto-generate dashboards — fast and secure.</div>
            <div style="margin-top:12px;">
                <a href="javascript:void(0)" class="cta" onclick="navigateTo('split')">🚀 Start Now</a>
                &nbsp;
                <a href="javascript:void(0)" class="cta secondary" onclick="navigateTo('dashboard')">📊 Try Dashboard</a>
            </div>
        </div>
        <div style="width:160px; text-align:center;">
            <!-- Logo removed for Streamlit Cloud compatibility -->
            <div style="color:#e6eef8; margin-top:6px; font-size:13px;">By <strong>Mohamed Abd ELGhany</strong></div>
        </div>
    </div>
    """, unsafe_allow_html=True)
with col2:
    st.markdown(f"""
    <div style='text-align:right; margin-right:20px;'>
        <button onclick="openWhatsAppJS()" style='background:#FFD700; color:black; border:none; padding:8px 12px; border-radius:8px; font-weight:700; cursor:pointer;'>📥 Contact (WhatsApp)</button>
        <div style='text-align:right; color:#a9c1df; margin-top:10px;'>Last updated: <strong>{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</strong></div>
    </div>
    """, unsafe_allow_html=True)

# ... [باقي أقسام Split, Merge, Images, Dashboard, Info كما هي] ...

# عند كل زر "Clear", استخدم:
# st.rerun()  ← بدل st.experimental_rerun()

# مثال:
# if st.button("🗑️ Clear Uploaded File", key="clear_split"):
#     st.session_state.clear_counter += 1
#     st.rerun()

# =============== Toast Trigger ===============
if st.session_state.show_toast:
    action_text = st.session_state.last_action or "Operation completed"
    st.markdown(f"""
    <script>
    setTimeout(() => {{
        try {{ showToastJS({action_text!r}); }} catch(e) {{ console.error(e); }}
    }}, 200);
    </script>
    """, unsafe_allow_html=True)
    st.session_state.show_toast = False

st.markdown("<div style='height:30px;'></div>", unsafe_allow_html=True)
st.markdown(f"<div style='text-align:center; color:#a9c1df; font-size:13px; margin-bottom:28px;'>© Averroes Pharma — {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>", unsafe_allow_html=True)
