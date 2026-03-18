import os
import base64
import tempfile
import subprocess
from io import BytesIO
from shutil import which

import fitz  # PyMuPDF
import gspread
import streamlit as st
from streamlit_ace import st_ace
from google.oauth2.service_account import Credentials
from pptx import Presentation
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt

# =========================
# PATHS / SETTINGS
# =========================
SOFFICE_PATH = os.environ.get("SOFFICE_PATH", "soffice")

SHEET_KEY = st.secrets["SHEET_KEY"]
WORKSHEET_NAME = st.secrets["WORKSHEET_NAME"]

FIRST_LAYOUT_NAME = "TEMPLATE_FIRST"
REST_LAYOUT_NAME = "TEMPLATE_REST"

# =========================
# GOOGLE SHEETS
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=
