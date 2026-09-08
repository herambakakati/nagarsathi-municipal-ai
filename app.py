import os
import base64
import streamlit.components.v1 as components
from pathlib import Path
from datetime import datetime
from pypdf import PdfReader
import streamlit as st
from PIL import Image
from dotenv import load_dotenv
from langchain_openai import (
    ChatOpenAI,
    OpenAIEmbeddings
)

from langchain_community.vectorstores import FAISS
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import StrOutputParser


# ============================================================
# HTML RENDER HELPER
# ============================================================

import textwrap
def render_html(html):

    st.html(
        textwrap.dedent(html).strip()
    )


# ============================================================
# 1. ENVIRONMENT
# ============================================================

load_dotenv()


# ============================================================
# 2. PAGE CONFIGURATION
# ============================================================

st.set_page_config(

    page_title="NagarSathi",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"

)

st.set_option("client.toolbarMode", "viewer")

# ============================================================
# 3. BASE PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
PDF_DIR = DATA_DIR / "pdf_file"
WORD_DIR = DATA_DIR / "Word_file"
EXCEL_DIR = DATA_DIR / "Excel_file"
URL_FILE = DATA_DIR / "url.txt"
VECTORSTORE_DIR = BASE_DIR / "vectorstore"
HERO_IMAGE = BASE_DIR / "assets" / "municipal_ai.png"


# ============================================================
# 4. PREMIUM CSS
# ============================================================

st.markdown(
    """
<style>

/* ============================================================
   GLOBAL
   ============================================================ */

* {
    box-sizing: border-box;
}

html, body, [class*="css"] {
    font-family:
        Inter,
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        sans-serif;
}

.stApp {

    background:
        radial-gradient(
            circle at 82% 4%,
            rgba(67, 67, 255, 0.15),
            transparent 28%
        ),
        radial-gradient(
            circle at 45% 45%,
            rgba(30, 55, 180, 0.07),
            transparent 38%
        ),
        #030817;

    color: #F7F8FF;

}


/* ============================================================
   REMOVE DEFAULT STREAMLIT SPACE
   ============================================================ */

header {
    background: transparent !important;
}

#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

.main .block-container,
section[data-testid="stMain"] .block-container,
div[data-testid="stMainBlockContainer"] {

    max-width: 100% !important;
    padding-top: 0 !important;
    padding-left: 0 !important;
    padding-right: 0 !important;
    padding-bottom: 2rem !important;
    margin-top: 0px !important;
}

/* ============================================================
   SIDEBAR — MATCH SECOND REFERENCE IMAGE
   ============================================================ */

section[data-testid="stSidebar"] {
    width: 340px !important;
    min-width: 340px !important;
    max-width: 340px !important;

    background:
        linear-gradient(
            180deg,
            #071024 0%,
            #050B1B 100%
        ) !important;

    border-right:
        1px solid
        rgba(89, 112, 180, 0.22) !important;
}


section[data-testid="stSidebar"] > div {
    padding-top: 0 !important;
    padding-left: 34px !important;
    padding-right: 34px !important;
    padding-bottom: 18px !important;
}


/* ============================================================
   BRAND
   ============================================================ */
.sidebar-brand {
    padding: 0 0 10px 0 !important;
    margin: 0 !important;
    text-align: left !important;
}


/* ============================================================
   BRAND ICON + NAME — SAME LINE
   ============================================================ */

.sidebar-brand-header {
    display: flex !important;
    align-items: center !important;
    gap: 10px !important;
    width: 100% !important;
    margin: 0 !important;
    padding: 0 !important;
}


.sidebar-brand-header .sidebar-brand-icon {
    width: 48px !important;
    height: 48px !important;
    flex-shrink: 0 !important;
    margin: 0 !important;

    display: flex !important;
    align-items: center !important;
    justify-content: center !important;

    font-size: 30px !important;
}


/* Keep NagarSathi beside the icon */

.sidebar-brand-header .sidebar-brand-title {
    margin: 0 !important;
    padding: 0 !important;

    font-size: 25px !important;
    line-height: 1 !important;

    white-space: nowrap !important;
}


/* LOGO */

.sidebar-brand-icon {
    width: 58px !important;
    height: 58px !important;

    display: flex !important;
    align-items: center !important;
    justify-content: center !important;

    margin: 0 0 8px 0 !important;

    background:
        radial-gradient(
            circle,
            rgba(109, 76, 255, 0.42),
            rgba(74, 54, 180, 0.10) 68%,
            transparent 72%
        ) !important;

    border-radius: 50% !important;

    font-size: 34px !important;

    filter:
        drop-shadow(
            0 0 10px
            rgba(104, 86, 255, 0.60)
        ) !important;
}


/* BRAND NAME */

.sidebar-brand-title {
    margin: 0 !important;
    padding: 0 !important;

    color: #F4F6FF !important;

    font-size: 25px !important;
    line-height: 1.1 !important;

    font-weight: 800 !important;
    letter-spacing: -0.7px !important;
}


/* SATHI */

.sidebar-brand-title .gradient {
    background:
        linear-gradient(
            90deg,
            #6384FF,
            #B15DFF
        ) !important;

    -webkit-background-clip: text !important;
    background-clip: text !important;

    -webkit-text-fill-color: transparent !important;
}


/* SUBTITLE */

.sidebar-brand-subtitle {
    margin: 7px 0 0 0 !important;
    padding: 0 !important;

    color: #8E9AB7 !important;

    font-size: 12px !important;
    line-height: 1.45 !important;

    font-weight: 400 !important;
}


/* ============================================================
   HOME — ACTIVE BUTTON
   ============================================================ */

section[data-testid="stSidebar"] .st-key-home_button {
    width: 100% !important;
    margin: 0 0 8px 0 !important;
    padding: 0 !important;
}


/* Actual Streamlit Home button */

section[data-testid="stSidebar"]
.st-key-home_button .stButton > button {

    width: 100% !important;
    height: 44px !important;
    min-height: 44px !important;
    max-height: 44px !important;
    margin: 0 !important;
    padding: 0 16px !important;
    border-radius: 9px !important;

    background:
        linear-gradient(
            90deg,
            rgba(99, 65, 237, 0.68),
            rgba(55, 82, 190, 0.45)
        ) !important;

    border:
        1px solid
        rgba(113, 86, 255, 0.78) !important;

    color: #FFFFFF !important;

    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;

    text-align: left !important;

    font-size: 14px !important;
    font-weight: 650 !important;
    line-height: 1 !important;

    box-shadow:
        0 0 18px
        rgba(90, 66, 255, 0.22) !important;

    transform: none !important;
}


/* Home text */

section[data-testid="stSidebar"]
.st-key-home_button .stButton > button p {

    margin: 0 !important;
    padding: 0 !important;
    color: #FFFFFF !important;
    font-size: 14px !important;
    font-weight: 650 !important;
    line-height: 1 !important;
}


/* Home hover */

section[data-testid="stSidebar"]
.st-key-home_button .stButton > button:hover {

    background:
        linear-gradient(
            90deg,
            rgba(99, 65, 237, 0.72),
            rgba(55, 82, 190, 0.48)
        ) !important;

    border-color:
        rgba(130, 105, 255, 0.85) !important;

    box-shadow:
        0 0 20px
        rgba(90, 66, 255, 0.28) !important;

    transform: none !important;
}


/* ============================================================
   EXPLORE
   ============================================================ */

.nav-section {
    margin: 0 0 8px 10px !important;
    padding: 0 !important;
    color: #68738F !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    letter-spacing: 1.4px !important;
    line-height: 1.2 !important;
}


/* ============================================================
   ALL SIDEBAR BUTTONS
   ============================================================ */

section[data-testid="stSidebar"] .stButton {
    margin: 0 !important;
    padding: 0 !important;
}


section[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    height: 36px !important;
    min-height: 36px !important;
    max-height: 36px !important;
    margin: 0 0 1px 0 !important;
    padding: 0 10px !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    border-radius: 9px !important;
    color: #E5E9F5 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    text-align: left !important;
    font-size: 14px !important;
    font-weight: 550 !important;
    line-height: 1 !important;
    box-shadow: none !important;
    transform: none !important;
}


/* BUTTON TEXT */

section[data-testid="stSidebar"]
.stButton > button p {
    margin: 0 !important;
    padding: 0 !important;
    font-size: 14px !important;
    line-height: 1 !important;
    font-weight: 550 !important;
}


/* ============================================================
   NAVIGATION SPACING
   ============================================================ */

.nav-section + div {
    margin-top: 0 !important;
}


.nav-section ~ .stButton {
    margin-bottom: 2px !important;
}


/* ============================================================
   SYSTEM STATUS
   ============================================================ */

.status-panel {

    margin-top: 18px !important;

    padding: 14px !important;

    border-radius: 11px !important;

    background:
        linear-gradient(
            145deg,
            rgba(15, 29, 59, 0.94),
            rgba(9, 17, 37, 0.95)
        ) !important;

    border:
        1px solid
        rgba(78, 104, 170, 0.25) !important;

    box-shadow:
        0 8px 20px
        rgba(0,0,0,0.10) !important;
}


.status-heading {
    margin: 0 0 8px 0 !important;
    color: #FFFFFF !important;
    font-size: 12px !important;
    font-weight: 700 !important;

    line-height: 1.2 !important;
}


.status-online {
    margin: 0 0 9px 0 !important;
    color: #2ED995 !important;
    font-size: 11px !important;
    line-height: 1.2 !important;
}


.status-line {

    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    padding: 4px 0 !important;
    margin: 0 !important;
    color: #8B97B3 !important;
    font-size: 11px !important;
    line-height: 1.2 !important;
}


.status-number {
    color: #F4F6FF !important;
    font-weight: 700 !important;
}


.ready-badge {
    padding: 3px 7px !important;

    border-radius: 20px !important;

    background:
        rgba(37, 209, 139, 0.17) !important;

    color: #36E19A !important;

    font-size: 9px !important;
    font-weight: 700 !important;
}


/* ============================================================
   REFRESH
   ============================================================ */

section[data-testid="stSidebar"] .status-panel + div {
    margin-top: 6px !important;
    margin-bottom: 0 !important;
}


section[data-testid="stSidebar"] .status-panel + div .stButton > button {
    height: 36px !important;
    min-height: 36px !important;
    line-height: 36px !important;
}


/* ============================================================
   TRUSTED & SECURE
   ============================================================ */

.trust-panel {

    margin-top: 8px !important;

    padding: 13px !important;

    min-height: 82px !important;

    box-sizing: border-box !important;

    border-radius: 10px !important;

    background:
        linear-gradient(
            135deg,
            rgba(82, 38, 160, 0.72),
            rgba(56, 28, 120, 0.48)
        ) !important;

    border:
        1px solid
        rgba(155, 100, 255, 0.28) !important;

    box-shadow:
        0 8px 22px
        rgba(66, 30, 150, 0.15) !important;
}


.trust-title {

    margin: 0 0 4px 0 !important;

    color: #FFFFFF !important;

    font-size: 11px !important;

    font-weight: 700 !important;

    line-height: 1.2 !important;
}


.trust-text {
    margin: 0 !important;
    color: #C5CBE0 !important;
    font-size: 11px !important;
    line-height: 1.4 !important;
}


/* ============================================================
   FOOTER
   ============================================================ */

section[data-testid="stSidebar"] .stCaption {

    margin-top: 20px !important;
    margin-bottom: 0 !important;

    color: #8E9AB7 !important;

    font-size: 10px !important;

    line-height: 1.45 !important;
}


/* ============================================================
   REMOVE STREAMLIT EXTRA GAPS
   ============================================================ */

section[data-testid="stSidebar"]
[data-testid="stVerticalBlock"] {
    gap: 0 !important;
}


section[data-testid="stSidebar"]
[data-testid="stVerticalBlockBorderWrapper"] {
    gap: 0 !important;

    margin: 0 !important;
    padding: 0 !important;
}


section[data-testid="stSidebar"]
[data-testid="element-container"] {
    margin: 0 !important;
    padding: 0 !important;
}


section[data-testid="stSidebar"]
[data-testid="stHorizontalBlock"] {
    gap: 0 !important;
}


/* ============================================================
   DEVELOPER FOOTER
   ============================================================ */

.sidebar-footer {
    margin-top: 18px !important;
    padding: 0 2px 8px 2px !important;

    text-align: left !important;
}


.sidebar-footer-copy {
    color: #8E9AB7 !important;

    font-size: 10px !important;
    line-height: 1.4 !important;
}


.sidebar-footer-project {
    margin-top: 3px !important;

    color: #8E9AB7 !important;

    font-size: 10px !important;
    line-height: 1.4 !important;
}


.sidebar-footer-developer {
    margin-top: 10px !important;

    color: #A9B2C8 !important;

    font-size: 10px !important;
    line-height: 1.4 !important;
}


.sidebar-footer-developer strong {
    color: #C9B8FF !important;
    font-weight: 600 !important;
}

/* ============================================================
   PREMIUM HERO — ORIGINAL IMAGE + EDGE BLEND
   ============================================================ */

.hero-wrapper {
    position: relative;
    width: calc(100% - 50px);
    height: 430px;
    min-height: 430px;
    margin-top: 18px;
    margin-left: 25px;
    margin-right: 25px;
    padding: 0;
    overflow: hidden;
    border-radius: 18px;
    background:
        radial-gradient(
            ellipse 620px 430px at 76% 52%,
            rgba(25, 53, 170, 0.14),
            transparent 72%
        ),
        linear-gradient(
            110deg,
            #071126 0%,
            #071126 52%,
            #07102A 100%
        );

    border: 1px solid rgba(78, 104, 235, 0.78);

    box-shadow:
            0 0 30px rgba(54, 66, 190, 0.12),
            inset 0 1px 0 rgba(255, 255, 255, 0.035);

    isolation: isolate;
}


/* REMOVE STREAMLIT TOP GAP ABOVE HERO */

div[data-testid="stMainBlockContainer"] > div:first-child {
    margin-top: 0 !important;
    padding-top: 0 !important;
}

div[data-testid="stMainBlockContainer"]
div[data-testid="stElementContainer"]:has(.hero-wrapper) {
    margin-top: 0 !important;
    padding-top: 0 !important;
}


/* ============================================================
   HERO TEXT — EXACT REFERENCE STYLE
   ============================================================ */

.hero-content {
    position: absolute;
    z-index: 20;
    left: 62px;
    top: 58px;
    width: 44%;
    max-width: 650px;
    padding: 0;
    margin: 0;
}


/* ============================================================
   GREETING
   ============================================================ */

.hero-greeting {
    margin: 0 0 32px 0;
    color: #B58BFF;
    font-family:
        Inter,
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        sans-serif;

    font-size: 27px;
    line-height: 1.1;
    font-weight: 700;
    letter-spacing: -0.6px;
}


/* ============================================================
   MAIN HEADING
   ============================================================ */

.hero-title {
    margin: 0;

    font-family:
        Inter,
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        sans-serif;

    font-size: 62px;
    line-height: 1.00;
    font-weight: 800;
    letter-spacing: -3px;
    color: #F7F9FF;
    white-space: nowrap;
}


/* ============================================================
   GRADIENT HEADING
   ============================================================ */

.hero-title .blue {
    display: inline-block;

    background:
        linear-gradient(
            90deg,
            #5485FF 0%,
            #7869FF 52%,
            #B05CFF 100%
        );

    -webkit-background-clip: text;
    background-clip: text;

    -webkit-text-fill-color: transparent;
    color: transparent;
}

/* ============================================================
   HERO DESCRIPTION 
   ============================================================ */

.hero-description {
    margin-top: 28px;

    width: 100%;
    max-width: 600px;

    color: #A9B6D5;

    font-family:
        Inter,
        -apple-system,
        BlinkMacSystemFont,
        "Segoe UI",
        sans-serif;

    font-size: 18px;
    line-height: 1.65;
    font-weight: 400;
    letter-spacing: 0;
    white-space: normal;
}


/* ============================================================
   HERO ARTWORK
   IMPORTANT:
   Use the ORIGINAL municipal_ai.png.
   Do NOT remove its internal background.
   Only blend the OUTER EDGE of the image.
   ============================================================ */

.hero-visual {
    position: absolute;

    z-index: 5;

    top: 0;
    right: 0;

    width: 61%;
    height: 100%;

    display: flex;

    align-items: center;
    justify-content: center;

    pointer-events: none;

    overflow: hidden;
}

/* ============================================================
   ORIGINAL IMAGE — EDGE ONLY BLEND
   ============================================================ */

.hero-visual img {
    display: block;

    width: 100%;
    max-width: 900px;
    max-height: 440px;

    height: auto;

    object-fit: contain;
    object-position: center center;

    margin: 0;
    position: relative;
    z-index: 1;

    padding: 0;

    background: transparent !important;

    border: 0 !important;
    outline: none !important;

    box-shadow: none !important;

    opacity: 1;

    filter: none !important;

    /*
       Only the OUTER LEFT and RIGHT edges fade.
       The centre of the original PNG remains untouched.
    */

    -webkit-mask-image:
        linear-gradient(
            90deg,
            transparent 0%,
            rgba(0,0,0,0.45) 3%,
            #000 9%,
            #000 86%,
            rgba(0,0,0,0.92) 90%,
            rgba(0,0,0,0.62) 94%,
            rgba(0,0,0,0.25) 98%,
            transparent 100%
        );

    mask-image:
        linear-gradient(
            90deg,
            transparent 0%,
            rgba(0,0,0,0.45) 3%,
            #000 9%,
            #000 86%,
            rgba(0,0,0,0.92) 90%,
            rgba(0,0,0,0.62) 94%,
            rgba(0,0,0,0.25) 98%,
            transparent 100%
        );

    -webkit-mask-repeat: no-repeat;
    mask-repeat: no-repeat;

    -webkit-mask-position: center;
    mask-position: center;

    -webkit-mask-size: 100% 100%;
    mask-size: 100% 100%;
}

/* ============================================================
   IMAGE EDGE BLEND
   ============================================================ */

.hero-visual::before {
    content: "";

    position: absolute;

    inset: 0;

    z-index: 11;

    pointer-events: none;

    background:
        linear-gradient(
            90deg,

            /* LEFT EDGE */
            #071126 0%,
            rgba(7,17,38,0.92) 6%,
            rgba(7,17,38,0.55) 13%,
            rgba(7,17,38,0.16) 20%,
            transparent 27%,

            /* ORIGINAL IMAGE AREA */
            transparent 70%,

            /* RIGHT EDGE — SOFT BLEND */
            rgba(7,16,42,0.08) 76%,
            rgba(7,16,42,0.20) 82%,
            rgba(7,16,42,0.38) 88%,
            rgba(7,16,42,0.58) 93%,
            rgba(7,16,42,0.76) 97%,
            #07102A 100%
        );
}

/* ============================================================
   RESPONSIVE
   ============================================================ */

@media (max-width: 1200px) {

    .hero-content {
        left: 48px;
        top: 68px;

        width: 46%;
    }

    .hero-title {
        font-size: 46px;
    }

    .hero-description {
        font-size: 14px;
        max-width: 500px;
    }

    .hero-visual {
        width: 61%;
    }
}


@media (max-width: 900px) {

    .hero-wrapper {
        width: calc(100% - 36px);
        margin-left: 18px;
        margin-right: 18px;

        height: 540px;
        min-height: 540px;
        margin-top: 10px;
    }

    .hero-content {
        position: absolute;

        left: 32px;
        right: 32px;
        top: 34px;

        width: auto;
        max-width: none;

        z-index: 20;
    }

    .hero-greeting {
        font-size: 16px;
        margin-bottom: 12px;
    }

    .hero-title {
        font-size: 38px;
        line-height: 1.05;

        white-space: normal;

        letter-spacing: -1.7px;
    }

    .hero-description {
        margin-top: 16px;

        max-width: 100%;

        font-size: 13px;
        line-height: 1.6;
    }

    .hero-visual {
        top: auto;
        bottom: -10px;

        left: 0;
        right: 0;

        width: 100%;
        height: 310px;

        overflow: hidden;
    }

    .hero-visual img {
        width: 100%;
        max-width: 620px;
        max-height: 320px;
    }
}


@media (max-width: 600px) {

    .hero-wrapper {
        width: calc(100% - 28px);
        margin-left: 14px;
        margin-right: 14px;
        
        height: 540px;
        min-height: 540px;

        border-radius: 18px;
    }

    .hero-content {
        left: 24px;
        right: 24px;
        top: 28px;
    }

    .hero-title {
        font-size: 32px;
    }

    .hero-description {
        font-size: 12px;
    }

    .hero-visual {
        height: 280px;
    }
}

/* ============================================================
   QUESTION / SEARCH PANEL
   ============================================================ */

.st-key-question_panel {
    width: calc(100% - 50px) !important;
    margin-top: 12px !important;
    margin-left: 25px !important;
    margin-right: 25px !important;
    padding: 19px 30px 18px 30px !important;
    min-height: 180px !important;
    height: auto !important;
    overflow: visible !important;
    padding-bottom: 18px !important;

    border-radius: 18px !important;

    background:
        linear-gradient(
            145deg,
            #091833 0%,
            #07142D 55%,
            #061126 100%
        ) !important;

    border: 1px solid rgba(78, 104, 235, 0.78); 

    box-shadow:
            0 0 30px rgba(54, 66, 190, 0.12),
            inset 0 1px 0 rgba(255, 255, 255, 0.035);
}


/* ============================================================
   QUESTION LABEL
   ============================================================ */

.st-key-question_panel .question-label {

    color: #B7A5E8;
    font-size: 17px;
    font-weight: 650;
    line-height: 1.2;
    margin: 0 0 8px 0;
}


/* ============================================================
   COLUMNS
   ============================================================ */

.st-key-question_panel
div[data-testid="column"] {

    padding-left: 0 !important;
    padding-right: 0 !important;
}



/* Keep every row inside the panel */

.st-key-question_panel
.stHorizontalBlock {

    width: 100% !important;
    margin-left: 0 !important;
    margin-right: 0 !important;
}


/* Remove Streamlit column padding */

.st-key-question_panel
div[data-testid="column"] {

    padding-left: 0 !important;

    padding-right: 0 !important;
}


/* ============================================================
   SEARCH INPUT CONTAINER
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"] {

    position: relative;
    width: 100% !important;
    margin: 0 !important;
}


/* ============================================================
   SEARCH ICON
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"]::before {

    content: "⌕";
    position: absolute;
    left: 15px;
    top: 50%;
    transform: translateY(-50%);
    z-index: 10;
    color: #8C70E8;
    font-size: 22px;
    line-height: 1;
    pointer-events: none;
}


/* ============================================================
   STREAMLIT INPUT OUTER WRAPPER
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"] > div {

    min-height: 48px !important;
    height: 48px !important;
    width: 100% !important;
}

/* ============================================================
SEARCH INPUT
============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"]
div[data-baseweb="input"] {

    height: 52px !important;
    min-height: 52px !important;
    width: 100% !important;
    
    background:
        linear-gradient(
            180deg,
            #111C35 0%,
            #0D1830 100%
        ) !important;

    border: 1px solid
        rgba(76, 101, 170, 0.62) !important;

    border-radius: 11px !important;

    box-shadow:
        inset 0 1px 0
        rgba(255,255,255,0.025) !important;
}


/* Actual text field */

.st-key-question_panel
div[data-testid="stTextInput"]
div[data-baseweb="input"]
input {

    height: 50px !important;
    min-height: 50px !important;
    width: 100% !important;
    padding:
        0 16px 0 47px !important;

    background: transparent !important;
    color: #F5F7FF !important;
    border: 0 !important;
    outline: none !important;
    font-size: 18px !important;
    font-weight: 500 !important;
    line-height: 1.25 !important;
    letter-spacing: 0 !important;
}


/* Placeholder */

.st-key-question_panel
div[data-testid="stTextInput"]
div[data-baseweb="input"]
input::placeholder {

    color: #E1E4EE !important;
    opacity: 1 !important;
}


/* ============================================================
   INPUT FIELD
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"]
div[data-baseweb="input"] input {

    height: 50px !important;
    min-height: 50px !important;
    width: 100% !important;
    padding:
        0 16px 0 47px !important;

    background: transparent !important;
    color: #F5F7FF !important;
    border: 0 !important;
    outline: none !important;
    font-size: 18px !important;
    font-weight: 500 !important;
    line-height: 1.25 !important;
    letter-spacing: 0 !important;
}


/* ============================================================
   PLACEHOLDER
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"]
div[data-baseweb="input"]
input::placeholder {

    color: rgba(225, 228, 238, 0.55) !important;
    opacity: 1 !important;
}


/* ============================================================
   INPUT FOCUS
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"]
div[data-baseweb="input"]:focus-within {

    border-color:
        #5D66C9 !important;

    box-shadow:
        0 0 0 1px
        rgba(93, 102, 201, 0.20) !important;
}


/* ============================================================
   REMOVE STREAMLIT LABEL
   ============================================================ */

.st-key-question_panel
div[data-testid="stTextInput"] label {

    display: none !important;
}


/* ============================================================
   SEARCH BUTTON
   ============================================================ */

.st-key-search_button {

    width: 52px !important;
    min-width: 52px !important;
    max-width: 52px !important;
    height: 52px !important;
    min-height: 52px !important;
    margin: 0 !important;
    padding: 0 !important;
}


/* ============================================================
   BUTTON CONTAINER
   ============================================================ */

.st-key-search_button
.stButton {

    width: 52px !important;
    min-width: 52px !important;
    max-width: 52px !important;
    height: 52px !important;
    margin: 0 !important;
    padding: 0 !important;
}


/* ============================================================
   SEARCH BUTTON
   ============================================================ */

.st-key-question_panel
.st-key-search_button
div[data-testid="stButton"]
button {

    width: 52px !important;

    min-width: 52px !important;

    max-width: 52px !important;

    height: 52px !important;

    min-height: 52px !important;

    max-height: 52px !important;

    margin: 0 !important;

    padding: 0 !important;

    border-radius: 11px !important;

    border:
        1px solid
        rgba(115, 105, 255, 0.85) !important;

    background:
        linear-gradient(
            135deg,
            #4C5BFF 0%,
            #6230D9 100%
        ) !important;

    color: #FFFFFF !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    line-height: 52px !important;
    display: flex !important;
    align-items: center !important;
    justify-content: center !important;

    box-shadow:
        0 5px 18px
        rgba(72, 65, 235, 0.28) !important;
}


/* ============================================================
   SEND BUTTON HOVER
   ============================================================ */

.st-key-search_button
.stButton > button:hover {
    width: 52px !important;
    min-width: 52px !important;
    max-width: 52px !important;
    height: 52px !important;
    min-height: 52px !important;
    border-color:
        #9A91FF !important;

    background:
        linear-gradient(
            135deg,
            #5968FF 0%,
            #713BE5 100%
        ) !important;

    box-shadow:
        0 6px 24px
        rgba(83, 76, 245, 0.42) !important;

    transform: none !important;
}


/* ============================================================
   SUGGESTION TITLE
   ============================================================ */

.st-key-question_panel .suggestion-title {

    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    width: 100% !important;
    height: 36px !important;
    min-height: 36px !important;
    margin: 0 !important;
    padding: 0 !important;
    color: #8792B0 !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    line-height: 1 !important;
    white-space: nowrap !important;
}


/* ============================================================
   SPARKLE ICON
   ============================================================ */

.st-key-question_panel
.suggestion-title .sparkle {

    display: inline-flex !important;
    align-items: center !important;
    justify-content: center !important;
    flex-shrink: 0 !important;
    width: 28px !important;
    height: 28px !important;
    margin-right: 6px !important;
    font-size: 13px !important;
    border-radius: 50% !important;
    background:
        rgba(24, 37, 68, 0.75) !important;

    color: #D9D7FF !important;

    font-size: 13px !important;
    line-height: 1 !important;
}


/* ============================================================
   SUGGESTION CONTAINERS
   ============================================================ */

.st-key-question_panel
[class*="st-key-suggestion_"] {

    width: 100% !important;
    height: 34px !important;
    min-height: 34px !important;
    margin: 4px 0 0 0 !important;
    padding: 0 !important;
}


/* ============================================================
   SUGGESTION BUTTON
   ============================================================ */

.st-key-question_panel
[class*="st-key-suggestion_"]
.stButton > button {

    width: 100% !important;

    height: 36px !important;
    min-height: 36px !important;
    max-height: 36px !important;
    margin: 0 !important;
    padding: 0 8px !important;
    border-radius: 9px !important;
    border:
        1px solid
        rgba(63, 91, 157, 0.55) !important;

    background:
        linear-gradient(
            180deg,
            #132448 0%,
            #0D1B38 100%
        ) !important;

    color: #D4DAEA !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    line-height: 36px !important;
    white-space: nowrap !important;
    overflow: hidden !important;
    text-overflow: ellipsis !important;
    box-shadow: none !important;
}


/* ============================================================
   SUGGESTION BUTTON TEXT
   ============================================================ */

.st-key-question_panel
[class*="st-key-suggestion_"]
.stButton > button p {

    margin: 0 !important;
    padding: 0 !important;
    font-size: 13px !important;
    font-weight: 500 !important;
    line-height: 1 !important;
    white-space: nowrap !important;
}


/* ============================================================
   SUGGESTION HOVER
   ============================================================ */

.st-key-question_panel
[class*="st-key-suggestion_"]
.stButton > button:hover {

    border-color:
        rgba(111, 119, 235, 0.70) !important;

    background:
        linear-gradient(
            180deg,
            rgba(27, 44, 82, 0.98),
            rgba(14, 28, 56, 0.98)
        ) !important;

    color: #FFFFFF !important;
}


/* ============================================================
   REMOVE EXTRA STREAMLIT SPACING
   ============================================================ */

.st-key-question_panel
div[data-testid="stVerticalBlock"] {

    gap: 0 !important;
}

.answer-card {
    position: relative;
    width: calc(100% - 50px);
    margin: 12px 25px 0 25px;
    padding: 24px 30px 23px 30px;
    
    border-radius: 18px;

    background:
        radial-gradient(
            circle at 88% 42%,
            rgba(74, 67, 255, 0.12),
            transparent 24%
        ),
        radial-gradient(
            circle at 100% 0%,
            rgba(72, 72, 255, 0.07),
            transparent 32%
        ),
        linear-gradient(
            145deg,
            #071228 0%,
            #050D1F 58%,
            #07102A 100%
        );

    border: 1px solid rgba(78, 104, 235, 0.78);

    box-shadow:
        0 0 30px rgba(54, 66, 190, 0.12),
        inset 0 1px 0 rgba(255, 255, 255, 0.035);

    overflow: hidden;
}

/* ============================================================
   INSUFFICIENT INFORMATION CARD
   ============================================================ */

.answer-card.no-info-card {

    min-height: 250px;

    padding-top: 38px;
    padding-bottom: 38px;

    display: flex;
    align-items: center;
}


.answer-card.no-info-card .answer-top {

    min-height: 170px;

    align-items: center;
}


.answer-card.no-info-card .answer-content {

    padding-top: 0;
}


.answer-card.no-info-card .answer-title {

    margin-bottom: 18px;
}


.answer-card.no-info-card .answer-text {

    max-width: 900px;

    font-size: 16px;

    line-height: 1.7;

    color: #E4E8F3;
}







/* ============================================================
   ANSWER TOP AREA
   ============================================================ */

.answer-top {
    position: relative;
    z-index: 5;

    display: flex;
    align-items: flex-start;

    width: 100%;
}


/* ============================================================
   ANSWER ICON
   ============================================================ */

.answer-icon {
    flex-shrink: 0;

    width: 60px;
    height: 60px;

    margin-right: 20px;

    display: flex;
    align-items: center;
    justify-content: center;

    border-radius: 50%;

    background:
        linear-gradient(
            145deg,
            #516BFF 0%,
            #713BE4 100%
        );

    border: 1px solid rgba(143, 142, 255, 0.65);

    box-shadow:
        0 8px 25px rgba(67, 61, 220, 0.30),
        inset 0 1px 0 rgba(255,255,255,0.18);

    color: #FFFFFF;

    font-size: 27px;
    line-height: 1;
}


/* ============================================================
   ANSWER CONTENT
   ============================================================ */

.answer-content {
    flex: 1;
    min-width: 0;

    padding-top: 2px;
}


/* ============================================================
   ANSWER TITLE
   ============================================================ */

.answer-title {
    margin: 0 0 14px 0;
    color: #FFFFFF;
    font-size: 20px;
    line-height: 1.25;

    font-weight: 800;

    letter-spacing: -0.25px;
}


/* ============================================================
   ANSWER TEXT
   ============================================================ */

.answer-text {
    max-width: 820px;

    color: #E4E8F3;

    font-size: 15px;
    line-height: 1.65;

    font-weight: 400;

    white-space: normal;
}


/* ============================================================
   DECORATIVE RIGHT ILLUSTRATION
   ============================================================ */

.answer-art {
    position: absolute;

    right: 28px;
    top: 26px;

    width: 190px;
    height: 145px;

    z-index: 1;

    pointer-events: none;

    opacity: 0.95;
}


/* Large glowing circular document/search area */

.answer-art-circle {
    position: absolute;

    right: 12px;
    top: 12px;

    width: 108px;
    height: 108px;

    border-radius: 50%;

    background:
        radial-gradient(
            circle,
            rgba(76, 111, 255, 0.35) 0%,
            rgba(69, 75, 220, 0.20) 42%,
            rgba(73, 65, 210, 0.08) 66%,
            transparent 72%
        );

    border:
        1px solid rgba(101, 104, 255, 0.48);

    box-shadow:
        0 0 28px rgba(70, 79, 255, 0.25);
}


/* Inner rings */

.answer-art-circle::before {
    content: "";

    position: absolute;

    inset: 12px;

    border-radius: 50%;

    border:
        2px solid rgba(93, 128, 255, 0.55);

    box-shadow:
        0 0 14px rgba(77, 112, 255, 0.20);
}

.answer-art-circle::after {
    content: "";

    position: absolute;

    inset: 24px;

    border-radius: 50%;

    border:
        1px solid rgba(133, 111, 255, 0.55);
}


/* Magnifying glass */

.answer-art-search {
    position: absolute;

    right: 31px;
    top: 37px;

    width: 68px;
    height: 68px;

    border-radius: 50%;

    border:
        8px solid #5C86FF;

    transform: rotate(-45deg);

    box-shadow:
        0 0 15px rgba(75, 111, 255, 0.55);
}

.answer-art-search::after {
    content: "";

    position: absolute;

    width: 38px;
    height: 8px;

    right: -30px;
    bottom: -18px;

    border-radius: 10px;

    background:
        linear-gradient(
            90deg,
            #4D67D9,
            #7149D8
        );

    box-shadow:
        0 0 12px rgba(77, 90, 230, 0.45);
}


/* Small decorative stars */

.answer-art-star {
    position: absolute;

    color: #776CFF;

    font-size: 18px;

    text-shadow:
        0 0 12px rgba(100, 92, 255, 0.8);
}

.answer-art-star.one {
    right: 150px;
    top: 8px;
}

.answer-art-star.two {
    right: 6px;
    top: 2px;

    font-size: 13px;
}

.answer-art-star.three {
    right: 158px;
    bottom: 14px;

    font-size: 12px;
}

.answer-sources {
    position: relative;
    z-index: 10;
    width: 100%;
    margin-top: 60px;
    padding: 0;
}


/* ============================================================
   REFERENCE DOCUMENTS PANEL
   SAME WIDTH AS HERO / SEARCH / ANSWER CARDS
   ============================================================ */

.reference-documents-panel {

    width:
        calc(100% - 50px) !important;

    margin:
        10px 25px 0 25px !important;

    padding:
        0 !important;

    box-sizing:
        border-box !important;
}


/* ============================================================
   COLLAPSIBLE CONTAINER
   ============================================================ */

.reference-documents-panel details {

    width:
        100% !important;

    margin:
        0 !important;

    padding:
        0 !important;

    border:
        1px solid
        rgba(62, 91, 157, 0.34);

    border-radius:
        11px;

    background:
        linear-gradient(
            180deg,
            rgba(11, 27, 56, 0.92),
            rgba(8, 21, 46, 0.92)
        );

    overflow:
        hidden;
}


/* ============================================================
   HEADER
   COLLAPSED BY DEFAULT
   ============================================================ */

.reference-documents-panel summary {

    width:
        100%;

    min-height:
        46px;

    padding:
        0 18px;

    box-sizing:
        border-box;

    display:
        flex;

    align-items:
        center;

    cursor:
        pointer;

    list-style:
        none;

    color:
        #E6EAF4;

    font-size:
        14px;

    font-weight:
        650;

    user-select:
        none;
}


/* Remove default marker */

.reference-documents-panel summary::-webkit-details-marker {

    display:
        none;
}


/* ============================================================
   FOLDER ICON
   ============================================================ */

.reference-folder-icon {

    flex-shrink:
        0;

    margin-right:
        9px;

    font-size:
        17px;

    line-height:
        1;
}


/* ============================================================
   TITLE
   ============================================================ */

.reference-panel-title {

    color:
        #E6EAF4;

    white-space:
        nowrap;
}


/* ============================================================
   COUNT
   ============================================================ */

.reference-panel-count {

    margin-left:
        6px;

    color:
        #8996B5;

    font-size:
        13px;

    font-weight:
        500;
}


/* ============================================================
   ARROW
   ============================================================ */

.reference-panel-arrow {

    margin-left:
        auto;

    color:
        #C4CCDF;

    font-size:
        26px;

    font-weight:
        300;

    line-height:
        1;

    transition:
        transform 0.18s ease;
}


/* Rotate when opened */

.reference-documents-panel details[open]
.reference-panel-arrow {

    transform:
        rotate(90deg);
}


/* ============================================================
   DOCUMENT LIST
   ============================================================ */

.reference-document-list {

    width:
        100%;

    padding:
        3px 18px 10px 18px;

    box-sizing:
        border-box;

    border-top:
        1px solid
        rgba(65, 91, 155, 0.22);
}


/* ============================================================
   DOCUMENT ROW
   ============================================================ */

.reference-document-row {

    width:
        100%;

    min-height:
        62px;

    display:
        flex;

    align-items:
        center;

    box-sizing:
        border-box;

    border-bottom:
        1px solid
        rgba(65, 91, 155, 0.14);
}


.reference-document-row:last-child {

    border-bottom:
        none;
}


/* ============================================================
   MAIN DOCUMENT AREA
   ============================================================ */

.reference-document-main {

    width:
        100%;

    display:
        flex;

    align-items:
        center;

    min-width:
        0;
}


/* ============================================================
   DOCUMENT ICON
   ============================================================ */

.reference-document-icon {

    width:
        30px;

    flex:
        0 0 30px;

    margin-right:
        1px;

    color:
        #D8DCE8;

    font-size:
        14px;

    text-align:
        center;
}


/* ============================================================
   DOCUMENT CONTENT
   ============================================================ */

.reference-document-content {

    min-width:
        0;

    flex:
        1;
}


/* ============================================================
   REFERENCE NUMBER
   ============================================================ */

.reference-document-title {

    color:
        #E5E9F4;

    font-size:
        11px;

    font-weight:
        650;

    line-height:
        1.2;

    margin-bottom:
        3px;
}


/* ============================================================
   DOCUMENT REFERENCE LINE
   ============================================================ */

.reference-document-reference {

    display:
        flex;

    align-items:
        center;

    min-width:
        0;

    white-space:
        nowrap;

    overflow:
        hidden;

    color:
        #AEB8CE;

    font-size:
        11px;

    line-height:
        1.35;
}


/* ============================================================
   DOCUMENT NAME
   ============================================================ */

.reference-document-name {

    min-width:
        0;

    max-width:
        none;

    overflow:
        hidden;

    text-overflow:
        ellipsis;

    white-space:
        nowrap;

    color:
        #AEB8CE;
}


/* ============================================================
   SEPARATOR
   ============================================================ */

.reference-document-separator {

    flex-shrink:
        0;

    margin:
        0 7px;

    color:
        #53617D;
}


/* ============================================================
   PAGE NUMBER
   ============================================================ */

.reference-document-page {

    flex-shrink:
        0;

    color:
        #7F8DAA;

    white-space:
        nowrap;
}


/* ============================================================
   VIEW BUTTON
   IMMEDIATELY AFTER PAGE NUMBER
   ============================================================ */

.reference-view-button {

    flex-shrink: 0;

    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 56px;
    height: 30px;
    margin-left: 7px;
    padding: 0;
    box-sizing: border-box;
    border:
        1px solid
        rgba(91, 105, 142, 0.55);

    border-radius: 8px;

    background:
        rgba(17, 28, 50, 0.90);

    color:
        #F1F4FC !important;

    font-family:
        inherit;

    font-size:
        11px;

    font-weight:
        650;

    line-height:
        1;

    cursor:
        pointer;

    appearance:
        none;

    -webkit-appearance:
        none;

    text-decoration:
        none !important;

    transition:
        background 0.15s ease,
        border-color 0.15s ease,
        transform 0.15s ease;
}


.reference-view-button:hover {

    background:
        rgba(38, 51, 80, 0.98);

    border-color:
        rgba(117, 132, 180, 0.78);

    color:
        #FFFFFF !important;

    transform:
        translateY(-1px);
}


/* ============================================================
   DOWNLOAD BUTTON
   IMMEDIATELY AFTER VIEW
   ============================================================ */

.reference-download-button {

    flex-shrink: 0;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 30px;
    height: 30px;
    margin-left: 6px;
    padding: 0;
    box-sizing: border-box;

    border:
        1px solid
        rgba(91, 105, 142, 0.55);

    border-radius: 8px;

    background:
        rgba(17, 28, 50, 0.90);

    color:
        #F1F4FC !important;

    font-family:
        inherit;

    font-size:
        16px;

    font-weight:
        700;

    line-height:
        1;

    cursor:
        pointer;

    appearance:
        none;

    -webkit-appearance:
        none;

    text-decoration:
        none !important;

    transition:
        background 0.15s ease,
        border-color 0.15s ease,
        transform 0.15s ease;
}


.reference-download-button:hover {

    background:
        rgba(38, 51, 80, 0.98);

    border-color:
        rgba(117, 132, 180, 0.78);

    color:
        #FFFFFF !important;

    transform:
        translateY(-1px);
}


/* ============================================================
   MOBILE
   ============================================================ */

@media (max-width: 900px) {

    .reference-documents-panel {

        width:
            calc(100% - 36px) !important;

        margin-left:
            18px !important;

        margin-right:
            18px !important;
    }

}


@media (max-width: 600px) {

    .reference-documents-panel {

        width:
            calc(100% - 28px) !important;

        margin-left:
            14px !important;

        margin-right:
            14px !important;
    }


    .reference-document-reference {

        font-size:
            10px;
    }


    .reference-view-button,
    .reference-download-button {

        text-decoration: none !important;

        cursor: pointer;

    }

    
    .reference-view-button {

        width:
            52px;

        height:
            30px;

        margin-left:
            6px;
    }


    .reference-download-button {

        width:
            30px;

        height:
            30px;

        margin-left:
            5px;
    }
}

/* ============================================================
   RESPONSIVE ANSWER CARD
   ============================================================ */

@media (max-width: 900px) {

    .st-key-question_panel {
        width: calc(100% - 36px) !important;
        margin-left: 18px !important;
        margin-right: 18px !important;
    }

    .answer-card {
        width: calc(100% - 36px);
        margin-left: 18px;
        margin-right: 18px;
        padding: 25px 22px 22px 22px;
    }

    .answer-art {
        opacity: 0.35;
        right: 10px;
    }

    .answer-text {
        max-width: 100%;
        font-size: 14px;
    }
}


@media (max-width: 600px) {

    .st-key-question_panel {
        width: calc(100% - 28px) !important;
        margin-left: 14px !important;
        margin-right: 14px !important;
    }

    .answer-card {
        width: calc(100% - 28px);

        margin-left: 14px;
        margin-right: 14px;

        padding: 20px 17px 18px 17px;

        border-radius: 15px;
    }

    .answer-icon {
        width: 48px;
        height: 48px;

        margin-right: 13px;

        font-size: 21px;
    }

    .answer-title {
        font-size: 17px;

        margin-bottom: 12px;
    }

    .answer-text {
        font-size: 13px;
        line-height: 1.55;
    }

    .answer-art {
        display: none;
    }

    .answer-sources {
        margin-top: 20px;
    }
}

/* ============================================================
   BOTTOM AI DISCLAIMER — CLEAN CENTERED STYLE
   ============================================================ */

.disclaimer {
    width: 100%;
    margin: 6px 0 4px 0;
    padding: 2px 20px;
    box-sizing: border-box;
    display: flex;
    align-items: flex-start;
    justify-content: center;
    gap: 10px;

    /* REMOVE BOX / BORDER */
    border: none;
    border-radius: 0;
    background: transparent;
    box-shadow: none;

    color: #9FAAC2;
    font-size: 12px;
    line-height: 1.35;

    /* CENTER TEXT */
    text-align: center;
}


/* ============================================================
   SHIELD ICON
   ============================================================ */

.disclaimer-icon {
    flex-shrink: 0;
    width: 28px;
    height: 28px;
    display: flex;
    align-items: center;
    justify-content: center;
    margin-top: 1px;
    background: transparent;
    border: none;
    border-radius: 0;

    font-size: 21px;
}


/* ============================================================
   TEXT CONTENT
   ============================================================ */

.disclaimer-content {
    max-width: 900px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    gap: 2px;
}


/* ============================================================
   FIRST LINE — NORMAL, NOT BOLD
   ============================================================ */

.disclaimer-main {
    color: #9FAAC2;
    font-size: 12px;

    /* NORMAL */
    font-weight: 400;

    line-height: 1.35;

    margin: 0;
    padding: 0;
}


/* ============================================================
   SECOND LINE
   ============================================================ */

.disclaimer-warning {
    color: #9FAAC2;

    font-size: 12px;

    font-weight: 400;

    line-height: 1.35;

    margin: 1px 0 0 0;
    padding: 0;
}


/* ============================================================
   THIRD LINE
   ============================================================ */

.disclaimer-ai {
    color: #9FAAC2;

    font-size: 12px;

    font-weight: 400;

    line-height: 1.35;

    margin: 1px 0 0 0;
    padding: 0;
}


/* ============================================================
   REMOVE ANY BOLD OVERRIDE
   ============================================================ */

.disclaimer-main strong,
.disclaimer-warning strong,
.disclaimer-ai strong {
    font-weight: 400 !important;
}

/* ============================================================
   STREAMLIT EXPANDER
   ============================================================ */

div[data-testid="stExpander"] {

    background:
        rgba(8, 17, 37, 0.90);

    border:
        1px solid
        rgba(65, 91, 160, 0.28);

    border-radius:
        11px;

}


/* ============================================================
   SELECTBOX
   ============================================================ */

div[data-baseweb="select"] > div {

    background:
        rgba(8, 16, 35, 0.85);

    border-color:
        rgba(80, 95, 165, 0.30);

}


/* ============================================================
   SCROLLBAR
   ============================================================ */

::-webkit-scrollbar {

    width:
        6px;

}


::-webkit-scrollbar-track {

    background:
        #050914;

}


::-webkit-scrollbar-thumb {

    background:
        #28345B;

    border-radius:
        10px;

}

/* ============================================================
   FINAL SIDEBAR DESIGN — MATCH SECOND REFERENCE IMAGE
   ============================================================ */

/* SIDEBAR WIDTH */
section[data-testid="stSidebar"] {
    width: 370px !important;
    min-width: 370px !important;
    max-width: 370px !important;

    background:
        linear-gradient(
            180deg,
            #071024 0%,
            #050B1B 100%
        ) !important;

    border-right: 1px solid
        rgba(89, 112, 180, 0.22) !important;
}


/* ============================================================
   SIDEBAR INNER POSITION — MATCH SECOND REFERENCE
   ============================================================ */

section[data-testid="stSidebar"] > div {
    padding-top: 0 !important;
    padding-left: 24px !important;
    padding-right: 24px !important;
    padding-bottom: 0 !important;
}

/* ============================================================
   REMOVE SIDEBAR HEADER HEIGHT — MOVE BRAND TO VERY TOP
   ============================================================ */

section[data-testid="stSidebar"]
div[data-testid="stSidebarHeader"] {
    height: 0 !important;
    min-height: 0 !important;
    max-height: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
    overflow: visible !important;
}

/* Keep sidebar collapse button available */
section[data-testid="stSidebar"]
div[data-testid="stSidebarHeader"] button {
    position: absolute !important;
    top: 8px !important;
    right: 8px !important;
    z-index: 9999 !important;
}

/* Remove Streamlit's internal top spacing */
section[data-testid="stSidebar"]
div[data-testid="stSidebarContent"] {
    padding-top: 0 !important;
}

section[data-testid="stSidebar"]
div[data-testid="stSidebarUserContent"] {
    padding-top: 18px !important;
    margin-top: 0 !important;
}

/* Remove first element top gap */
section[data-testid="stSidebar"]
div[data-testid="stSidebarUserContent"]
> div:first-child {
    margin-top: 0 !important;
    padding-top: 0 !important;
}

/* ============================================================
   BRAND — MATCH SECOND REFERENCE
   ============================================================ */

.sidebar-brand {
    width: 100% !important;
    margin: 0 !important;
    padding: 0 0 10px 0 !important;
    text-align: left !important;
}


/* LOGO + NAGARSATHI SAME LINE */
.sidebar-brand-header {
    display: flex !important;
    align-items: center !important;
    width: 100% !important;
    gap: 18px !important;
    margin: 0 !important;
    padding: 0 !important;
}


/* LOGO */
.sidebar-brand-header .sidebar-brand-icon {
    width: 58px !important;
    height: 58px !important;

    flex: 0 0 58px !important;

    display: flex !important;
    align-items: center !important;
    justify-content: center !important;

    margin: 0 !important;
    padding: 0 !important;

    border-radius: 50% !important;

    background:
        radial-gradient(
            circle,
            rgba(109, 76, 255, 0.45) 0%,
            rgba(74, 54, 180, 0.12) 58%,
            transparent 74%
        ) !important;

    font-size: 34px !important;

    filter:
        drop-shadow(
            0 0 11px
            rgba(104, 86, 255, 0.60)
        ) !important;
}

/* NAGARSATHI */
.sidebar-brand-header .sidebar-brand-title {
    margin: 0 !important;
    padding: 0 !important;

    color: #F5F7FF !important;

    font-size: 26px !important;
    line-height: 1.05 !important;

    font-weight: 800 !important;

    letter-spacing: -0.8px !important;

    white-space: nowrap !important;
}


/* GRADIENT SATHI */
.sidebar-brand-title .gradient {
    background:
        linear-gradient(
            90deg,
            #6384FF 0%,
            #B15DFF 100%
        ) !important;

    -webkit-background-clip: text !important;
    background-clip: text !important;

    -webkit-text-fill-color: transparent !important;
}


/* SUBTITLE */
.sidebar-brand-subtitle {
    margin: 6px 0 0 76px !important;
    padding: 0 !important;
    color: #8E9AB7 !important;
    font-size: 13px !important;
    line-height: 1.45 !important;
    font-weight: 400 !important;
}


/* ============================================================
   EXPLORE TITLE
   ============================================================ */

section[data-testid="stSidebar"] .nav-section {
    margin: 0 0 3px 10px !important;
    padding: 0 !important;
    color: #68738F !important;
    font-size: 11px !important;
    font-weight: 700 !important;
    letter-spacing: 1.5px !important;
    line-height: 1.2 !important;
}


/* ============================================================
   NAVIGATION BUTTONS
   ============================================================ */

section[data-testid="stSidebar"] .stButton {
    width: 100% !important;

    margin: 0 !important;
    padding: 0 !important;
}


section[data-testid="stSidebar"] .stButton > button {
    width: 100% !important;
    height: 34px !important;
    min-height: 34px !important;
    max-height: 34px !important;
    margin: 0 !important;
    padding: 0 10px !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    border-radius: 9px !important;
    color: #E5E9F5 !important;
    display: flex !important;
    align-items: center !important;
    justify-content: flex-start !important;
    text-align: left !important;
    font-size: 14px !important;
    font-weight: 500 !important;
    line-height: 1 !important;
    box-shadow: none !important;
    transform: none !important;
}


/* NAVIGATION TEXT */
section[data-testid="stSidebar"]
.stButton > button p {

    margin: 0 !important;
    padding: 0 !important;

    color: #E5E9F5 !important;

    font-size: 14px !important;

    font-weight: 500 !important;

    line-height: 1 !important;
}


/* NAVIGATION HOVER */
section[data-testid="stSidebar"]
.stButton > button:hover {

    background:
        rgba(84, 72, 180, 0.18) !important;

    border-color:
        rgba(112, 92, 240, 0.32) !important;

    color: #FFFFFF !important;
}


/* ============================================================
   SYSTEM STATUS
   ============================================================ */

section[data-testid="stSidebar"] .status-panel {
    width: calc(100% - 20px) !important;
    margin: 6px 10px 0 10px !important;
    padding: 11px 14px !important;
    border-radius: 11px !important;

    background:
        linear-gradient(
            145deg,
            rgba(15, 29, 59, 0.96),
            rgba(9, 17, 37, 0.98)
        ) !important;

    border:
        1px solid
        rgba(78, 104, 170, 0.25) !important;

    box-shadow:
        0 8px 20px
        rgba(0,0,0,0.12) !important;

    box-sizing: border-box !important;
}


/* STATUS HEADING */
section[data-testid="stSidebar"] .status-heading {
    margin: 0 0 5px 0 !important;
    color: #FFFFFF !important;
    font-size: 13px !important;
    font-weight: 700 !important;
    line-height: 1.2 !important;
}


/* ONLINE */
section[data-testid="stSidebar"] .status-online {
    margin: 0 0 4px 0 !important;
    color: #2ED995 !important;
    font-size: 11px !important;
    line-height: 1.2 !important;
}


/* STATUS ROW */
section[data-testid="stSidebar"] .status-line {

    display: flex !important;
    align-items: center !important;
    justify-content: space-between !important;
    width: 100% !important;
    padding: 3px 0 !important;
    margin: 0 !important;
    color: #8B97B3 !important;
    font-size: 11px !important;
    line-height: 1.2 !important;
}


/* STATUS NUMBERS */
section[data-testid="stSidebar"] .status-number {

    color: #F4F6FF !important;

    font-weight: 700 !important;
}


/* READY */
section[data-testid="stSidebar"] .ready-badge {

    padding: 3px 8px !important;

    border-radius: 20px !important;

    background:
        rgba(37, 209, 139, 0.17) !important;

    color: #36E19A !important;

    font-size: 9px !important;

    font-weight: 700 !important;
}


/* ============================================================
   REFRESH KNOWLEDGE BASE
   ============================================================ */

section[data-testid="stSidebar"]
.status-panel + div {

    margin-top: 3px !important;
    margin-bottom: 0 !important;
}


section[data-testid="stSidebar"]
.status-panel + div .stButton > button {

    height: 32px !important;
    min-height: 32px !important;
    max-height: 32px !important;
    margin: 0 !important;
    padding: 0 10px !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    color: #E5E9F5 !important;
    font-size: 14px !important;
    line-height: 1 !important;
}


/* ============================================================
   TRUSTED & SECURE
   ============================================================ */

section[data-testid="stSidebar"] .trust-panel {
    width: calc(100% - 20px) !important;
    margin: 3px 10px 0 10px !important;
    padding: 10px 12px !important;
    min-height: 70px !important;   
    box-sizing: border-box !important;
    border-radius: 10px !important;

    background:
        linear-gradient(
            135deg,
            rgba(82, 38, 160, 0.76),
            rgba(56, 28, 120, 0.52)
        ) !important;

    border:
        1px solid
        rgba(155, 100, 255, 0.30) !important;

    box-shadow:
        0 8px 22px
        rgba(66, 30, 150, 0.16) !important;
}


/* TRUST TITLE */
section[data-testid="stSidebar"] .trust-title {

    margin: 0 0 3px 0 !important;
    color: #FFFFFF !important;
    font-size: 12px !important;
    font-weight: 700 !important;
    line-height: 1.2 !important;
}


/* TRUST TEXT */
section[data-testid="stSidebar"] .trust-text {
    margin: 0 !important;
    color: #C5CBE0 !important;
    font-size: 11px !important;
    line-height: 1.4 !important;
}


/* ============================================================
   FOOTER
   ============================================================ */

section[data-testid="stSidebar"] .sidebar-footer {
    margin-top: 10px !important;
    padding: 0 2px 4px 2px !important;
    text-align: left !important;
}


section[data-testid="stSidebar"] .sidebar-footer-copy {
    color: #8E9AB7 !important;
    font-size: 10px !important;
    line-height: 1.25 !important;
}


section[data-testid="stSidebar"] .sidebar-footer-project {
    margin-top: 2px !important;
    color: #8E9AB7 !important;
    font-size: 10px !important;
    line-height: 1.25 !important;
}


section[data-testid="stSidebar"] .sidebar-footer-developer {
    margin-top: 6px !important;
    color: #A9B2C8 !important;
    font-size: 10px !important;
    line-height: 1.25 !important;
}


section[data-testid="stSidebar"]
.sidebar-footer-developer strong {

    color: #C9B8FF !important;
    font-weight: 600 !important;
}


/* ============================================================
   MOBILE / SMALL SCREEN
   ============================================================ */

@media (max-width: 600px) {

    section[data-testid="stSidebar"] {
        width: 310px !important;
        min-width: 310px !important;
        max-width: 310px !important;
    }

    section[data-testid="stSidebar"] > div {
        padding-left: 20px !important;
        padding-right: 20px !important;
    }

    section[data-testid="stSidebar"]
    div[data-testid="stSidebarUserContent"] {
        padding-top: 14px !important;
    }

    .sidebar-brand {
        padding-bottom: 8px !important;
    }

    .sidebar-brand-header {
        gap: 14px !important;
    }

    .sidebar-brand-header .sidebar-brand-icon {
        width: 52px !important;
        height: 52px !important;
        flex-basis: 52px !important;
        font-size: 30px !important;
    }

    .sidebar-brand-header .sidebar-brand-title {
        font-size: 22px !important;
    }

    .sidebar-brand-subtitle {
        margin: 5px 0 0 66px !important;
        font-size: 11px !important;
        line-height: 1.35 !important;
    }

    section[data-testid="stSidebar"] .nav-section {
        margin-bottom: 4px !important;
    }

    section[data-testid="stSidebar"] .stButton > button {
        height: 34px !important;
        min-height: 34px !important;
        max-height: 34px !important;
    }

    section[data-testid="stSidebar"] .status-panel {
        margin-top: 7px !important;
        padding: 10px 12px !important;
    }

    section[data-testid="stSidebar"]
    .status-panel + div {
        margin-top: 3px !important;
    }

    section[data-testid="stSidebar"] .trust-panel {
        margin-top: 3px !important;
        padding: 9px 10px !important;
        min-height: 66px !important;
    }

    section[data-testid="stSidebar"] .sidebar-footer {
        margin-top: 8px !important;
        padding-bottom: 3px !important;
    }
}

</style>
""",
    unsafe_allow_html=True
)


# ============================================================
# 5. HELPER FUNCTIONS
# ============================================================

def count_sources():
    count = 0

    for folder in [
        PDF_DIR,
        WORD_DIR,
        EXCEL_DIR
    ]:
        # Only scan paths that are actually directories.
        if folder.is_dir():
            try:
                count += sum(
                    1
                    for item in folder.iterdir()
                    if item.is_file()
                )
            except (OSError, PermissionError):
                continue

    # Count configured URLs safely.
    if URL_FILE.is_file():
        try:
            count += sum(
                1
                for line in URL_FILE.read_text(
                    encoding="utf-8"
                ).splitlines()
                if line.strip()
            )
        except (OSError, UnicodeDecodeError):
            pass

    return count


# ============================================================
# PDF REFERENCE HELPERS
# ============================================================

def get_pdf_path(source_file):

    source_file = str(source_file)

    if source_file.startswith(
        ("http://", "https://")
    ):
        return None

    pdf_path = Path(source_file)

    if not pdf_path.is_absolute():
        pdf_path = PDF_DIR / pdf_path.name

    if pdf_path.exists():
        return pdf_path

    return None


def get_pdf_page_number(page):

    if page is None:
        return 1

    try:
        page_number = int(page)

        # Your metadata may be zero-based.
        # PDF viewer pages are one-based.
        return max(1, page_number)

    except Exception:
        return 1

    
# ============================================================
# COUNT CHUNKS
# ============================================================

def count_chunks():

    index_file = (
        VECTORSTORE_DIR /
        "index.faiss"
    )


    if not index_file.exists():

        return 0


    try:

        embeddings = OpenAIEmbeddings(

            model="text-embedding-3-small"

        )


        db = FAISS.load_local(

            str(VECTORSTORE_DIR),

            embeddings,

            allow_dangerous_deserialization=True

        )


        return len(
            db.index_to_docstore_id
        )


    except Exception:

        return 0


# ============================================================
# LAST UPDATED
# ============================================================

def get_last_updated():

    if not VECTORSTORE_DIR.exists():

        return "Not built"


    files = list(
        VECTORSTORE_DIR.glob("*")
    )


    if not files:

        return "Not built"


    latest = max(

        file.stat().st_mtime

        for file in files

    )


    return datetime.fromtimestamp(
        latest
    ).strftime(
        "%b %d, %Y"
    )

# ============================================================
# EXTRACT DOCUMENT TITLE FROM PDF
# ============================================================

@st.cache_data(show_spinner=False)
def extract_document_title(source_file):

    try:

        source_file = str(source_file)

        # ----------------------------------------------------
        # Do not process URLs
        # ----------------------------------------------------

        if source_file.startswith(
            ("http://", "https://")
        ):
            return None

        # ----------------------------------------------------
        # Find PDF
        # ----------------------------------------------------

        pdf_path = Path(source_file)

        if not pdf_path.is_absolute():
            pdf_path = PDF_DIR / pdf_path.name

        if not pdf_path.exists():
            return None

        # ----------------------------------------------------
        # Read first page
        # ----------------------------------------------------

        reader = PdfReader(
            str(pdf_path)
        )

        if len(reader.pages) == 0:
            return None

        text = reader.pages[0].extract_text() or ""

        if not text.strip():
            return None

        # ----------------------------------------------------
        # Clean text
        # ----------------------------------------------------

        lines = []

        for line in text.splitlines():

            line = " ".join(
                line.split()
            ).strip()

            if not line:
                continue

            # Ignore page-number-only lines
            if line.isdigit():
                continue

            lines.append(line)

        if not lines:
            return None

        # ----------------------------------------------------
        # Find Act / Rule / Notification title
        # ----------------------------------------------------

        keywords = (
            "ACT",
            "RULE",
            "RULES",
            "REGULATION",
            "REGULATIONS",
            "NOTIFICATION",
            "ORDER",
            "BYE-LAW",
            "BYE-LAWS",
            "ORDINANCE",
            "SCHEME"
        )

        title_lines = []

        for i, line in enumerate(lines):

            upper_line = line.upper()

            if any(
                keyword in upper_line
                for keyword in keywords
            ):

                # Take the lines immediately before
                # and around the Act/Rule heading
                start = max(
                    0,
                    i - 4
                )

                title_lines = lines[
                    start:i + 1
                ]

                # Remove obvious document numbers
                title_lines = [
                    x
                    for x in title_lines
                    if not (
                        x.startswith("(")
                        and "OF" in x.upper()
                        and ")" in x
                    )
                ]

                break

        # ----------------------------------------------------
        # Build title
        # ----------------------------------------------------

        if title_lines:

            title = " ".join(
                title_lines
            )

            title = " ".join(
                title.split()
            ).strip()

            if len(title) >= 20:
                return title

        # ----------------------------------------------------
        # Fallback
        # ----------------------------------------------------

        for line in lines[:10]:

            if len(line) >= 25:

                return " ".join(
                    line.split()
                ).strip()

    except Exception:
        return None

    return None


# ============================================================
# VECTORSTORE
# ============================================================

@st.cache_resource
def get_vectorstore():

    embeddings = OpenAIEmbeddings(

        model="text-embedding-3-small"

    )


    return FAISS.load_local(

        str(VECTORSTORE_DIR),

        embeddings,

        allow_dangerous_deserialization=True

    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    # --------------------------------------------------------
    # BRAND
    # --------------------------------------------------------

    render_html("""
        <div class="sidebar-brand">

            <div class="sidebar-brand-header">

                <div class="sidebar-brand-icon">
                    🏛️
                </div>

                <div class="sidebar-brand-title">
                    Nagar<span class="gradient">Sathi</span>
                </div>

            </div>

            <div class="sidebar-brand-subtitle">
                AI-Powered Municipal<br>
                Knowledge & Service Assistant
            </div>

        </div>
    """)


    # --------------------------------------------------------
    # HOME
    # --------------------------------------------------------

    st.button(
        ":material/home:  Home",
        key="home_button",
        use_container_width=True
    )


    # --------------------------------------------------------
    # EXPLORE
    # --------------------------------------------------------

    render_html("""
        <div class="nav-section">
            EXPLORE
        </div>
    """)


    st.button(
        ":material/menu_book:  Knowledge Base",
        use_container_width=True
    )

    st.button(
        ":material/grid_view:  Services",
        use_container_width=True
    )

    st.button(
        ":material/account_balance:  Acts & Rules",
        use_container_width=True
    )

    st.button(
        ":material/notifications_none:  Notifications",
        use_container_width=True
    )

    st.button(
        ":material/help_outline:  Help Center",
        use_container_width=True
    )


    # --------------------------------------------------------
    # SYSTEM STATUS
    # --------------------------------------------------------

    source_count = count_sources()
    updated = get_last_updated()

    render_html(f"""
        <div class="status-panel">

            <div class="status-heading">
                System Status
            </div>

            <div class="status-online">
                ● &nbsp; All Systems Operational
            </div>

            <div class="status-line">
                <span>
                    Vector Database
                </span>

                <span class="ready-badge">
                    Ready
                </span>
            </div>

            <div class="status-line">
                <span>
                    Documents Indexed
                </span>

                <span class="status-number">
                    {source_count}
                </span>
            </div>

           
            <div class="status-line">
                <span>
                    Last Updated
                </span>

                <span class="status-number">
                    {updated}
                </span>
            </div>

        </div>
    """)


    # --------------------------------------------------------
    # REFRESH KNOWLEDGE BASE
    # --------------------------------------------------------

    if st.button(
        "⟳   Refresh Knowledge Base",
        use_container_width=True
    ):

        with st.spinner(
            "Rebuilding municipal knowledge base..."
        ):

            from rag_build import build_vectorstore
            success = build_vectorstore()


        if success:

            st.cache_resource.clear()

            st.success(
                "Knowledge Base refreshed."
            )

            st.rerun()

        else:

            st.error(
                "No municipal documents found."
            )


    # --------------------------------------------------------
    # TRUSTED & SECURE
    # --------------------------------------------------------

    render_html("""
        <div class="trust-panel">

            <div class="trust-title">
                🛡️ &nbsp; Trusted & Secure
            </div>

            <div class="trust-text">
                Your information is protected.
                Answers are grounded in the
                municipal knowledge base.
            </div>

        </div>
    """)


    # --------------------------------------------------------
    # FOOTER
    # --------------------------------------------------------

    render_html("""
        <div class="sidebar-footer">

            <div class="sidebar-footer-copy">
                © 2026 NagarSathi
            </div>

            <div class="sidebar-footer-project">
                Municipal Knowledge Assistant
            </div>

            <div class="sidebar-footer-developer">
                👨‍💻 Developed By
                <strong>Heramba Kakati</strong>
            </div>

        </div>
    """)

# ============================================================
# HERO IMAGE — USE ORIGINAL IMAGE
# ============================================================

hero_image_base64 = None

from io import BytesIO


def load_original_hero_image(image_path):
    """
    Load the original municipal_ai.png without changing its pixels.

    The image itself is NOT recolored, background-removed, blurred,
    sharpened, or converted through an alpha-threshold algorithm.

    Boundary blending is handled by CSS only, so the original
    building, platform, icons, neon glow, and internal background
    remain exactly as supplied.
    """

    with open(image_path, "rb") as image_file:
        return base64.b64encode(
            image_file.read()
        ).decode("utf-8")


if HERO_IMAGE.exists():

    try:

        hero_image_base64 = load_original_hero_image(
            HERO_IMAGE
        )

    except Exception:

        hero_image_base64 = None


# ============================================================
# HERO
# ============================================================

if hero_image_base64:

    st.html(
        f"""
        <div class="hero-wrapper">

            <div class="hero-content">

                <div class="hero-greeting">
                    Hello! 👋
                </div>

                <div class="hero-title">
                    How can I
                    <br>
                    <span class="blue">
                        help you today?
                    </span>
                </div>

                <div class="hero-description">
                Ask any municipal question and get answers<br>
                from official documents, Acts, Rules,<br>
                Regulations and notifications.
            </div>

            </div>

            <div class="hero-visual">

                <img
                    src="data:image/png;base64,{hero_image_base64}"
                    alt="NagarSathi Municipal AI"
                >

            </div>

        </div>
        """
    )

else:

    st.html(
        """
        <div class="hero-wrapper">

            <div class="hero-content">

                <div class="hero-greeting">
                    Hello! 👋
                </div>

                <div class="hero-title">
                    How can I
                    <br>
                    <span class="blue">
                        help you today?
                    </span>
                </div>

                <div class="hero-description">
                    Ask any municipal question and get answers
                    from official documents, Acts, Rules,
                    Regulations and notifications.
                </div>

            </div>

        </div>
        """
    )

    st.warning(
        "Hero image not found: "
        "assets/municipal_ai.png"
    )


# ============================================================
# QUESTION PANEL
# ============================================================

with st.container(key="question_panel"):

    # --------------------------------------------------------
    # QUESTION LABEL
    # --------------------------------------------------------

    st.markdown(
        """
        <div class="question-label">
            Ask your municipal question:
        </div>
        """,
        unsafe_allow_html=True
    )


    # --------------------------------------------------------
    # SEARCH ROW
    # --------------------------------------------------------

    input_col, search_col = st.columns(
        [1, 0.075],
        gap="small"
    )


    # --------------------------------------------------------
    # SEARCH INPUT
    # --------------------------------------------------------

    with input_col:

        query = st.text_input(
            "Question",

            placeholder=(
                "What is the procedure for registering "
                "a public grievance?"
            ),

            label_visibility="collapsed"
        )


        # ========================================================
        # SUGGESTIONS — FULL WIDTH INSIDE QUESTION CARD
        # ========================================================

        suggestion_cols = st.columns(
            [
                0.82,   # Try these
                1.00,   # Property tax
                1.18,   # Building permission
                0.98,   # Trade license
                1.28,   # Electricity connection
                0.88,   # Sanitation
                1.15    # Birth certificate
            ],
            gap="small"
        )

        # --------------------------------------------------------
        # TRY THESE LABEL
        # --------------------------------------------------------

        with suggestion_cols[0]:
            st.markdown(
                """
                <div class="suggestion-title">
                    <span class="sparkle">✨</span>
                    <span>Try these:</span>
                </div>
                """,
                unsafe_allow_html=True
            )

        # --------------------------------------------------------
        # SUGGESTIONS
        # --------------------------------------------------------

        suggestions = [
            "🏠  Property tax",
            "🏗️  Building permission",
            "📜  Trade license",
            "⚡  Electricity connection",
            "♻️  Sanitation",
            "🎂  Birth certificate"
        ]

        # --------------------------------------------------------
        # CREATE SIX BUTTONS
        # --------------------------------------------------------

        for index, (column, suggestion) in enumerate(
            zip(
                suggestion_cols[1:],
                suggestions
            ),
            start=1
        ):

            with column:

                with st.container(
                    key=f"suggestion_{index}"
                ):

                    clicked = st.button(
                        suggestion,
                        use_container_width=True
                    )

                if clicked:
                    query = suggestion.split(
                        "  ",
                        1
                    )[-1]


    # --------------------------------------------------------
    # SEARCH BUTTON
    # --------------------------------------------------------

    with search_col:

        with st.container(
            key="search_button"
        ):

            search_clicked = st.button(
                "➤",
                use_container_width=True
            )

# ============================================================
# QUESTION PROCESSING
# ============================================================

if query:

    if not (
        VECTORSTORE_DIR /
        "index.faiss"
    ).exists():

        st.error(
            "Knowledge Base is not available."
        )

        st.info(
            "Place your files inside the data folders "
            "and run: python rag_build.py"
        )

        st.stop()


    try:

        # ----------------------------------------------------
        # LOAD DATABASE
        # ----------------------------------------------------

        vectorstore = get_vectorstore()


        # ----------------------------------------------------
        # RETRIEVER
        # ----------------------------------------------------

        retriever = vectorstore.as_retriever(

            search_kwargs={
                "k": 6
            }

        )


        # ----------------------------------------------------
        # SEARCH
        # ----------------------------------------------------

        with st.spinner(
            "Searching official municipal documents..."
        ):

            docs = retriever.invoke(
                query
            )


        # ----------------------------------------------------
        # BUILD CONTEXT
        # ----------------------------------------------------

        context_parts = []


        for index, doc in enumerate(
            docs,
            start=1
        ):

            metadata = doc.metadata


            source = metadata.get(
                "source_file",
                "Unknown source"
            )


            file_type = metadata.get(
                "file_type",
                "Unknown"
            )


            page = metadata.get(
                "page"
            )


            sheet = metadata.get(
                "sheet"
            )


            source_location = source


            if page:

                source_location += (
                    f" | Page {page}"
                )


            if sheet:

                source_location += (
                    f" | Sheet {sheet}"
                )


            context_parts.append(

                f"""
[Source {index}]

Source:
{source_location}

Type:
{file_type}

Content:
{doc.page_content}
"""

            )


        context = "\n\n".join(
            context_parts
        )


        # ----------------------------------------------------
        # PROMPT
        # ----------------------------------------------------

        prompt = ChatPromptTemplate.from_template(

            """
You are NagarSathi, an AI assistant for a Municipal Board.

Answer the user's question using ONLY the supplied
municipal CONTEXT.

RULES:

1. Never invent information.

2. Never guess municipal rules, fees,
   dates, sections or procedures.

3. Do not use outside knowledge.

4. Preserve:
   - Act names
   - Section numbers
   - Rule numbers
   - Dates
   - Fees
   - Conditions
   - Exceptions

5. Use simple English.

6. Use bullets where appropriate.

7. Do not include [Source 1], [Source 2] or any
   other source labels in the final answer.

8. Source information is displayed separately by the application.

9. If the answer is not present in the context,
   say exactly:

"I could not find sufficient information in the
available municipal documents to answer this question."

10. Do not reveal internal reasoning.

CONTEXT:

{context}

QUESTION:

{question}

ANSWER:
"""

        )


        # ----------------------------------------------------
        # MODEL
        # ----------------------------------------------------

        llm = ChatOpenAI(

            model="gpt-4.1-mini",

            temperature=0

        )


        chain = (

            prompt

            | llm

            | StrOutputParser()

        )


        # ----------------------------------------------------
        # ANSWER
        # ----------------------------------------------------

        response = chain.invoke(

            {

                "context":
                    context,

                "question":
                    query

            }

        )


        # ====================================================
        # ANSWER CARD — REFERENCE DESIGN
        # ====================================================

        import html
        import re

        # ----------------------------------------------------
        # PREPARE RESPONSE FOR HTML
        # ----------------------------------------------------

        # Remove [Source 1], [Source 2], etc.
        # because sources are already displayed separately below.
        clean_response = re.sub(
            r"\[Source\s+\d+\]",
            "",
            response,
            flags=re.IGNORECASE
        )

        # Remove excessive blank lines created after removing citations
        clean_response = re.sub(
            r"\n\s*\n+",
            "\n",
            clean_response
        ).strip()

        # ============================================================
        # DETECT "NO SUFFICIENT INFORMATION" RESPONSE
        # ============================================================

        insufficient_information = (
            "I could not find sufficient information in the "
            "available municipal documents to answer this question."
            in clean_response.strip()
        )

        safe_response = html.escape(clean_response)

        # Preserve line breaks from the LLM response
        safe_response = safe_response.replace("\n", "<br>")


        # ----------------------------------------------------
        # BUILD SOURCE DETAILS
        # ----------------------------------------------------

        source_items_html = ""

        # Store PDF references separately.
        # Streamlit buttons will be rendered after the HTML card.
        pdf_references = []


        for index, doc in enumerate(
            docs,
            start=1
        ):

            metadata = doc.metadata

            # ------------------------------------------------
            # SOURCE FILE / URL
            # ------------------------------------------------

            source_file = metadata.get(
                "source_file",
                "Unknown source"
            )

            # ------------------------------------------------
            # PAGE
            # ------------------------------------------------

            page = metadata.get(
                "page"
            )

            # ------------------------------------------------
            # SHEET
            # ------------------------------------------------

            sheet = metadata.get(
                "sheet"
            )

            # ------------------------------------------------
            # CHECK URL
            # ------------------------------------------------

            is_url = str(
                source_file
            ).startswith(
                ("http://", "https://")
            )

            # =================================================
            # WEB SOURCE
            # =================================================

            if is_url:

                safe_url = html.escape(
                    str(source_file),
                    quote=True
                )

                source_reference = f"""
                <a
                    href="{safe_url}"
                    target="_blank"
                    class="source-link"
                >
                    {html.escape(str(source_file))}
                </a>
                """

            # =================================================
            # PDF / DOCUMENT SOURCE
            # =================================================

            else:

                document_title = (
                    metadata.get("document_title")
                    or metadata.get("title")
                    or metadata.get("heading")
                    or metadata.get("notification_title")
                )

                # ------------------------------------------------
                # Extract internal PDF heading
                # ------------------------------------------------

                if not document_title:

                    document_title = extract_document_title(
                        source_file
                    )

                # ------------------------------------------------
                # Final fallback
                # ------------------------------------------------

                if not document_title:

                    document_title = Path(
                        str(source_file)
                    ).stem

                source_reference = html.escape(
                    str(document_title)
                )

                # ------------------------------------------------
                # Store PDF for View / Download buttons
                # ------------------------------------------------

                pdf_path = get_pdf_path(
                    source_file
                )

                if pdf_path:

                    pdf_references.append({
                        "index": index,
                        "title": str(document_title),
                        "path": pdf_path,
                        "page": get_pdf_page_number(page)
                    })

            # ------------------------------------------------
            # PAGE NUMBER
            # ------------------------------------------------

            if page is not None:

                source_reference += (
                    f" | Page "
                    f"{html.escape(str(page))}"
                )

            # ------------------------------------------------
            # SHEET
            # ------------------------------------------------

            if sheet:

                source_reference += (
                    f" | Sheet "
                    f"{html.escape(str(sheet))}"
                )

            # ------------------------------------------------
            # SOURCE CARD
            # ------------------------------------------------

            source_items_html += f"""
            <div class="source-item">

                <div class="source-item-title">
                    Source {index}
                </div>

                <div class="source-item-reference">
                    {source_reference}
                </div>

            </div>
            """

        # ============================================================
        # ANSWER CARD STYLE
        # ============================================================

        answer_card_class = (
            "answer-card no-info-card"
            if insufficient_information
            else "answer-card"
        )

        # ----------------------------------------------------
        # COMPLETE ANSWER CARD
        # ----------------------------------------------------

        render_html(
            f"""
            <div class="{answer_card_class}">

                <!-- ==========================================
                    RIGHT DECORATIVE ARTWORK
                    ========================================== -->

                <div class="answer-art">

                    <div class="answer-art-circle"></div>

                    <div class="answer-art-search"></div>

                    <div class="answer-art-star one">✦</div>
                    <div class="answer-art-star two">✦</div>
                    <div class="answer-art-star three">✦</div>

                </div>


                <!-- ==========================================
                    ANSWER HEADER + RESPONSE
                    ========================================== -->

                <div class="answer-top">
                    <div class="answer-icon">
                        📄
                    </div>

                    <div class="answer-content">
                        <div class="answer-title">
                            Here's what I found
                        </div>

                        <div class="answer-text">
                            {safe_response}
                        </div>
                    </div>
                </div>
            </div>
            """
        )


        #============================================================
        # REFERENCE DOCUMENTS
        # VIEW = OPEN PDF IN NEW TAB
        # DOWNLOAD = DIRECT PDF DOWNLOAD
        # ============================================================
        
        if pdf_references:
        
            reference_rows = ""
        
            for ref in pdf_references:
        
                # --------------------------------------------------------
                # READ PDF
                # --------------------------------------------------------
        
                try:
                    with open(ref["path"], "rb") as pdf_file:
                        pdf_bytes = pdf_file.read()
        
                except Exception:
                    continue
        
                # --------------------------------------------------------
                # CONVERT PDF TO BASE64
                # --------------------------------------------------------
        
                pdf_base64 = base64.b64encode(
                    pdf_bytes
                ).decode("utf-8")
        
                # --------------------------------------------------------
                # SAFE DISPLAY VALUES
                # --------------------------------------------------------
        
                source_index = html.escape(
                    str(ref["index"])
                )
        
                document_title = html.escape(
                    str(ref["title"])
                )
        
                page_number = html.escape(
                    str(ref["page"])
                )
        
                download_filename = html.escape(
                    ref["path"].name,
                    quote=True
                )
        
                # --------------------------------------------------------
                # ONE REFERENCE ROW
                # --------------------------------------------------------
        
                reference_rows += f"""
                <div class="reference-document-row">
        
                    <div class="reference-document-main">
        
                        <div class="reference-document-icon">
                            📄
                        </div>
        
                        <div class="reference-document-content">
        
                            <div class="reference-document-title">
                                Reference {source_index}
                            </div>
        
                            <div class="reference-document-reference">
        
                                <span class="reference-document-name">
                                    {document_title}
                                </span>
        
                                <span class="reference-document-separator">
                                    |
                                </span>
        
                                <span class="reference-document-page">
                                    Page {page_number}
                                </span>
        
                                <!-- =====================================
                                     VIEW BUTTON
                                     ===================================== -->
        
                                <button
                                    type="button"
                                    class="reference-view-button"
                                    onclick="viewPDF(
                                        '{pdf_base64}',
                                        {int(ref["page"])}
                                    )"
                                >
                                    View
                                </button>
        
                                <!-- =====================================
                                     DOWNLOAD BUTTON
                                     ===================================== -->
        
                                <button
                                    type="button"
                                    class="reference-download-button"
                                    onclick="downloadPDF(
                                        '{pdf_base64}',
                                        '{download_filename}'
                                    )"
                                    title="Download Reference {source_index}"
                                >
                                    ↓
                                </button>
        
                            </div>
        
                        </div>
        
                    </div>
        
                </div>
                """
        
            # ============================================================
            # COMPLETE REFERENCE PANEL
            # ============================================================
        
            reference_html = f"""
            <!DOCTYPE html>
        
            <html>
        
            <head>
        
                <meta charset="UTF-8">
        
                <style>
        
                    * {{
                        box-sizing: border-box;
                    }}
        
                    html,
                    body {{
                        margin: 0;
                        padding: 0;
                        background: transparent;
                        font-family:
                            Inter,
                            -apple-system,
                            BlinkMacSystemFont,
                            "Segoe UI",
                            sans-serif;
                    }}
        
                    .reference-documents-panel {{
                        width: 100%;
                        margin: 0;
                        padding: 0;
        
                        border:
                            1px solid
                            rgba(62, 91, 157, 0.34);
        
                        border-radius: 11px;
        
                        background:
                            linear-gradient(
                                180deg,
                                rgba(11, 27, 56, 0.92),
                                rgba(8, 21, 46, 0.92)
                            );
        
                        overflow: hidden;
                    }}
        
                    .reference-header {{
                        width: 100%;
                        min-height: 46px;
        
                        padding: 0 18px;
        
                        display: flex;
                        align-items: center;
        
                        color: #E6EAF4;
        
                        font-size: 14px;
                        font-weight: 650;
                    }}
        
                    .reference-folder-icon {{
                        flex-shrink: 0;
                        margin-right: 9px;
                        font-size: 17px;
                        line-height: 1;
                    }}
        
                    .reference-panel-title {{
                        color: #E6EAF4;
                        white-space: nowrap;
                    }}
        
                    .reference-panel-count {{
                        margin-left: 6px;
                        color: #8996B5;
                        font-size: 13px;
                        font-weight: 500;
                    }}
        
                    .reference-document-list {{
                        width: 100%;
        
                        padding: 3px 18px 10px 18px;
        
                        border-top:
                            1px solid
                            rgba(65, 91, 155, 0.22);
                    }}
        
                    .reference-document-row {{
                        width: 100%;
                        min-height: 62px;
        
                        display: flex;
                        align-items: center;
        
                        border-bottom:
                            1px solid
                            rgba(65, 91, 155, 0.14);
                    }}
        
                    .reference-document-row:last-child {{
                        border-bottom: none;
                    }}
        
                    .reference-document-main {{
                        width: 100%;
        
                        display: flex;
                        align-items: center;
        
                        min-width: 0;
                    }}
        
                    .reference-document-icon {{
                        width: 30px;
                        min-width: 30px;
        
                        display: flex;
                        align-items: center;
                        justify-content: flex-start;
        
                        font-size: 16px;
                    }}
        
                    .reference-document-content {{
                        width: 100%;
                        min-width: 0;
                    }}
        
                    .reference-document-title {{
                        margin-bottom: 3px;
        
                        color: #F2F5FF;
        
                        font-size: 12px;
                        font-weight: 700;
                        line-height: 1.2;
                    }}
        
                    .reference-document-reference {{
                        width: 100%;
        
                        display: flex;
                        align-items: center;
        
                        min-width: 0;
        
                        color: #B7C2DD;
        
                        font-size: 12px;
                        line-height: 1.3;
                    }}
        
                    .reference-document-name {{
                        min-width: 0;
                        max-width: calc(100% - 125px);
        
                        overflow: hidden;
                        white-space: nowrap;
                        text-overflow: ellipsis;
        
                        color: #B8C7E6;
                    }}
        
                    .reference-document-separator {{
                        margin: 0 8px;
                        color: #6D7A99;
                    }}
        
                    .reference-document-page {{
                        flex-shrink: 0;
                        color: #9AA8C7;
                        white-space: nowrap;
                    }}
        
                    /* ====================================================
                       VIEW BUTTON
                       ==================================================== */
        
                    .reference-view-button {{
                        flex-shrink: 0;
        
                        width: 52px;
                        height: 34px;
        
                        margin-left: 10px;
        
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
        
                        padding: 0;
        
                        border:
                            1px solid
                            rgba(89, 112, 180, 0.55);
        
                        border-radius: 9px;
        
                        background:
                            rgba(17, 30, 59, 0.82);
        
                        color: #F1F4FF;
        
                        font-family: inherit;
                        font-size: 12px;
                        font-weight: 650;
        
                        cursor: pointer;
        
                        transition:
                            background 0.15s ease,
                            border-color 0.15s ease,
                            transform 0.15s ease;
                    }}
        
                    .reference-view-button:hover {{
                        background:
                            rgba(38, 51, 80, 0.98);
        
                        border-color:
                            rgba(115, 135, 200, 0.85);
        
                        transform: translateY(-1px);
                    }}
        
                    /* ====================================================
                       DOWNLOAD BUTTON
                       ==================================================== */
        
                    .reference-download-button {{
                        flex-shrink: 0;
        
                        width: 34px;
                        height: 34px;
        
                        margin-left: 6px;
        
                        display: inline-flex;
                        align-items: center;
                        justify-content: center;
        
                        padding: 0;
        
                        border:
                            1px solid
                            rgba(89, 112, 180, 0.55);
        
                        border-radius: 9px;
        
                        background:
                            rgba(17, 30, 59, 0.82);
        
                        color: #F1F4FF;
        
                        font-family: inherit;
                        font-size: 18px;
                        font-weight: 500;
        
                        line-height: 1;
        
                        cursor: pointer;
        
                        transition:
                            background 0.15s ease,
                            border-color 0.15s ease,
                            transform 0.15s ease;
                    }}
        
                    .reference-download-button:hover {{
                        background:
                            rgba(38, 51, 80, 0.98);
        
                        border-color:
                            rgba(115, 135, 200, 0.85);
        
                        transform: translateY(-1px);
                    }}
        
                    @media (max-width: 700px) {{
        
                        .reference-document-name {{
                            max-width: 150px;
                        }}
        
                    }}
        
                </style>
        
            </head>
        
            <body>
        
                <div class="reference-documents-panel">
        
                    <div class="reference-header">
        
                        <span class="reference-folder-icon">
                            📁
                        </span>
        
                        <span class="reference-panel-title">
                            Reference documents
                        </span>
        
                        <span class="reference-panel-count">
                            {len(pdf_references)}
                        </span>
        
                    </div>
        
                    <div class="reference-document-list">
        
                        {reference_rows}
        
                    </div>
        
                </div>
        
        
                <script>
        
                    // ====================================================
                    // BASE64 → BLOB
                    // ====================================================
        
                    function base64ToBlob(
                        base64,
                        contentType
                    ) {{
        
                        const byteCharacters =
                            atob(base64);
        
                        const byteArrays = [];
        
                        const sliceSize = 1024;
        
                        for (
                            let offset = 0;
                            offset < byteCharacters.length;
                            offset += sliceSize
                        ) {{
        
                            const slice =
                                byteCharacters.slice(
                                    offset,
                                    offset + sliceSize
                                );
        
                            const byteNumbers =
                                new Array(slice.length);
        
                            for (
                                let i = 0;
                                i < slice.length;
                                i++
                            ) {{
        
                                byteNumbers[i] =
                                    slice.charCodeAt(i);
        
                            }}
        
                            const byteArray =
                                new Uint8Array(
                                    byteNumbers
                                );
        
                            byteArrays.push(
                                byteArray
                            );
        
                        }}
        
                        return new Blob(
                            byteArrays,
                            {{
                                type: contentType
                            }}
                        );
                    }}
        
        
                    // ====================================================
                    // VIEW PDF
                    // OPENS PDF IN NEW BROWSER TAB
                    // ====================================================
        
                    function viewPDF(
                        base64
                    ) {{
        
                        try {{
        
                            const blob =
                                base64ToBlob(
                                    base64,
                                    "application/pdf"
                                );
        
                            const pdfURL =
                                URL.createObjectURL(
                                    blob
                                );
        
                            window.open(
                                pdfURL,
                                "_blank"
                            );
        
                        }} catch (error) {{
        
                            console.error(
                                "Unable to open PDF:",
                                error
                            );
        
                        }}
        
                    }}
        
        
                    // ====================================================
                    // DOWNLOAD PDF
                    // ====================================================
        
                    function downloadPDF(
                        base64,
                        filename
                    ) {{
        
                        try {{
        
                            const blob =
                                base64ToBlob(
                                    base64,
                                    "application/pdf"
                                );
        
                            const pdfURL =
                                URL.createObjectURL(
                                    blob
                                );
        
                            const link =
                                document.createElement(
                                    "a"
                                );
        
                            link.href = pdfURL;
                            link.download = filename;
        
                            document.body.appendChild(
                                link
                            );
        
                            link.click();
        
                            document.body.removeChild(
                                link
                            );
        
                            setTimeout(
                                function() {{
                                    URL.revokeObjectURL(
                                        pdfURL
                                    );
                                }},
                                1000
                            );
        
                        }} catch (error) {{
        
                            console.error(
                                "Unable to download PDF:",
                                error
                            );
        
                        }}
        
                    }}
        
                </script>
        
            </body>
        
            </html>
            """
        
            # ============================================================
            # RENDER REFERENCE PANEL
            # ============================================================
        
            components.html(
                reference_html,
                height=(
                    48
                    + (len(pdf_references) * 62)
                    + 14
                ),
                scrolling=False
            )


# ============================================================
# OUTER ERROR HANDLER
# ============================================================

    except Exception as error:

        st.error(
            "Unable to process the question."
            )
        
        st.exception(error)

# ============================================================
# BOTTOM AI DISCLAIMER
# ============================================================

render_html("""
<div class="disclaimer">


    <div class="disclaimer-content">

        <div class="disclaimer-main">
            All answers are based only on official municipal documents.
        </div>

     
        <div class="disclaimer-ai">
            This is AI-generated information.
            AI may make mistakes.
            Please verify important information
            before using it.
        </div>

    </div>

</div>
""")
