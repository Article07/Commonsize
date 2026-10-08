"""
Look and feel, following the firm's website (acumenfc.co.in): white header with the
Acumen logo, red accent (#EF4444), slate-navy panels (#1E293B), Roboto, red section
headings. Pure presentation; no logic lives here.
"""

import html

import streamlit as st

RED = "#EF4444"
NAVY = "#1E293B"
SLATE = "#334155"
GREY = "#4B5563"
LOGO_URL = "https://documents.acumenfc.co.in/Acumen-logo-with-R-min.png"

_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
html, body, .stApp, .stApp p, .stApp label, .stApp li, .stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp td, .stApp th,
button, input, textarea {{ font-family: 'Roboto', sans-serif; }}
[data-testid="stIconMaterial"], [class*="material-symbols"], .material-icons {{
    font-family: 'Material Symbols Rounded', 'Material Icons' !important; }}
#MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"] {{ display: none !important; }}
header[data-testid="stHeader"] {{ background: transparent; height: 0; }}
.block-container {{ max-width: 1180px; padding-top: 0 !important; padding-bottom: 3rem; }}

.stApp {{ border-top: 4px solid {RED}; }}
.ac-top {{ display: none; }}
html, body {{ overflow-x: hidden; }}
.ac-nav {{ display: flex; align-items: center; justify-content: space-between; padding: 16px 0 14px;
          border-bottom: 1px solid #E5E5E5; margin-bottom: 22px; }}
.ac-nav img {{ height: 46px; }}
.ac-nav .tag {{ color: {GREY}; font-size: 14px; letter-spacing: .04em; text-transform: uppercase; }}
.ac-nav .tag b {{ color: {RED}; font-weight: 500; }}

.ac-hero {{ background: {NAVY}; color: #fff; border-radius: 6px; padding: 34px 40px; margin-bottom: 8px;
           border-left: 6px solid {RED}; }}
.ac-hero h1 {{ color: #fff !important; font-weight: 400; font-size: 34px; line-height: 1.25; margin: 0 0 10px; padding: 0; }}
.ac-hero p {{ color: #CBD5E1; font-size: 16px; margin: 0; max-width: 760px; line-height: 1.6; }}

.ac-step {{ display: flex; align-items: center; gap: 14px; margin: 38px 0 6px; }}
.ac-step .n {{ background: {RED}; color: #fff; min-width: 34px; height: 34px; border-radius: 50%;
              display: flex; align-items: center; justify-content: center; font-weight: 500; font-size: 16px; }}
.ac-step h3 {{ color: {RED}; font-weight: 400; font-size: 26px; margin: 0; padding: 0; }}
.ac-hint {{ color: {GREY}; font-size: 14px; margin: 0 0 10px 48px; }}

.ac-cards {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 14px; margin: 10px 0 16px; }}
.ac-card {{ background: {NAVY}; color: #fff; border-radius: 6px; padding: 18px 20px; border-top: 4px solid {SLATE}; }}
.ac-card .k {{ color: #94A3B8; font-size: 12px; text-transform: uppercase; letter-spacing: .06em; }}
.ac-card .v {{ font-size: 22px; margin-top: 6px; font-weight: 400; }}
.ac-card.ok {{ border-top-color: #22C55E; }}
.ac-card.bad {{ border-top-color: {RED}; }}
.ac-card.bad .v {{ color: #FCA5A5; }}
.ac-card.ok .v {{ color: #86EFAC; }}

.ac-foot {{ margin-top: 56px; padding-top: 16px; border-top: 1px solid #E5E5E5; color: {GREY}; font-size: 13px;
           display: flex; justify-content: space-between; flex-wrap: wrap; gap: 8px; }}
.st-key-access_code input {{ -webkit-text-security: disc; text-security: disc; letter-spacing: .12em; }}
.ac-login {{ max-width: 460px; margin: 40px auto 0; text-align: center; }}
.ac-login img {{ height: 64px; margin-bottom: 18px; }}
.ac-login h2 {{ color: {NAVY}; font-weight: 400; margin: 0 0 4px; }}
.ac-login p {{ color: {GREY}; font-size: 14px; margin-bottom: 18px; }}

[data-testid="stDataFrame"], [data-testid="stDataEditor"] {{ border: 1px solid #E5E5E5; border-radius: 6px; }}
[data-testid="stAlert"] {{ border-radius: 6px; }}
.stRadio label p, .stSelectbox label p, .stTextInput label p, .stMultiSelect label p, .stFileUploader label p {{
    color: {NAVY}; font-weight: 500; }}
[data-testid="stFileUploaderDropzone"] {{
    border: 2px dashed {RED}88; background: #FEF2F2; border-radius: 8px; min-height: 170px; padding: 22px;
    display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 12px; text-align: center;
    transition: background .15s, border-color .15s; }}
[data-testid="stFileUploaderDropzone"]:hover, [data-testid="stFileUploaderDropzone"]:focus-within {{
    background: #FEE2E2; border-color: {RED}; }}
[data-testid="stFileUploaderDropzoneInstructions"] {{ display: flex; flex-direction: column; align-items: center; }}
[data-testid="stFileUploaderDropzone"]:not(:has([data-testid="stFileUploaderDropzoneInstructions"]))::after {{
    content: "Drop more files anywhere in this box"; color: {GREY}; font-size: 13px; }}
[data-testid="stFileUploaderDropzoneInstructions"]::before {{
    content: "Drag and drop your files anywhere in this box"; display: block; color: {NAVY}; font-weight: 500;
    font-size: 16px; margin-bottom: 4px; }}
div.stButton > button[kind="primary"], div.stDownloadButton > button {{
    background: {RED}; border: 0; color: #fff; font-weight: 500; padding: .55rem 1.4rem; border-radius: 4px; }}
div.stButton > button[kind="primary"]:hover, div.stDownloadButton > button:hover {{ background: #DC2626; color: #fff; }}
</style>
"""


def apply_theme():
    st.markdown(_CSS, unsafe_allow_html=True)


def header(subtitle="Common Size Generator"):
    st.markdown(
        f'<div class="ac-top"></div><div class="ac-nav"><img src="{LOGO_URL}" alt="Acumen M&amp;A Advisors">'
        f'<div class="tag"><b>{html.escape(subtitle)}</b> &nbsp;|&nbsp; Internal tool</div></div>',
        unsafe_allow_html=True)


def hero(title, text):
    st.markdown(f'<div class="ac-hero"><h1>{html.escape(title)}</h1><p>{html.escape(text)}</p></div>',
                unsafe_allow_html=True)


def step(number, title, hint=""):
    hint_html = f'<div class="ac-hint">{html.escape(hint)}</div>' if hint else ""
    st.markdown(f'<div class="ac-step"><div class="n">{number}</div><h3>{html.escape(title)}</h3></div>{hint_html}',
                unsafe_allow_html=True)


def cards(items):
    """items: [(label, value, state)] with state 'ok' | 'bad' | ''."""
    body = "".join(
        f'<div class="ac-card {state}"><div class="k">{html.escape(label)}</div><div class="v">{html.escape(str(value))}</div></div>'
        for label, value, state in items)
    st.markdown(f'<div class="ac-cards">{body}</div>', unsafe_allow_html=True)


def login_header():
    st.markdown(
        f'<div class="ac-top"></div><div class="ac-login"><img src="{LOGO_URL}" alt="Acumen M&amp;A Advisors">'
        '<h2>Common Size Generator</h2><p>For Acumen team members. Enter the shared password to continue.</p></div>',
        unsafe_allow_html=True)


def footer():
    st.markdown('<div class="ac-foot"><span>Acumen M&amp;A Advisors LLP. Confidential: client financials are processed '
                'on this server and are not stored.</span><span>Rule-based; no AI service is used.</span></div>',
                unsafe_allow_html=True)
