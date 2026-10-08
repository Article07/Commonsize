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


_LOGIN_CSS = f"""
<style>
.stApp {{ background: radial-gradient(1200px 600px at 85% -10%, #FEE2E2 0%, rgba(254,226,226,0) 60%),
                       linear-gradient(180deg, #F8FAFC 0%, #EEF2F7 100%); }}
.block-container {{ max-width: 1100px; padding-top: 7vh !important; }}

.ac-lp {{ position: relative; overflow: hidden; background: {NAVY}; color: #fff; border-radius: 14px;
          padding: 48px 44px; min-height: 460px; border-left: 6px solid {RED};
          box-shadow: 0 24px 48px -24px rgba(15, 23, 42, .55); }}
.ac-lp::before {{ content: ""; position: absolute; width: 340px; height: 340px; right: -120px; top: -140px;
                  border-radius: 50%; background: radial-gradient(circle, rgba(239,68,68,.35), rgba(239,68,68,0) 70%); }}
.ac-lp::after {{ content: ""; position: absolute; width: 260px; height: 260px; left: -90px; bottom: -130px;
                 border-radius: 50%; border: 1px solid rgba(255,255,255,.08); }}
.ac-lp .kicker {{ color: #FCA5A5; font-size: 12px; letter-spacing: .16em; text-transform: uppercase; font-weight: 500; }}
.ac-lp h1 {{ color: #fff !important; font-weight: 400; font-size: 38px; line-height: 1.2; margin: 14px 0 14px; padding: 0; }}
.ac-lp h1 b {{ color: {RED}; font-weight: 500; }}
.ac-lp p.lead {{ color: #CBD5E1; font-size: 16px; line-height: 1.65; margin: 0 0 30px; max-width: 440px; }}
.ac-lp ul {{ list-style: none; padding: 0; margin: 0; display: grid; gap: 16px; }}
.ac-lp li {{ display: flex; gap: 14px; align-items: flex-start; color: #E2E8F0; font-size: 14.5px; line-height: 1.5; margin: 0; }}
.ac-lp li .dot {{ flex: 0 0 30px; height: 30px; border-radius: 8px; background: rgba(239,68,68,.16); color: #FCA5A5;
                  display: flex; align-items: center; justify-content: center; font-weight: 500; font-size: 13px; }}
.ac-lp li b {{ color: #fff; font-weight: 500; }}

.st-key-login_card {{ background: #fff; border-radius: 14px; padding: 40px 36px 30px; min-height: 460px;
                      border: 1px solid #E5E7EB; box-shadow: 0 24px 48px -28px rgba(15, 23, 42, .35); }}
.ac-card-head img {{ height: 52px; margin-bottom: 26px; }}
.ac-card-head h2 {{ color: {NAVY}; font-weight: 500; font-size: 26px; margin: 0 0 6px; padding: 0; }}
.ac-card-head p {{ color: {GREY}; font-size: 14px; margin: 0 0 22px; }}
.st-key-login_card [data-testid="stForm"] {{ border: 0; padding: 0; }}
.st-key-login_card .st-key-access_code input {{ height: 46px; font-size: 16px; }}
.st-key-login_card div.stFormSubmitButton > button {{ background: {RED}; color: #fff; border: 0; height: 46px;
    font-weight: 500; font-size: 15px; border-radius: 6px; margin-top: 6px; }}
.st-key-login_card div.stFormSubmitButton > button:hover {{ background: #DC2626; color: #fff; }}
.ac-card-help {{ color: {GREY}; font-size: 12.5px; margin-top: 18px; padding-top: 16px; border-top: 1px solid #F1F5F9; }}
.ac-card-help b {{ color: {NAVY}; font-weight: 500; }}
.ac-showpw {{ display: inline-flex; align-items: center; gap: 8px; color: {GREY}; font-size: 13.5px; cursor: pointer;
              user-select: none; margin: 2px 0 4px; }}
.ac-showpw input {{ width: 16px; height: 16px; accent-color: {RED}; cursor: pointer; margin: 0; }}
.stApp:has(#ac-showpw:checked) .st-key-access_code input {{ -webkit-text-security: none; text-security: none;
                                                             letter-spacing: normal; }}
@media (max-width: 700px) {{
  .ac-lp {{ display: none; }}
  .block-container {{ padding-top: 3vh !important; }}
  .st-key-login_card {{ min-height: 0; padding: 30px 22px 22px; }}
}}
</style>
"""


def login_styles():
    st.markdown(_LOGIN_CSS, unsafe_allow_html=True)


def login_panel():
    items = [
        ("PDF or Excel", "digital or scanned financials, one year or several at once"),
        ("Classified", "every line placed under the firm's taxonomy; new items highlighted yellow"),
        ("Checked", "the balance sheet must tally and the income statement must match the input"),
    ]
    rows = "".join(f'<li><span class="dot">{i}</span><span><b>{html.escape(t)}</b> &mdash; {html.escape(d)}</span></li>'
                   for i, (t, d) in enumerate(items, 1))
    st.markdown(
        '<div class="ac-lp"><div class="kicker">Acumen M&amp;A Advisors &nbsp;|&nbsp; Internal tool</div>'
        '<h1>Common Size <b>Generator</b></h1>'
        '<p class="lead">From a company\'s financial statements to the firm\'s Common Size workbook, '
        'reconciled and ready for analysis.</p>'
        f'<ul>{rows}</ul></div>',
        unsafe_allow_html=True)


def login_card_head():
    st.markdown(
        f'<div class="ac-card-head"><img src="{LOGO_URL}" alt="Acumen M&amp;A Advisors">'
        '<h2>Sign in</h2><p>Enter the team password to continue.</p></div>',
        unsafe_allow_html=True)


def show_password_toggle():
    """A plain HTML tick box; CSS (:has) unmasks the password box while it is ticked. No rerun, nothing sent."""
    st.markdown('<label class="ac-showpw" for="ac-showpw"><input type="checkbox" id="ac-showpw"> Show password</label>',
                unsafe_allow_html=True)


def login_card_help():
    st.markdown('<div class="ac-card-help"><b>Need access?</b> Ask the Commonsize administrator for the team '
                'password.</div>', unsafe_allow_html=True)


def footer():
    st.markdown('<div class="ac-foot"><span>Acumen M&amp;A Advisors LLP. Confidential: client financials are processed '
                'on this server and are not stored.</span></div>',
                unsafe_allow_html=True)
