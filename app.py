import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
import re
import difflib

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="GST Reconciliation Pro",
    page_icon="🧾",
    layout="wide"
)


# =========================================================
# LOGIN / AUTHENTICATION
# =========================================================
# Login credentials requested for the current GST Reconciliation Pro build.
LOGIN_ID = "9560838810"
LOGIN_PASSWORD = "nishant@123"

if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    # =====================================================
    # CODE-ONLY LOGIN CHARACTER
    # No PNG / no extra folder / no external asset required.
    # Pure inline SVG + CSS animation.
    # =====================================================
    login_visual = """
    <div class="login-visual-scene">
        <svg class="login-character-svg"
             viewBox="0 0 620 620"
             role="img"
             aria-label="Animated GST Reconciliation Pro business character"
             xmlns="http://www.w3.org/2000/svg">

            <defs>
                <linearGradient id="sceneBg" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0%" stop-color="#f8fbff"/>
                    <stop offset="58%" stop-color="#eef5ff"/>
                    <stop offset="100%" stop-color="#ffffff"/>
                </linearGradient>
                <linearGradient id="suitGrad" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0%" stop-color="#263b73"/>
                    <stop offset="100%" stop-color="#101c45"/>
                </linearGradient>
                <linearGradient id="shirtGrad" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stop-color="#ffffff"/>
                    <stop offset="100%" stop-color="#e8eef8"/>
                </linearGradient>
                <linearGradient id="skinGrad" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0%" stop-color="#ffd5b5"/>
                    <stop offset="100%" stop-color="#e7a579"/>
                </linearGradient>
                <linearGradient id="briefGrad" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0%" stop-color="#26324c"/>
                    <stop offset="100%" stop-color="#0e1629"/>
                </linearGradient>
                <filter id="softShadow" x="-40%" y="-40%" width="180%" height="180%">
                    <feDropShadow dx="0" dy="12" stdDeviation="12" flood-color="#183b70" flood-opacity=".16"/>
                </filter>
                <filter id="smallShadow" x="-40%" y="-40%" width="180%" height="180%">
                    <feDropShadow dx="0" dy="6" stdDeviation="6" flood-color="#183b70" flood-opacity=".13"/>
                </filter>
            </defs>

            <!-- Soft corporate background -->
            <rect x="0" y="0" width="620" height="620" rx="30" fill="url(#sceneBg)"/>
            <circle cx="70" cy="90" r="120" fill="#dcecff" opacity=".42"/>
            <circle cx="555" cy="95" r="105" fill="#e6f1ff" opacity=".70"/>
            <path d="M0 500 C150 430 245 530 350 465 C455 400 535 455 620 405 L620 620 L0 620Z"
                  fill="#eaf3ff" opacity=".8"/>

            <!-- Floating GST card -->
            <g class="float-card card-one" filter="url(#smallShadow)">
                <rect x="45" y="72" width="132" height="82" rx="16" fill="#fff" stroke="#d8e6f7"/>
                <rect x="60" y="88" width="42" height="42" rx="10" fill="#e9f2ff"/>
                <text x="81" y="115" text-anchor="middle" font-size="18" font-weight="900" fill="#1769aa">GST</text>
                <rect x="113" y="92" width="45" height="7" rx="3.5" fill="#dbe7f5"/>
                <rect x="113" y="107" width="34" height="7" rx="3.5" fill="#e7eef7"/>
                <rect x="113" y="122" width="42" height="7" rx="3.5" fill="#edf2f8"/>
            </g>

            <!-- Excel card -->
            <g class="float-card card-two" filter="url(#smallShadow)">
                <rect x="33" y="222" width="118" height="72" rx="16" fill="#fff" stroke="#d8e6f7"/>
                <rect x="49" y="237" width="43" height="43" rx="9" fill="#e8f8ee"/>
                <text x="70.5" y="266" text-anchor="middle" font-size="25" font-weight="900" fill="#1f9d55">X</text>
                <rect x="103" y="242" width="33" height="6" rx="3" fill="#dbe7f5"/>
                <rect x="103" y="256" width="25" height="6" rx="3" fill="#e7eef7"/>
                <rect x="103" y="270" width="30" height="6" rx="3" fill="#edf2f8"/>
            </g>

            <!-- Analytics card -->
            <g class="float-card card-three" filter="url(#smallShadow)">
                <rect x="360" y="80" width="154" height="100" rx="18" fill="#fff" stroke="#d8e6f7"/>
                <text x="380" y="108" font-size="12" font-weight="800" fill="#64748b">ITC ANALYTICS</text>
                <rect x="382" y="142" width="12" height="20" rx="3" fill="#9cc6ff"/>
                <rect x="403" y="132" width="12" height="30" rx="3" fill="#6faeff"/>
                <rect x="424" y="119" width="12" height="43" rx="3" fill="#3f8df5"/>
                <path d="M380 135 L405 126 L426 112 L454 122 L482 100"
                      fill="none" stroke="#1769aa" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
                <path d="M474 101 L484 99 L481 110" fill="none" stroke="#1769aa" stroke-width="3" stroke-linecap="round"/>
            </g>

            <!-- Invoice card -->
            <g class="float-card card-four" filter="url(#smallShadow)">
                <rect x="405" y="218" width="166" height="91" rx="18" fill="#fff" stroke="#d8e6f7"/>
                <circle cx="432" cy="249" r="16" fill="#eaf3ff"/>
                <text x="432" y="255" text-anchor="middle" font-size="15" font-weight="900" fill="#1769aa">₹</text>
                <text x="458" y="246" font-size="12" font-weight="900" fill="#172033">INVOICE</text>
                <rect x="458" y="258" width="83" height="6" rx="3" fill="#dbe7f5"/>
                <rect x="458" y="271" width="66" height="6" rx="3" fill="#e7eef7"/>
                <rect x="458" y="284" width="76" height="6" rx="3" fill="#edf2f8"/>
            </g>

            <!-- Reconciliation checks -->
            <g class="float-card card-five" filter="url(#smallShadow)">
                <rect x="410" y="348" width="154" height="108" rx="18" fill="#fff" stroke="#d8e6f7"/>
                <text x="429" y="373" font-size="11" font-weight="900" fill="#64748b">RECONCILE</text>
                <circle cx="431" cy="394" r="9" fill="#dcfce7"/>
                <path d="M426 394 l4 4 l7 -8" fill="none" stroke="#16a34a" stroke-width="2.5"/>
                <text x="447" y="398" font-size="11" fill="#334155">2B Portal</text>
                <circle cx="431" cy="419" r="9" fill="#dcfce7"/>
                <path d="M426 419 l4 4 l7 -8" fill="none" stroke="#16a34a" stroke-width="2.5"/>
                <text x="447" y="423" font-size="11" fill="#334155">Books</text>
                <circle cx="431" cy="444" r="9" fill="#dcfce7"/>
                <path d="M426 444 l4 4 l7 -8" fill="none" stroke="#16a34a" stroke-width="2.5"/>
                <text x="447" y="448" font-size="11" fill="#334155">Matched</text>
            </g>

            <!-- Desk / floor -->
            <ellipse cx="265" cy="543" rx="198" ry="30" fill="#cfe0f4" opacity=".55"/>
            <rect x="120" y="496" width="305" height="17" rx="8.5" fill="#dce9f7"/>
            <rect x="143" y="512" width="13" height="55" rx="6" fill="#c6d8ed"/>
            <rect x="389" y="512" width="13" height="55" rx="6" fill="#c6d8ed"/>

            <!-- Books -->
            <g class="books" filter="url(#smallShadow)">
                <rect x="88" y="450" width="105" height="18" rx="5" fill="#24477c"/>
                <rect x="94" y="432" width="105" height="18" rx="5" fill="#3569a8"/>
                <rect x="100" y="414" width="105" height="18" rx="5" fill="#173b70"/>
                <text x="151" y="427" text-anchor="middle" font-size="8" font-weight="900" fill="#fff">COMPLIANCE</text>
                <text x="146" y="445" text-anchor="middle" font-size="8" font-weight="900" fill="#fff">TAX</text>
                <text x="140" y="463" text-anchor="middle" font-size="8" font-weight="900" fill="#fff">GST</text>
            </g>

            <!-- Character shadow -->
            <ellipse cx="275" cy="505" rx="78" ry="15" fill="#9eb6d5" opacity=".35"/>

            <!-- Animated character -->
            <g class="character">
                <!-- back arm -->
                <g class="char-arm-back">
                    <path d="M225 333 C205 350 195 374 183 398"
                          fill="none" stroke="#172b59" stroke-width="24" stroke-linecap="round"/>
                    <circle cx="181" cy="400" r="12" fill="url(#skinGrad)"/>
                </g>

                <!-- legs -->
                <g class="char-leg-a">
                    <path d="M250 425 L238 482" fill="none" stroke="#18264b" stroke-width="28" stroke-linecap="round"/>
                    <path d="M235 485 C224 486 216 493 213 502 C212 508 218 511 229 511 L257 511 C261 503 253 492 235 485Z"
                          fill="#fff"/>
                </g>
                <g class="char-leg-b">
                    <path d="M292 425 L313 478" fill="none" stroke="#1d2e58" stroke-width="28" stroke-linecap="round"/>
                    <path d="M308 478 C318 480 327 487 330 496 C332 503 326 508 315 508 L287 508 C284 500 291 488 308 478Z"
                          fill="#fff"/>
                </g>

                <!-- body -->
                <g class="char-body">
                    <path d="M222 318 C235 303 274 299 294 318 L307 420
                             C292 434 246 434 224 420Z"
                          fill="url(#suitGrad)" filter="url(#smallShadow)"/>
                    <path d="M252 316 L274 316 L280 350 L264 372 L247 350Z"
                          fill="url(#shirtGrad)"/>
                    <path d="M260 325 L273 325 L276 371 L267 385 L258 371Z"
                          fill="#26304d"/>
                    <path d="M240 338 L224 353 L232 399" fill="none" stroke="#334c83" stroke-width="6" opacity=".7"/>
                    <path d="M286 337 L302 352 L296 399" fill="none" stroke="#334c83" stroke-width="6" opacity=".7"/>
                </g>

                <!-- presenting arm -->
                <g class="char-arm-front">
                    <path d="M294 334 C314 347 325 361 346 366"
                          fill="none" stroke="#172b59" stroke-width="25" stroke-linecap="round"/>
                    <path d="M345 366 C358 368 369 362 380 354"
                          fill="none" stroke="url(#skinGrad)" stroke-width="13" stroke-linecap="round"/>
                    <circle cx="382" cy="353" r="11" fill="url(#skinGrad)"/>
                </g>

                <!-- neck -->
                <path d="M252 299 L252 320 Q264 331 277 320 L277 299Z" fill="url(#skinGrad)"/>

                <!-- head -->
                <g class="char-head">
                    <ellipse cx="265" cy="262" rx="48" ry="54" fill="url(#skinGrad)" filter="url(#smallShadow)"/>
                    <!-- ears -->
                    <circle cx="218" cy="270" r="10" fill="#efb18a"/>
                    <circle cx="312" cy="270" r="10" fill="#efb18a"/>
                    <!-- hair -->
                    <path d="M220 249 C217 220 236 197 267 198
                             C294 197 311 215 310 245
                             C300 232 290 226 278 226
                             C267 216 251 218 240 229
                             C233 235 227 243 220 249Z"
                          fill="#2a2530"/>
                    <path d="M228 235 C241 214 258 207 275 210"
                          fill="none" stroke="#493746" stroke-width="7" stroke-linecap="round" opacity=".65"/>
                    <!-- face -->
                    <ellipse cx="247" cy="263" rx="4" ry="6" fill="#342b31"/>
                    <ellipse cx="284" cy="263" rx="4" ry="6" fill="#342b31"/>
                    <path d="M257 281 Q266 288 275 281" fill="none" stroke="#a34f4c" stroke-width="3" stroke-linecap="round"/>
                    <path d="M258 273 Q265 277 272 273" fill="none" stroke="#c87f68" stroke-width="2" stroke-linecap="round"/>
                    <circle cx="232" cy="279" r="7" fill="#f29b91" opacity=".24"/>
                    <circle cx="297" cy="279" r="7" fill="#f29b91" opacity=".24"/>
                </g>

                <!-- briefcase -->
                <g class="char-briefcase">
                    <rect x="175" y="370" width="65" height="49" rx="9" fill="url(#briefGrad)" filter="url(#smallShadow)"/>
                    <path d="M195 370 V360 Q195 353 202 353 H214 Q221 353 221 360 V370"
                          fill="none" stroke="#46526e" stroke-width="5"/>
                    <rect x="175" y="390" width="65" height="5" fill="#3b4965"/>
                    <circle cx="207" cy="393" r="3" fill="#c9d4e4"/>
                </g>
            </g>

            <!-- Plant -->
            <g class="plant">
                <path d="M505 500 C500 463 508 440 526 423" fill="none" stroke="#4a8b5b" stroke-width="5" stroke-linecap="round"/>
                <path d="M513 468 C492 456 483 440 489 426 C507 429 518 443 513 468Z" fill="#79b47d"/>
                <path d="M516 454 C530 437 546 433 559 441 C549 457 533 464 516 454Z" fill="#5f9f69"/>
                <path d="M504 492 C486 482 478 469 481 458 C498 460 508 473 504 492Z" fill="#8ac38b"/>
                <path d="M498 495 H542 L535 526 Q520 536 505 526Z" fill="#f8fafc" stroke="#d7e1ec"/>
                <path d="M501 501 H539" stroke="#d2dce8" stroke-width="3"/>
            </g>

            <!-- Floating sparkles -->
            <g class="sparkles" fill="#6ea8ef">
                <circle cx="190" cy="150" r="4"/>
                <circle cx="332" cy="205" r="3"/>
                <circle cx="570" cy="168" r="4"/>
                <circle cx="380" cy="320" r="3"/>
            </g>
        </svg>
    </div>
    """


    st.markdown('<div class="login-page"><div class="login-shell">', unsafe_allow_html=True)

    left_col, right_col = st.columns([1.08, 0.92], gap="small")

    with left_col:
        st.markdown(
            f'<div class="login-image-wrap">{login_visual}</div>',
            unsafe_allow_html=True
        )

    with right_col:
        st.markdown('<div class="login-form-wrap">', unsafe_allow_html=True)
        st.markdown('<div class="login-brand">GST • TAX • COMPLIANCE</div>', unsafe_allow_html=True)
        st.markdown('<h1 class="login-title">GST Reconciliation Pro</h1>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle">Match&nbsp;&nbsp;|&nbsp;&nbsp;Verify&nbsp;&nbsp;|&nbsp;&nbsp;Reconcile&nbsp;&nbsp;|&nbsp;&nbsp;Stay Compliant</div>', unsafe_allow_html=True)
        st.markdown('<h2 style="text-align:center;color:#172033;font-size:24px;margin:0 0 4px;">Welcome Back</h2>', unsafe_allow_html=True)
        st.markdown('<div class="login-subtitle" style="margin-bottom:22px;">Login to continue to your reconciliation dashboard</div>', unsafe_allow_html=True)

        with st.form("gst_login_form", clear_on_submit=False):
            login_id_input = st.text_input(
                "Login ID",
                placeholder="Enter your login ID",
                key="gst_login_id"
            )
            password_input = st.text_input(
                "Password",
                type="password",
                placeholder="Enter your password",
                key="gst_login_password"
            )
            remember_me = st.checkbox("Remember me", value=True, key="gst_remember_me")
            submitted = st.form_submit_button(
                "🔐  SUBMIT",
                type="primary",
                use_container_width=True
            )

        if submitted:
            if login_id_input.strip() == LOGIN_ID and password_input == LOGIN_PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Invalid Login ID or Password. Please try again.")

        st.markdown('<div class="login-locked-note">✓ Secure access • Professional GST workspace</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="login-security">🔒 Your reconciliation workspace is protected by login authentication.</div>', unsafe_allow_html=True)
        st.markdown('<div class="login-small">GST Reconciliation Pro • Smart GST 2B vs Books Analysis</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('</div></div>', unsafe_allow_html=True)
    st.stop()


# =========================================================
# CSS
# =========================================================

st.markdown("""
<style>
.stApp {
    background-color: #f4f7fb;
}

[data-testid="stHeader"] {
    background-color: transparent;
}

.hero {
    background: linear-gradient(135deg, #0f2a5f, #1769aa);
    padding: 30px;
    border-radius: 20px;
    margin-bottom: 25px;
}

.hero h1 {
    color: white !important;
    font-size: 34px;
    margin: 0;
}

.hero p {
    color: #dbeafe !important;
    margin-top: 8px;
}

.section-title {
    color: #0f2a5f;
    font-size: 23px;
    font-weight: 800;
    margin-top: 25px;
    margin-bottom: 15px;
}

.kpi {
    background: white;
    border-radius: 16px;
    padding: 20px;
    border: 1px solid #dbe3ef;
    box-shadow: 0 5px 18px rgba(15,42,95,.07);
    min-height: 125px;
}

.kpi-title {
    color: #64748b;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: .03em;
}

.kpi-value {
    color: #172033;
    font-size: 30px;
    font-weight: 900;
    margin-top: 8px;
}

.kpi-desc {
    color: #94a3b8;
    font-size: 12px;
}

.blue {
    border-top: 4px solid #2563eb;
}

.green {
    border-top: 4px solid #16a34a;
}

.red {
    border-top: 4px solid #dc2626;
}

.orange {
    border-top: 4px solid #ea580c;
}

.purple {
    border-top: 4px solid #7c3aed;
}

.teal {
    border-top: 4px solid #0d9488;
}

.info-box {
    background: white;
    padding: 18px;
    border-radius: 13px;
    border-left: 5px solid #2563eb;
    border-top: 1px solid #dbe3ef;
    border-right: 1px solid #dbe3ef;
    border-bottom: 1px solid #dbe3ef;
    margin: 15px 0;
}

.warn-box {
    background: #fff7ed;
    padding: 14px 18px;
    border-radius: 13px;
    border-left: 5px solid #ea580c;
    margin: 10px 0;
    color: #7c2d12;
    font-size: 13px;
}

[data-testid="stFileUploaderDropzone"] {
    background: white;
    border: 2px dashed #b8c9df;
    border-radius: 14px;
}

.stButton button {
    border-radius: 10px;
    font-weight: 700;
}

.stDownloadButton button {
    background: #0f2a5f !important;
    color: white !important;
    border-radius: 10px;
    font-weight: 700;
}

.badge {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 11px;
    font-weight: 700;
}

/* =========================================================
   YOOM-STYLE ANIMATIONS
   ========================================================= */

@keyframes floatUpDown {
    0%, 100% { transform: translateY(0px); }
    50% { transform: translateY(-14px); }
}
@keyframes fadeInUp {
    from { opacity: 0; transform: translateY(18px); }
    to { opacity: 1; transform: translateY(0); }
}
@keyframes gradientShift {
    0% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
@keyframes blinkCursor {
    0%, 50% { opacity: 1; }
    51%, 100% { opacity: 0; }
}

.hero {
    background: linear-gradient(-45deg, #0f2a5f, #1769aa, #2563eb, #0f2a5f);
    background-size: 300% 300%;
    animation: gradientShift 8s ease infinite;
    position: relative;
    overflow: hidden;
}
.hero::before, .hero::after {
    content: "";
    position: absolute;
    border-radius: 50%;
    background: rgba(255,255,255,0.08);
}
.hero::before {
    width: 180px; height: 180px;
    top: -60px; right: 40px;
    animation: floatUpDown 6s ease-in-out infinite;
}
.hero::after {
    width: 110px; height: 110px;
    bottom: -30px; left: 60px;
    animation: floatUpDown 5s ease-in-out infinite reverse;
}
.hero h1 { animation: fadeInUp 0.8s ease; }
.hero p { animation: fadeInUp 0.8s ease 0.2s backwards; }
.hero-cursor {
    display: inline-block;
    width: 3px; height: 1em;
    background: #fff;
    margin-left: 5px;
    animation: blinkCursor 1s step-start infinite;
    vertical-align: middle;
}

.float-icons {
    display: flex;
    gap: 18px;
    justify-content: center;
    margin: 22px 0 30px;
    flex-wrap: wrap;
}
.float-icon {
    width: 56px; height: 56px;
    border-radius: 50%;
    background: white;
    display: flex; align-items: center; justify-content: center;
    font-size: 24px;
    box-shadow: 0 6px 18px rgba(15,42,95,.12);
    animation: floatUpDown 3.5s ease-in-out infinite;
}
.float-icon:nth-child(2) { animation-delay: .3s; }
.float-icon:nth-child(3) { animation-delay: .6s; }
.float-icon:nth-child(4) { animation-delay: .9s; }
.float-icon:nth-child(5) { animation-delay: 1.2s; }
.float-icon:nth-child(6) { animation-delay: 1.5s; }

.feature-card {
    background: white;
    border-radius: 16px;
    padding: 18px 20px;
    border: 1px solid #dbe3ef;
    box-shadow: 0 5px 18px rgba(15,42,95,.06);
    opacity: 0;
    animation: fadeInUp 0.6s ease forwards;
    min-height: 150px;
    transition: transform .2s ease;
}
.feature-card:hover { transform: translateY(-4px); }
.feature-card:nth-child(1) { animation-delay: .1s; }
.feature-card:nth-child(2) { animation-delay: .25s; }
.feature-card:nth-child(3) { animation-delay: .4s; }
.feature-card:nth-child(4) { animation-delay: .55s; }

/* =========================================================
   LOADING MASCOTS (Yoom-style cute characters)
   ========================================================= */

@keyframes mascotBounce {
    0%, 100% { transform: translateY(0) rotate(0deg); }
    50% { transform: translateY(-22px) rotate(-6deg); }
}
@keyframes mascotBounce2 {
    0%, 100% { transform: translateY(0) rotate(0deg); }
    50% { transform: translateY(-16px) rotate(6deg); }
}
@keyframes mascotWave {
    0%, 100% { transform: rotate(0deg); }
    25% { transform: rotate(-10deg); }
    75% { transform: rotate(10deg); }
}
@keyframes textPulse {
    0%, 100% { opacity: 0.6; }
    50% { opacity: 1; }
}

.loading-mascot {
    background: linear-gradient(135deg, #eef4ff, #f4f7fb);
    border-radius: 20px;
    padding: 34px 20px;
    text-align: center;
    border: 1px solid #dbe3ef;
    margin: 10px 0 18px;
}
.mascot-row {
    display: flex;
    justify-content: center;
    align-items: flex-end;
    gap: 26px;
    margin-bottom: 14px;
}
.mascot-char {
    width: 64px; height: 64px;
    border-radius: 50%;
    display: flex; align-items: center; justify-content: center;
    font-size: 30px;
    box-shadow: 0 8px 20px rgba(15,42,95,.15);
}
.mascot-1 { background: #fde68a; animation: mascotBounce 1.6s ease-in-out infinite; }
.mascot-2 { background: #fbcfe8; animation: mascotWave 1.4s ease-in-out infinite; animation-delay: .15s; }
.mascot-3 { background: #a7f3d0; animation: mascotBounce2 1.8s ease-in-out infinite; animation-delay: .3s; }
.mascot-4 { background: #bfdbfe; animation: mascotBounce 2s ease-in-out infinite; animation-delay: .45s; }
.loading-text {
    color: #0f2a5f;
    font-weight: 700;
    font-size: 15px;
    animation: textPulse 1.4s ease-in-out infinite;
    margin: 0;
}

/* =========================================================
   HERO SIDE MASCOT (waving character next to title)
   ========================================================= */

.hero-flex {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 20px;
}
.hero-text-block {
    flex: 1;
}
.hero-mascot {
    width: 92px; height: 92px;
    border-radius: 50%;
    background: rgba(255,255,255,0.15);
    display: flex; align-items: center; justify-content: center;
    font-size: 46px;
    flex-shrink: 0;
    animation: floatUpDown 3s ease-in-out infinite;
    position: relative;
}
.hero-mascot .hero-hand {
    position: absolute;
    top: -6px; right: -6px;
    font-size: 26px;
    animation: mascotWave 1.1s ease-in-out infinite;
    transform-origin: 70% 70%;
}

/* =========================================================
   RUNNING LOADER OVERLAY (blur background + running character)
   ========================================================= */

@keyframes runCycle {
    0%   { transform: translateX(-120px) scaleX(1); }
    48%  { transform: translateX(90px) scaleX(1); }
    50%  { transform: translateX(90px) scaleX(-1); }
    98%  { transform: translateX(-120px) scaleX(-1); }
    100% { transform: translateX(-120px) scaleX(1); }
}
@keyframes legBounce {
    0%, 100% { transform: translateY(0); }
    50% { transform: translateY(-6px); }
}
@keyframes blurIn {
    from { backdrop-filter: blur(0px); opacity: 0; }
    to { backdrop-filter: blur(6px); opacity: 1; }
}

.run-overlay {
    background: rgba(244,247,251,0.75);
    animation: blurIn 0.5s ease forwards;
    border-radius: 20px;
    padding: 46px 20px;
    text-align: center;
    border: 1px solid #dbe3ef;
    margin: 10px 0 18px;
    overflow: hidden;
    position: relative;
}
.run-track {
    position: relative;
    height: 70px;
    display: flex;
    align-items: center;
    justify-content: center;
}
.runner {
    font-size: 44px;
    display: inline-block;
    animation: runCycle 2.2s linear infinite, legBounce 0.3s ease-in-out infinite;
}
.run-ground {
    width: 260px;
    height: 3px;
    background: repeating-linear-gradient(90deg, #b8c9df 0 14px, transparent 14px 28px);
    margin: 6px auto 0;
    opacity: 0.6;
}


/* =========================================================
   LOGIN SCREEN — 3D LIGHT THEME
   ========================================================= */
@keyframes loginImageIn {
    from { opacity: 0; transform: translateX(-35px) scale(.96); }
    to { opacity: 1; transform: translateX(0) scale(1); }
}
@keyframes loginFloat {
    0%, 100% { transform: translateY(0) scale(1); }
    50% { transform: translateY(-8px) scale(1.008); }
}
@keyframes loginCardIn {
    from { opacity: 0; transform: translateX(35px); }
    to { opacity: 1; transform: translateX(0); }
}
.login-page {
    min-height: 76vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 24px 0 40px;
}
.login-shell {
    width: 100%;
    max-width: 1180px;
    background: rgba(255,255,255,.94);
    border: 1px solid #dbe7f5;
    border-radius: 28px;
    box-shadow: 0 22px 70px rgba(15,42,95,.12);
    overflow: hidden;
    animation: fadeInUp .65s ease;
}
.login-brand {
    text-align: center;
    color: #0f2a5f;
    font-weight: 900;
    font-size: 13px;
    letter-spacing: .08em;
    text-transform: uppercase;
    margin: 0 0 10px;
}
.login-title {
    color: #172033 !important;
    font-size: 34px !important;
    font-weight: 900 !important;
    line-height: 1.1 !important;
    margin: 0 !important;
    text-align: center;
}
.login-subtitle {
    color: #64748b !important;
    text-align: center;
    margin: 10px 0 24px;
    font-size: 14px;
}
.login-visual-scene {
    width: 100%;
    max-width: 640px;
    display: flex;
    align-items: center;
    justify-content: center;
    animation: loginImageIn .85s ease both;
}
.login-character-svg {
    width: 100%;
    height: auto;
    display: block;
    overflow: visible;
    filter: drop-shadow(0 22px 35px rgba(15,42,95,.12));
}
.character {
    transform-box: fill-box;
    transform-origin: center bottom;
    animation: characterFloat 3.8s ease-in-out infinite;
}
.char-body {
    transform-box: fill-box;
    transform-origin: center bottom;
    animation: bodyBreath 2.8s ease-in-out infinite;
}
.char-head {
    transform-box: fill-box;
    transform-origin: center bottom;
    animation: headNod 3.4s ease-in-out infinite;
}
.char-arm-front {
    transform-box: fill-box;
    transform-origin: 294px 334px;
    animation: presentArm 2.2s ease-in-out infinite;
}
.char-arm-back {
    transform-box: fill-box;
    transform-origin: 225px 333px;
    animation: backArm 2.6s ease-in-out infinite;
}
.char-leg-a {
    transform-box: fill-box;
    transform-origin: 250px 425px;
    animation: walkLegA 1.05s ease-in-out infinite;
}
.char-leg-b {
    transform-box: fill-box;
    transform-origin: 292px 425px;
    animation: walkLegB 1.05s ease-in-out infinite;
}
.char-briefcase {
    transform-box: fill-box;
    transform-origin: 208px 395px;
    animation: briefcaseSwing 1.05s ease-in-out infinite;
}
.float-card {
    transform-box: fill-box;
    transform-origin: center;
}
.card-one { animation: cardFloat1 4.8s ease-in-out infinite; }
.card-two { animation: cardFloat2 4.2s ease-in-out .4s infinite; }
.card-three { animation: cardFloat3 5.1s ease-in-out .2s infinite; }
.card-four { animation: cardFloat4 4.6s ease-in-out .7s infinite; }
.card-five { animation: cardFloat5 4.9s ease-in-out .3s infinite; }
.books { animation: tinyFloat 3.5s ease-in-out infinite; }
.plant { animation: plantSway 4s ease-in-out infinite; transform-box: fill-box; transform-origin: center bottom; }
.sparkles { animation: sparklePulse 2.4s ease-in-out infinite; }

@keyframes characterFloat {
    0%, 100% { transform: translateY(0); }
    50% { transform: translateY(-7px); }
}
@keyframes bodyBreath {
    0%, 100% { transform: scaleY(1); }
    50% { transform: scaleY(1.012); }
}
@keyframes headNod {
    0%, 100% { transform: rotate(0deg); }
    45% { transform: rotate(-2deg); }
    75% { transform: rotate(1.5deg); }
}
@keyframes presentArm {
    0%, 100% { transform: rotate(0deg); }
    30% { transform: rotate(-4deg); }
    60% { transform: rotate(4deg); }
}
@keyframes backArm {
    0%, 100% { transform: rotate(0deg); }
    50% { transform: rotate(5deg); }
}
@keyframes walkLegA {
    0%, 100% { transform: rotate(3deg); }
    50% { transform: rotate(-7deg); }
}
@keyframes walkLegB {
    0%, 100% { transform: rotate(-7deg); }
    50% { transform: rotate(3deg); }
}
@keyframes briefcaseSwing {
    0%, 100% { transform: rotate(3deg); }
    50% { transform: rotate(-5deg); }
}
@keyframes cardFloat1 {
    0%, 100% { transform: translate(0,0) rotate(0deg); }
    50% { transform: translate(4px,-9px) rotate(-1deg); }
}
@keyframes cardFloat2 {
    0%, 100% { transform: translate(0,0) rotate(0deg); }
    50% { transform: translate(-5px,7px) rotate(1deg); }
}
@keyframes cardFloat3 {
    0%, 100% { transform: translate(0,0) rotate(0deg); }
    50% { transform: translate(5px,-8px) rotate(1deg); }
}
@keyframes cardFloat4 {
    0%, 100% { transform: translate(0,0) rotate(0deg); }
    50% { transform: translate(-5px,-6px) rotate(-1deg); }
}
@keyframes cardFloat5 {
    0%, 100% { transform: translate(0,0) rotate(0deg); }
    50% { transform: translate(4px,7px) rotate(1deg); }
}
@keyframes tinyFloat {
    0%, 100% { transform: translateY(0); }
    50% { transform: translateY(-3px); }
}
@keyframes plantSway {
    0%, 100% { transform: rotate(-1deg); }
    50% { transform: rotate(2deg); }
}
@keyframes sparklePulse {
    0%, 100% { opacity: .35; }
    50% { opacity: 1; }
}

.login-image-wrap {
    display: flex;
    align-items: center;
    justify-content: center;
    min-height: 620px;
    padding: 18px 12px 18px 24px;
    background: linear-gradient(135deg,#f8fbff 0%,#eef5ff 55%,#ffffff 100%);
    overflow: hidden;
}
.login-image {
    width: 100%;
    max-width: 640px;
    border-radius: 22px;
    display: block;
}
.login-form-wrap {
    padding: 54px 58px 46px;
    animation: loginCardIn .85s ease both;
}
.login-security {
    text-align: center;
    color: #94a3b8;
    font-size: 11px;
    margin-top: 22px;
}
.login-divider {
    height: 1px;
    background: #e6edf6;
    margin: 26px 0 0;
}
.login-small {
    text-align: center;
    color: #94a3b8;
    font-size: 11px;
    margin-top: 14px;
}
.login-locked-note {
    text-align: center;
    color: #16a34a;
    font-size: 12px;
    font-weight: 700;
    margin-top: 8px;
}
@media (max-width: 900px) {
    .login-form-wrap { padding: 36px 28px; }
    .login-image-wrap { min-height: 430px; padding: 12px; }
    .login-title { font-size: 28px !important; }
}

</style>
""", unsafe_allow_html=True)


# =========================================================
# CONSTANTS
# =========================================================

REQUIRED_FIELDS = ["GSTIN", "Invoice", "Party", "Taxable"]

OPTIONAL_FIELDS = [
    "Date",
    "IGST",
    "CGST",
    "SGST",
    "InvoiceValue"
]

STATUS_COLORS = {
    "Matched": "16A34A",
    "Missing in 2B": "DC2626",
    "Missing in Books": "EA580C",
    "Value Mismatch": "7C3AED",
}


# =========================================================
# HELPERS
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_gstin(value):
    if pd.isna(value):
        return ""
    return str(value).upper().replace(" ", "").strip()


def clean_invoice(value):
    if pd.isna(value):
        return ""

    value = str(value).upper().strip()

    if value.endswith(".0"):
        value = value[:-2]

    return re.sub(r"[^A-Z0-9]", "", value)


def clean_amount(value):
    if pd.isna(value):
        return 0.0

    try:
        text = str(value)
        text = text.replace(",", "")
        text = text.replace("₹", "")
        text = text.replace("Rs.", "")
        text = text.replace("Rs", "")
        text = text.strip()

        if text in ("", "-", "nan"):
            return 0.0

        return float(text)

    except Exception:
        return 0.0


def is_valid_gstin(gstin):
    if not gstin or len(gstin) != 15:
        return False

    pattern = (
        r"^[0-9]{2}[A-Z]{5}[0-9]{4}"
        r"[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$"
    )

    return bool(re.match(pattern, gstin))


# =========================================================
# COLUMN ALIASES
# =========================================================

ALIASES = {
    "GSTIN": [
        "GSTIN",
        "GSTIN/UIN",
        "GSTIN of Supplier",
        "GSTIN of supplier",
        "Supplier GSTIN",
        "GST Number",
        "GST No",
        "GST No.",
        "GSTIN of Vendor",
        "GSTIN/UIN of Supplier",
        "Vendor GSTIN"
    ],

    "Invoice": [
        "Invoice Number",
        "Invoice number",
        "Invoice No",
        "Invoice No.",
        "Document Number",
        "Document No",
        "Bill No",
        "Bill Number",
        "Invoice"
    ],

    "Date": [
        "Invoice Date",
        "Invoice date",
        "Document Date",
        "Document date",
        "Bill Date",
        "Date"
    ],

    "Party": [
        "Party Name",
        "Supplier Name",
        "Supplier",
        "Vendor Name",
        "Vendor",
        "Trade Name",
        "Legal Name",
        "Party",
        "Supplier Trade Name",
        "Trade/Legal Name"
    ],

    "Taxable": [
        "Taxable Value",
        "Taxable value",
        "Taxable Amount",
        "Taxable Amount (₹)",
        "Taxable Amt",
        "Taxable"
    ],

    "IGST": [
        "IGST",
        "IGST Amount",
        "IGST Amt",
        "Integrated Tax",
        "Integrated Tax Amount",
        "Integrated Tax Amount (₹)"
    ],

    "CGST": [
        "CGST",
        "CGST Amount",
        "CGST Amt",
        "Central Tax",
        "Central Tax Amount",
        "Central Tax Amount (₹)"
    ],

    "SGST": [
        "SGST",
        "SGST Amount",
        "SGST Amt",
        "UTGST",
        "State Tax",
        "State Tax Amount",
        "State/UT Tax",
        "State/UT Tax Amount",
        "State/UT Tax Amount (₹)"
    ],

    "InvoiceValue": [
        "Invoice Value",
        "Invoice value",
        "Total Invoice Value",
        "Total Value",
        "Document Value",
        "Invoice Amount",
        "Total Invoice Amount",
        "Invoice Value (₹)"
    ]
}


FIELD_LABELS = {
    "GSTIN": "GSTIN",
    "Invoice": "Invoice Number",
    "Date": "Invoice Date",
    "Party": "Party / Supplier Name",
    "Taxable": "Taxable Value",
    "IGST": "IGST",
    "CGST": "CGST",
    "SGST": "SGST",
    "InvoiceValue": "Invoice Value (Total)",
}


def normalize_column(value):
    return re.sub(r"[^A-Z0-9]", "", str(value).upper())


def find_column(df, aliases):
    columns = {
        normalize_column(c): c
        for c in df.columns
    }

    for alias in aliases:
        key = normalize_column(alias)

        if key in columns:
            return columns[key]

    for column in df.columns:
        current = normalize_column(column)

        for alias in aliases:
            target = normalize_column(alias)

            if target in current or current in target:
                return column

    return None


def auto_detect(df):
    return {
        key: find_column(df, ALIASES[key])
        for key in ALIASES
    }


# =========================================================
# HEADER ROW AUTO-DETECTION
# =========================================================

def detect_header_row(raw_df, max_scan_rows=15):
    """
    Existing header detection logic preserved.
    Scans first rows and finds a row containing
    both GSTIN and Invoice style headers.
    """

    scan_limit = min(max_scan_rows, len(raw_df))

    for row_idx in range(scan_limit):

        row_values = raw_df.iloc[row_idx].tolist()

        norm_values = [
            normalize_column(v)
            for v in row_values
            if pd.notna(v)
        ]

        if not norm_values:
            continue

        has_gstin = any(
            any(
                normalize_column(alias) in val
                or val in normalize_column(alias)
                for alias in ALIASES["GSTIN"]
            )
            for val in norm_values
        )

        has_invoice = any(
            any(
                normalize_column(alias) in val
                or val in normalize_column(alias)
                for alias in ALIASES["Invoice"]
            )
            for val in norm_values
        )

        if has_gstin and has_invoice:
            return row_idx

    return None


# =========================================================
# EXISTING EXCEL READER
# =========================================================

def read_excel(file):
    excel = pd.ExcelFile(file)

    all_data = []
    valid_sheets = []
    skipped_sheets = []

    for sheet in excel.sheet_names:

        try:

            raw = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None
            )

            if raw.empty:
                skipped_sheets.append(sheet)
                continue

            header_row_idx = detect_header_row(raw)

            if header_row_idx is None:
                skipped_sheets.append(sheet)
                continue

            df = pd.read_excel(
                file,
                sheet_name=sheet,
                header=header_row_idx
            )

            df = df.dropna(
                axis=1,
                how="all"
            )

            df = df.loc[
                :,
                ~df.columns.astype(str).str.startswith("Unnamed")
            ]

            gstin_column = find_column(
                df,
                ALIASES["GSTIN"]
            )

            invoice_column = find_column(
                df,
                ALIASES["Invoice"]
            )

            if gstin_column and invoice_column:

                df["_SOURCE_SHEET"] = sheet

                all_data.append(df)
                valid_sheets.append(sheet)

            else:
                skipped_sheets.append(sheet)

        except Exception:
            skipped_sheets.append(sheet)
            continue

    if not all_data:
        raise ValueError(
            "GSTIN aur Invoice Number wale columns kisi bhi sheet mein nahi mile."
        )

    combined = pd.concat(
        all_data,
        ignore_index=True
    )

    return combined, valid_sheets, skipped_sheets


# =========================================================
# NEW PORTAL-SPECIFIC READER
# EXISTING read_excel() IS NOT TOUCHED
# =========================================================

def portal_sheet_type(sheet_name):
    """
    Detects GST portal 2B sheet type from sheet name.
    """

    name = normalize_column(sheet_name)

    if "B2B" in name:
        return "B2B"

    if (
        "CDN" in name
        or "CREDITDEBIT" in name
        or "CREDITNOTE" in name
        or "DEBITNOTE" in name
    ):
        return "CDN"

    if (
        "RCM" in name
        or "REVERSECHARGE" in name
    ):
        return "RCM"

    if "ITCSUMMARY" in name:
        return "ITC Summary"

    if "SUMMARY" in name:
        return "Summary"

    if "HSN" in name:
        return "HSN"

    return "Other"


def read_portal_2b(file):
    """
    Separate portal reader.

    IMPORTANT:
    Existing read_excel() remains untouched.

    Portal logic:
    - B2B sheet is the primary 2B reconciliation source.
    - CDN / credit-debit note sheets are stored separately.
    - RCM sheets are stored separately.
    - ITC Summary and other portal sheets are not mixed into B2B.
    """

    excel = pd.ExcelFile(file)

    b2b_frames = []
    cdn_frames = []
    rcm_frames = []

    valid_b2b_sheets = []
    cdn_sheets = []
    rcm_sheets = []
    skipped_sheets = []

    for sheet in excel.sheet_names:

        try:

            sheet_type = portal_sheet_type(sheet)

            raw = pd.read_excel(
                file,
                sheet_name=sheet,
                header=None
            )

            if raw.empty:
                skipped_sheets.append(sheet)
                continue

            header_row_idx = detect_header_row(raw)

            if header_row_idx is None:

                # Some portal sheets may have slightly different
                # structures. Try a normal header read as fallback.
                try:
                    temp = pd.read_excel(
                        file,
                        sheet_name=sheet
                    )
                except Exception:
                    temp = pd.DataFrame()

                if temp.empty:
                    skipped_sheets.append(sheet)
                    continue

                df = temp

            else:

                df = pd.read_excel(
                    file,
                    sheet_name=sheet,
                    header=header_row_idx
                )

            df = df.dropna(
                axis=1,
                how="all"
            )

            df = df.loc[
                :,
                ~df.columns.astype(str).str.startswith("Unnamed")
            ]

            df["_SOURCE_SHEET"] = sheet

            gstin_column = find_column(
                df,
                ALIASES["GSTIN"]
            )

            invoice_column = find_column(
                df,
                ALIASES["Invoice"]
            )

            # -------------------------------------------------
            # B2B
            # -------------------------------------------------

            if sheet_type == "B2B":

                if gstin_column and invoice_column:
                    b2b_frames.append(df)
                    valid_b2b_sheets.append(sheet)
                else:
                    skipped_sheets.append(sheet)

            # -------------------------------------------------
            # CDN
            # -------------------------------------------------

            elif sheet_type == "CDN":

                cdn_frames.append(df)
                cdn_sheets.append(sheet)

            # -------------------------------------------------
            # RCM
            # -------------------------------------------------

            elif sheet_type == "RCM":

                rcm_frames.append(df)
                rcm_sheets.append(sheet)

            else:

                skipped_sheets.append(sheet)

        except Exception:
            skipped_sheets.append(sheet)

    # ---------------------------------------------------------
    # If no explicit B2B sheet found, fallback to existing reader
    # ---------------------------------------------------------

    if not b2b_frames:

        fallback_raw, fallback_sheets, fallback_skipped = read_excel(file)

        b2b_raw = fallback_raw

        valid_b2b_sheets = fallback_sheets

        skipped_sheets.extend(fallback_skipped)

    else:

        b2b_raw = pd.concat(
            b2b_frames,
            ignore_index=True
        )

    # ---------------------------------------------------------
    # CDN
    # ---------------------------------------------------------

    if cdn_frames:

        cdn_raw = pd.concat(
            cdn_frames,
            ignore_index=True
        )

    else:

        cdn_raw = pd.DataFrame()

    # ---------------------------------------------------------
    # RCM
    # ---------------------------------------------------------

    if rcm_frames:

        rcm_raw = pd.concat(
            rcm_frames,
            ignore_index=True
        )

    else:

        rcm_raw = pd.DataFrame()

    return {
        "b2b_raw": b2b_raw,
        "cdn_raw": cdn_raw,
        "rcm_raw": rcm_raw,
        "b2b_sheets": valid_b2b_sheets,
        "cdn_sheets": cdn_sheets,
        "rcm_sheets": rcm_sheets,
        "skipped_sheets": skipped_sheets,
        "all_sheets": excel.sheet_names
    }


# =========================================================
# STANDARDIZE
# =========================================================

def standardize(df, overrides=None):

    overrides = overrides or {}

    result = pd.DataFrame()
    detected = {}

    for key in ALIASES:

        manual = overrides.get(key)

        if (
            manual
            and manual != "— None —"
            and manual in df.columns
        ):
            detected[key] = manual

        else:
            detected[key] = find_column(
                df,
                ALIASES[key]
            )

    # GSTIN

    if detected["GSTIN"]:
        result["GSTIN"] = df[
            detected["GSTIN"]
        ].apply(clean_gstin)
    else:
        result["GSTIN"] = ""

    # Invoice

    if detected["Invoice"]:
        result["Invoice Number"] = df[
            detected["Invoice"]
        ].apply(clean_invoice)
    else:
        result["Invoice Number"] = ""

    # Date

    if detected["Date"]:
        result["Invoice Date"] = pd.to_datetime(
            df[detected["Date"]],
            errors="coerce"
        )
    else:
        result["Invoice Date"] = pd.NaT

    # Party

    if detected["Party"]:
        result["Party Name"] = df[
            detected["Party"]
        ].apply(clean_text)
    else:
        result["Party Name"] = ""

    # Taxable

    if detected["Taxable"]:
        result["Taxable Value"] = df[
            detected["Taxable"]
        ].apply(clean_amount)
    else:
        result["Taxable Value"] = 0.0

    # IGST

    if detected["IGST"]:
        result["IGST"] = df[
            detected["IGST"]
        ].apply(clean_amount)
    else:
        result["IGST"] = 0.0

    # CGST

    if detected["CGST"]:
        result["CGST"] = df[
            detected["CGST"]
        ].apply(clean_amount)
    else:
        result["CGST"] = 0.0

    # SGST

    if detected["SGST"]:
        result["SGST"] = df[
            detected["SGST"]
        ].apply(clean_amount)
    else:
        result["SGST"] = 0.0

    # Invoice Value

    if detected["InvoiceValue"]:

        result["Invoice Value"] = df[
            detected["InvoiceValue"]
        ].apply(clean_amount)

    else:

        result["Invoice Value"] = (
            result["Taxable Value"]
            + result["IGST"]
            + result["CGST"]
            + result["SGST"]
        )

    result["Source Sheet"] = (
        df["_SOURCE_SHEET"]
        if "_SOURCE_SHEET" in df.columns
        else ""
    )

    total_rows = len(result)

    blank_mask = (
        (result["GSTIN"] == "")
        |
        (result["Invoice Number"] == "")
    )

    dropped_rows = int(
        blank_mask.sum()
    )

    cleaned = result[
        ~blank_mask
    ].copy()

    dup_mask = cleaned.duplicated(
        subset=[
            "GSTIN",
            "Invoice Number"
        ],
        keep=False
    )

    duplicate_count = int(
        dup_mask.sum()
    )

    invalid_gstin_count = int(
        cleaned["GSTIN"]
        .apply(
            lambda g: not is_valid_gstin(g)
        )
        .sum()
    )

    quality = {
        "total_rows": total_rows,
        "dropped_rows": dropped_rows,
        "duplicate_rows": duplicate_count,
        "invalid_gstin_rows": invalid_gstin_count,
        "detected_columns": detected,
    }

    return (
        cleaned.reset_index(drop=True),
        quality
    )


# =========================================================
# NEW PORTAL NOTE / RCM ANALYSIS
# =========================================================

def find_note_type_column(df):
    possible = [
        "Note Type",
        "Document Type",
        "Debit/Credit Note",
        "Debit Credit Note",
        "Type",
        "Note Type (D/C)"
    ]

    return find_column(
        df,
        possible
    )


def classify_notes(cdn_raw):
    """
    Converts portal CDN data into separate Debit Note
    and Credit Note dataframes.
    """

    if cdn_raw is None or cdn_raw.empty:
        return pd.DataFrame(), pd.DataFrame()

    note_col = find_note_type_column(cdn_raw)

    if not note_col:
        return (
            cdn_raw.copy(),
            pd.DataFrame()
        )

    values = (
        cdn_raw[note_col]
        .fillna("")
        .astype(str)
        .str.upper()
    )

    debit_mask = (
        values.str.contains("DEBIT", na=False)
        |
        values.str.contains(r"\bDN\b", regex=True, na=False)
    )

    credit_mask = (
        values.str.contains("CREDIT", na=False)
        |
        values.str.contains(r"\bCN\b", regex=True, na=False)
    )

    debit_df = cdn_raw[
        debit_mask
    ].copy()

    credit_df = cdn_raw[
        credit_mask
    ].copy()

    return debit_df, credit_df


def find_rcm_column(df):
    possible = [
        "Reverse Charge",
        "Reverse charge",
        "RCM",
        "Reverse Charge Flag",
        "Supply Attract Reverse Charge",
        "Is Reverse Charge"
    ]

    return find_column(
        df,
        possible
    )


def analyse_rcm(two_b_raw, rcm_raw=None):
    """
    RCM can appear inside B2B itself as Reverse Charge = Yes.
    Separate RCM sheet is also considered if present.
    """

    rcm_b2b = pd.DataFrame()
    rcm_sheet = pd.DataFrame()

    if two_b_raw is not None and not two_b_raw.empty:

        rcm_col = find_rcm_column(two_b_raw)

        if rcm_col:

            values = (
                two_b_raw[rcm_col]
                .fillna("")
                .astype(str)
                .str.upper()
                .str.strip()
            )

            rcm_mask = values.isin(
                [
                    "YES",
                    "Y",
                    "TRUE",
                    "1"
                ]
            )

            rcm_b2b = two_b_raw[
                rcm_mask
            ].copy()

    if rcm_raw is not None and not rcm_raw.empty:
        rcm_sheet = rcm_raw.copy()

    if not rcm_b2b.empty and not rcm_sheet.empty:

        combined = pd.concat(
            [rcm_b2b, rcm_sheet],
            ignore_index=True
        )

    elif not rcm_b2b.empty:

        combined = rcm_b2b

    else:

        combined = rcm_sheet

    return combined


def portal_document_stats(
    two_b_raw,
    cdn_raw,
    rcm_raw
):

    debit_df, credit_df = classify_notes(
        cdn_raw
    )

    rcm_df = analyse_rcm(
        two_b_raw,
        rcm_raw
    )

    return {
        "debit_notes": debit_df,
        "credit_notes": credit_df,
        "rcm": rcm_df,
        "debit_count": len(debit_df),
        "credit_count": len(credit_df),
        "rcm_count": len(rcm_df),
    }


# =========================================================
# RECONCILIATION
# =========================================================

def build_dedup_key(df):

    df = df.copy()

    base_key = (
        df["GSTIN"]
        + "|"
        + df["Invoice Number"]
    )

    occurrence = (
        base_key.groupby(base_key)
        .cumcount()
    )

    df["KEY"] = (
        base_key
        + "|"
        + occurrence.astype(str)
    )

    return df


def suggest_close_matches(
    missing_row,
    other_df,
    field="Invoice Number",
    cutoff=0.82
):

    same_gstin = other_df[
        other_df["GSTIN"]
        == missing_row["GSTIN"]
    ]

    if same_gstin.empty:
        return None

    candidates = same_gstin[
        field
    ].tolist()

    best = difflib.get_close_matches(
        missing_row[field],
        candidates,
        n=1,
        cutoff=cutoff
    )

    if best:

        match_row = same_gstin[
            same_gstin[field]
            == best[0]
        ].iloc[0]

        return (
            best[0],
            match_row.get(
                "Invoice Value",
                None
            )
        )

    return None


def reconcile(
    two_b,
    books,
    tolerance,
    enable_fuzzy=True
):

    two_b_k = build_dedup_key(
        two_b
    )

    books_k = build_dedup_key(
        books
    )

    result = pd.merge(
        books_k,
        two_b_k,
        on="KEY",
        how="outer",
        suffixes=(
            "_Books",
            "_2B"
        ),
        indicator=True
    )

    statuses = []
    differences = []
    suggestions = []

    for _, row in result.iterrows():

        if row["_merge"] == "left_only":

            statuses.append(
                "Missing in 2B"
            )

            differences.append(
                row["IGST_Books"]
                + row["CGST_Books"]
                + row["SGST_Books"]
            )

            suggestion = ""

            if enable_fuzzy:

                probe = {
                    "GSTIN": row[
                        "GSTIN_Books"
                    ],
                    "Invoice Number": row[
                        "Invoice Number_Books"
                    ]
                }

                found = suggest_close_matches(
                    probe,
                    two_b
                )

                if found:

                    suggestion = (
                        f"Possible match in 2B: "
                        f"'{found[0]}' — "
                        f"check formatting/typo"
                    )

            suggestions.append(
                suggestion
            )

        elif row["_merge"] == "right_only":

            statuses.append(
                "Missing in Books"
            )

            differences.append(
                row["IGST_2B"]
                + row["CGST_2B"]
                + row["SGST_2B"]
            )

            suggestion = ""

            if enable_fuzzy:

                probe = {
                    "GSTIN": row[
                        "GSTIN_2B"
                    ],
                    "Invoice Number": row[
                        "Invoice Number_2B"
                    ]
                }

                found = suggest_close_matches(
                    probe,
                    books
                )

                if found:

                    suggestion = (
                        f"Possible match in Books: "
                        f"'{found[0]}' — "
                        f"check formatting/typo"
                    )

            suggestions.append(
                suggestion
            )

        else:

            taxable_diff = abs(
                row["Taxable Value_Books"]
                -
                row["Taxable Value_2B"]
            )

            igst_diff = abs(
                row["IGST_Books"]
                -
                row["IGST_2B"]
            )

            cgst_diff = abs(
                row["CGST_Books"]
                -
                row["CGST_2B"]
            )

            sgst_diff = abs(
                row["SGST_Books"]
                -
                row["SGST_2B"]
            )

            invoice_diff = abs(
                row["Invoice Value_Books"]
                -
                row["Invoice Value_2B"]
            )

            total_diff = (
                igst_diff
                + cgst_diff
                + sgst_diff
            )

            if (
                taxable_diff <= tolerance
                and igst_diff <= tolerance
                and cgst_diff <= tolerance
                and sgst_diff <= tolerance
                and invoice_diff <= tolerance
            ):

                statuses.append(
                    "Matched"
                )

            else:

                statuses.append(
                    "Value Mismatch"
                )

            differences.append(
                total_diff
            )

            suggestions.append("")

    result["Status"] = statuses

    result["ITC Difference"] = differences

    result["Match Suggestion"] = suggestions

    return result


# =========================================================
# DISPLAY DATA
# =========================================================

def prepare_display(df):

    output = pd.DataFrame()

    def get_column(name):

        if name in df.columns:
            return df[name]

        return pd.Series(
            [""] * len(df),
            index=df.index
        )

    gstin_books = (
        get_column("GSTIN_Books")
        .fillna("")
    )

    gstin_2b = (
        get_column("GSTIN_2B")
        .fillna("")
    )

    output["GSTIN"] = gstin_books.where(
        gstin_books != "",
        gstin_2b
    )

    party_books = (
        get_column("Party Name_Books")
        .fillna("")
    )

    party_2b = (
        get_column("Party Name_2B")
        .fillna("")
    )

    output["Party Name"] = party_books.where(
        party_books != "",
        party_2b
    )

    inv_books = (
        get_column("Invoice Number_Books")
        .fillna("")
    )

    inv_2b = (
        get_column("Invoice Number_2B")
        .fillna("")
    )

    output["Invoice Number"] = inv_books.where(
        inv_books != "",
        inv_2b
    )

    output["Invoice Date"] = (
        get_column("Invoice Date_Books")
        .where(
            get_column(
                "Invoice Date_Books"
            ).notna(),
            get_column(
                "Invoice Date_2B"
            )
        )
    )

    output["Taxable - Books"] = get_column(
        "Taxable Value_Books"
    )

    output["Taxable - 2B"] = get_column(
        "Taxable Value_2B"
    )

    output["IGST - Books"] = get_column(
        "IGST_Books"
    )

    output["IGST - 2B"] = get_column(
        "IGST_2B"
    )

    output["CGST - Books"] = get_column(
        "CGST_Books"
    )

    output["CGST - 2B"] = get_column(
        "CGST_2B"
    )

    output["SGST - Books"] = get_column(
        "SGST_Books"
    )

    output["SGST - 2B"] = get_column(
        "SGST_2B"
    )

    output["Invoice Value - Books"] = get_column(
        "Invoice Value_Books"
    )

    output["Invoice Value - 2B"] = get_column(
        "Invoice Value_2B"
    )

    output["ITC Difference"] = get_column(
        "ITC Difference"
    )

    output["Status"] = get_column(
        "Status"
    )

    output["Match Suggestion"] = get_column(
        "Match Suggestion"
    )

    return output.reset_index(
        drop=True
    )


def apply_match_overrides(display_df):

    overrides = st.session_state.get(
        "match_suggestion_overrides",
        {}
    )

    if not overrides or display_df.empty:
        return display_df

    keys = (
        display_df["GSTIN"]
        + "|"
        + display_df["Invoice Number"]
    )

    display_df = display_df.copy()

    display_df["Match Suggestion"] = [
        overrides.get(
            k,
            v
        )
        for k, v in zip(
            keys,
            display_df["Match Suggestion"]
        )
    ]

    return display_df


def vendor_summary(display_df):

    grouped = display_df.groupby(
        [
            "GSTIN",
            "Party Name"
        ],
        dropna=False
    ).agg(
        Total_Invoices=(
            "Invoice Number",
            "count"
        ),

        Matched=(
            "Status",
            lambda s:
            (s == "Matched").sum()
        ),

        Missing_in_2B=(
            "Status",
            lambda s:
            (s == "Missing in 2B").sum()
        ),

        Missing_in_Books=(
            "Status",
            lambda s:
            (s == "Missing in Books").sum()
        ),

        Value_Mismatch=(
            "Status",
            lambda s:
            (s == "Value Mismatch").sum()
        ),

        ITC_at_Risk=(
            "ITC Difference",
            "sum"
        ),
    ).reset_index()

    grouped = grouped.sort_values(
        "ITC_at_Risk",
        ascending=False
    )

    grouped["ITC_at_Risk"] = (
        grouped["ITC_at_Risk"]
        .round(2)
    )

    return grouped


# =========================================================
# EXCEL REPORT
# =========================================================

def _style_worksheet_header(
    ws,
    ncols
):

    header_fill = PatternFill(
        start_color="0F2A5F",
        end_color="0F2A5F",
        fill_type="solid"
    )

    header_font = Font(
        name="Arial",
        bold=True,
        color="FFFFFF",
        size=10
    )

    thin = Side(
        style="thin",
        color="D9D9D9"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    for col in range(
        1,
        ncols + 1
    ):

        c = ws.cell(
            row=1,
            column=col
        )

        c.fill = header_fill

        c.font = header_font

        c.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        c.border = border

    ws.freeze_panes = "A2"


def _autofit(
    ws,
    max_width=45
):

    for col_cells in ws.columns:

        length = max(
            (
                len(str(c.value))
                if c.value is not None
                else 0
            )
            for c in col_cells
        )

        col_letter = get_column_letter(
            col_cells[0].column
        )

        ws.column_dimensions[
            col_letter
        ].width = min(
            max(length + 3, 10),
            max_width
        )


def _format_data_rows(
    ws,
    ncols,
    status_col_index=None
):

    thin = Side(
        style="thin",
        color="EDEDED"
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin
    )

    money_cols = set()

    for idx, cell in enumerate(
        ws[1],
        start=1
    ):

        header = str(
            cell.value or ""
        )

        if any(
            tag in header
            for tag in [
                "Taxable",
                "IGST",
                "CGST",
                "SGST",
                "Invoice Value",
                "ITC"
            ]
        ):

            money_cols.add(idx)

    for row in ws.iter_rows(
        min_row=2,
        max_row=ws.max_row,
        max_col=ncols
    ):

        for cell in row:

            cell.font = Font(
                name="Arial",
                size=10
            )

            cell.border = border

            if cell.column in money_cols:

                cell.number_format = (
                    "#,##0.00"
                )

        if status_col_index:

            status_val = row[
                status_col_index - 1
            ].value

            color = STATUS_COLORS.get(
                status_val
            )

            if color:

                fill = PatternFill(
                    start_color=color,
                    end_color=color,
                    fill_type="solid"
                )

                font = Font(
                    name="Arial",
                    size=9,
                    bold=True,
                    color="FFFFFF"
                )

                cell = row[
                    status_col_index - 1
                ]

                cell.fill = fill

                cell.font = font

                cell.alignment = Alignment(
                    horizontal="center"
                )


def create_excel(
    result,
    quality_2b,
    quality_books,
    tolerance
):

    display = prepare_display(
        result
    )

    display = apply_match_overrides(
        display
    )

    v_summary = vendor_summary(
        display
    )

    total = len(display)

    matched = int(
        (
            display["Status"]
            == "Matched"
        ).sum()
    )

    missing_2b = int(
        (
            display["Status"]
            == "Missing in 2B"
        ).sum()
    )

    missing_books = int(
        (
            display["Status"]
            == "Missing in Books"
        ).sum()
    )

    mismatch = int(
        (
            display["Status"]
            == "Value Mismatch"
        ).sum()
    )

    itc_at_risk = display.loc[
        display["Status"] != "Matched",
        "ITC Difference"
    ].sum()

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df = pd.DataFrame({

            "Metric": [
                "Total Invoices Analysed",
                "Matched",
                "Missing in 2B",
                "Missing in Books",
                "Value Mismatch",
                "Match Rate (%)",
                "Total ITC at Risk (Rs.)",
                "Mismatch Tolerance Used (Rs.)",
            ],

            "Value": [
                total,
                matched,
                missing_2b,
                missing_books,
                mismatch,
                round(
                    (
                        matched
                        / total
                        * 100
                    ),
                    2
                )
                if total
                else 0,

                round(
                    itc_at_risk,
                    2
                ),

                tolerance,
            ]
        })

        summary_df.to_excel(
            writer,
            sheet_name="Summary",
            index=False
        )

        display_export = display.copy()

        if "Invoice Date" in display_export.columns:

            display_export[
                "Invoice Date"
            ] = pd.to_datetime(
                display_export[
                    "Invoice Date"
                ],
                errors="coerce"
            ).dt.strftime(
                "%d-%m-%Y"
            )

        display_export.to_excel(
            writer,
            sheet_name="Complete Reconciliation",
            index=False
        )

        for status in [
            "Matched",
            "Missing in 2B",
            "Missing in Books",
            "Value Mismatch"
        ]:

            temp = display_export[
                display_export["Status"]
                == status
            ]

            temp.to_excel(
                writer,
                sheet_name=status[:31],
                index=False
            )

        v_summary.to_excel(
            writer,
            sheet_name="Vendor Summary",
            index=False
        )

    output.seek(0)

    wb = load_workbook(
        output
    )

    ws_summary = wb[
        "Summary"
    ]

    _style_worksheet_header(
        ws_summary,
        ws_summary.max_column
    )

    _format_data_rows(
        ws_summary,
        ws_summary.max_column
    )

    _autofit(
        ws_summary
    )

    main_sheet_names = [
        "Complete Reconciliation",
        "Matched",
        "Missing in 2B",
        "Missing in Books",
        "Value Mismatch"
    ]

    for name in main_sheet_names:

        ws = wb[name]

        if ws.max_row < 1:
            continue

        _style_worksheet_header(
            ws,
            ws.max_column
        )

        status_idx = None

        for i, cell in enumerate(
            ws[1],
            start=1
        ):

            if cell.value == "Status":

                status_idx = i

                break

        _format_data_rows(
            ws,
            ws.max_column,
            status_col_index=status_idx
        )

        _autofit(ws)

    ws_vendor = wb[
        "Vendor Summary"
    ]

    _style_worksheet_header(
        ws_vendor,
        ws_vendor.max_column
    )

    _format_data_rows(
        ws_vendor,
        ws_vendor.max_column
    )

    _autofit(
        ws_vendor
    )

    final_output = BytesIO()

    wb.save(
        final_output
    )

    final_output.seek(0)

    return final_output


# =========================================================
# HEADER
# =========================================================

st.markdown("""
<div class="hero">
    <div class="hero-flex">
        <div class="hero-text-block">
            <h1>🧾 GST Reconciliation Pro<span class="hero-cursor"></span></h1>
            <p>GST 2B vs Books • ITC Reconciliation • Vendor & Invoice Exception Analysis</p>
        </div>
        <div class="hero-mascot">
            😊
            <span class="hero-hand">👋</span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.header(
        "⚙️ Reconciliation Settings"
    )

    tolerance = st.number_input(
        "Mismatch Tolerance ₹",
        min_value=0.0,
        value=2.0,
        step=0.50,
        help=(
            "Differences within this amount "
            "are still treated as a Match."
        )
    )

    enable_fuzzy = st.checkbox(
        "🔍 Suggest possible matches for gaps",
        value=True,
        help=(
            "Flags likely formatting/typo differences "
            "instead of a genuine missing invoice."
        )
    )

    st.divider()

    st.subheader(
        "🔑 Matching Key"
    )

    st.info(
        "GSTIN + Invoice Number\n\n"
        "(duplicate invoices are paired in sequence)"
    )

    st.divider()

    st.subheader(
        "📌 Status Legend"
    )

    st.markdown(
        "🟢 **Matched** — tallies within tolerance"
    )

    st.markdown(
        "🔴 **Missing in 2B** — in Books, not filed by supplier"
    )

    st.markdown(
        "🟠 **Missing in Books** — filed by supplier, not booked"
    )

    st.markdown(
        "🟣 **Value Mismatch** — amounts differ beyond tolerance"
    )


# =========================================================
# UPLOAD
# =========================================================

st.markdown(
    '<div class="section-title">📂 Upload GST Data</div>',
    unsafe_allow_html=True
)

col1, col2 = st.columns(2)


with col1:

    st.subheader(
        "📥 GST 2B File"
    )

    st.caption(
        "Upload GST portal 2B Excel file. Multiple sheets supported."
    )

    two_b_file = st.file_uploader(
        "Choose GST 2B Excel",
        type=["xlsx", "xls"],
        key="gst_2b_upload"
    )


with col2:

    st.subheader(
        "📚 Books File"
    )

    st.caption(
        "Upload your purchase/books Excel file."
    )

    books_file = st.file_uploader(
        "Choose Books Excel",
        type=["xlsx", "xls"],
        key="books_upload"
    )


# =========================================================
# OVERRIDES
# =========================================================

overrides_2b = {}
overrides_books = {}


# =========================================================
# LOAD FILES
# =========================================================

if two_b_file and books_file:

    try:

        # -----------------------------------------------------
        # PORTAL 2B
        # -----------------------------------------------------

        if (
            "portal_b2b_raw"
            not in st.session_state
            or
            st.session_state.get(
                "two_b_name"
            )
            != two_b_file.name
        ):

            portal_data = read_portal_2b(
                two_b_file
            )

            st.session_state[
                "portal_b2b_raw"
            ] = portal_data[
                "b2b_raw"
            ]

            st.session_state[
                "portal_cdn_raw"
            ] = portal_data[
                "cdn_raw"
            ]

            st.session_state[
                "portal_rcm_raw"
            ] = portal_data[
                "rcm_raw"
            ]

            st.session_state[
                "two_b_sheets"
            ] = portal_data[
                "b2b_sheets"
            ]

            st.session_state[
                "portal_cdn_sheets"
            ] = portal_data[
                "cdn_sheets"
            ]

            st.session_state[
                "portal_rcm_sheets"
            ] = portal_data[
                "rcm_sheets"
            ]

            st.session_state[
                "two_b_skipped"
            ] = portal_data[
                "skipped_sheets"
            ]

            st.session_state[
                "portal_all_sheets"
            ] = portal_data[
                "all_sheets"
            ]

            st.session_state[
                "two_b_name"
            ] = two_b_file.name

            st.session_state[
                "match_suggestion_overrides"
            ] = {}

        # -----------------------------------------------------
        # BOOKS
        # -----------------------------------------------------

        if (
            "books_raw"
            not in st.session_state
            or
            st.session_state.get(
                "books_name"
            )
            != books_file.name
        ):

            raw, sheets, skipped = read_excel(
                books_file
            )

            st.session_state[
                "books_raw"
            ] = raw

            st.session_state[
                "books_sheets"
            ] = sheets

            st.session_state[
                "books_skipped"
            ] = skipped

            st.session_state[
                "books_name"
            ] = books_file.name

            st.session_state[
                "match_suggestion_overrides"
            ] = {}

        two_b_raw = st.session_state[
            "portal_b2b_raw"
        ]

        books_raw = st.session_state[
            "books_raw"
        ]

        # =====================================================
        # COLUMN MAPPING
        # =====================================================

        with st.expander(
            "🔧 Column Mapping (verify or override auto-detection)"
        ):

            st.caption(
                "Har alag Excel export column names alag rakhta hai — "
                "humne best-guess mapping kar di hai. "
                "Zaroorat ho to yahan se change karein."
            )

            m1, m2 = st.columns(2)

            # -------------------------------------------------
            # 2B
            # -------------------------------------------------

            with m1:

                st.markdown(
                    "**GST 2B B2B columns**"
                )

                auto_2b = auto_detect(
                    two_b_raw
                )

                options_2b = [
                    "— None —"
                ] + list(
                    two_b_raw.columns
                )

                for field in (
                    REQUIRED_FIELDS
                    + OPTIONAL_FIELDS
                ):

                    default_val = (
                        auto_2b.get(field)
                        or
                        "— None —"
                    )

                    default_idx = (
                        options_2b.index(
                            default_val
                        )
                        if
                        default_val
                        in options_2b
                        else 0
                    )

                    overrides_2b[field] = st.selectbox(
                        FIELD_LABELS[field],
                        options_2b,
                        index=default_idx,
                        key=f"2b_{field}"
                    )

            # -------------------------------------------------
            # BOOKS
            # -------------------------------------------------

            with m2:

                st.markdown(
                    "**Books columns**"
                )

                auto_books = auto_detect(
                    books_raw
                )

                options_books = [
                    "— None —"
                ] + list(
                    books_raw.columns
                )

                for field in (
                    REQUIRED_FIELDS
                    + OPTIONAL_FIELDS
                ):

                    default_val = (
                        auto_books.get(field)
                        or
                        "— None —"
                    )

                    default_idx = (
                        options_books.index(
                            default_val
                        )
                        if
                        default_val
                        in options_books
                        else 0
                    )

                    overrides_books[field] = st.selectbox(
                        FIELD_LABELS[field],
                        options_books,
                        index=default_idx,
                        key=f"bk_{field}"
                    )

    except Exception as error:

        st.error(
            "❌ File padhne mein error aayi"
        )

        st.error(
            str(error)
        )


# =========================================================
# PROCESS
# =========================================================

if (
    two_b_file
    and books_file
    and "portal_b2b_raw"
    in st.session_state
):

    st.write("")

    if st.button(
        "🚀 START GST RECONCILIATION",
        type="primary",
        use_container_width=True
    ):

        try:

            mascot_placeholder = st.empty()

            mascot_placeholder.markdown("""
            <div class="run-overlay">
                <div class="run-track">
                    <span class="runner">🏃</span>
                </div>
                <div class="run-ground"></div>
                <p class="loading-text">
                    Matching your invoices, GSTIN by GSTIN...
                </p>
            </div>
            """, unsafe_allow_html=True)

            with st.spinner(
                "Analysing GST 2B and Books..."
            ):

                # -------------------------------------------------
                # IMPORTANT:
                # Only B2B is used for normal invoice reconciliation.
                # CDN and RCM remain separate.
                # -------------------------------------------------

                two_b, quality_2b = standardize(
                    st.session_state[
                        "portal_b2b_raw"
                    ],
                    overrides_2b
                )

                books, quality_books = standardize(
                    st.session_state[
                        "books_raw"
                    ],
                    overrides_books
                )

                if two_b.empty:

                    raise ValueError(
                        "GST 2B B2B sheet mein valid invoice data nahi mila."
                    )

                if books.empty:

                    raise ValueError(
                        "Books file mein valid invoice data nahi mila."
                    )

                result = reconcile(
                    two_b,
                    books,
                    tolerance,
                    enable_fuzzy=enable_fuzzy
                )

                # -------------------------------------------------
                # PORTAL DOCUMENT ANALYSIS
                # -------------------------------------------------

                portal_stats = portal_document_stats(
                    st.session_state.get(
                        "portal_b2b_raw",
                        pd.DataFrame()
                    ),
                    st.session_state.get(
                        "portal_cdn_raw",
                        pd.DataFrame()
                    ),
                    st.session_state.get(
                        "portal_rcm_raw",
                        pd.DataFrame()
                    )
                )

                st.session_state[
                    "result"
                ] = result

                st.session_state[
                    "two_b"
                ] = two_b

                st.session_state[
                    "books"
                ] = books

                st.session_state[
                    "quality_2b"
                ] = quality_2b

                st.session_state[
                    "quality_books"
                ] = quality_books

                st.session_state[
                    "tolerance_used"
                ] = tolerance

                st.session_state[
                    "selected_status"
                ] = "Missing in 2B"

                st.session_state[
                    "portal_stats"
                ] = portal_stats

                st.session_state.setdefault(
                    "match_suggestion_overrides",
                    {}
                )

            mascot_placeholder.empty()

            st.success(
                "✅ GST Reconciliation completed successfully."
            )

        except Exception as error:

            mascot_placeholder.empty()

            st.error(
                "❌ File processing error"
            )

            st.error(
                str(error)
            )


# =========================================================
# DATA QUALITY PANEL
# =========================================================

if (
    "quality_2b"
    in st.session_state
    and
    "result"
    in st.session_state
):

    q2b = st.session_state[
        "quality_2b"
    ]

    qbk = st.session_state[
        "quality_books"
    ]

    quality_flags = []

    if (
        q2b["dropped_rows"]
        or
        qbk["dropped_rows"]
    ):

        quality_flags.append(
            f"⚠️ Rows skipped due to blank GSTIN/Invoice — "
            f"2B: {q2b['dropped_rows']}, "
            f"Books: {qbk['dropped_rows']}"
        )

    if (
        q2b["duplicate_rows"]
        or
        qbk["duplicate_rows"]
    ):

        quality_flags.append(
            f"⚠️ Duplicate GSTIN+Invoice rows detected — "
            f"2B: {q2b['duplicate_rows']}, "
            f"Books: {qbk['duplicate_rows']} "
            "(paired in sequence during matching)"
        )

    if (
        q2b["invalid_gstin_rows"]
        or
        qbk["invalid_gstin_rows"]
    ):

        quality_flags.append(
            f"⚠️ Rows with non-standard GSTIN format — "
            f"2B: {q2b['invalid_gstin_rows']}, "
            f"Books: {qbk['invalid_gstin_rows']}"
        )

    with st.expander(
        "🩺 Data Quality Checks",
        expanded=bool(quality_flags)
    ):

        if quality_flags:

            for flag in quality_flags:

                st.markdown(
                    f'<div class="warn-box">{flag}</div>',
                    unsafe_allow_html=True
                )

        else:

            st.success(
                "Koi data quality issue nahi mila — files clean hain."
            )


# =========================================================
# DASHBOARD
# =========================================================

if "result" in st.session_state:

    result = st.session_state[
        "result"
    ]

    two_b = st.session_state[
        "two_b"
    ]

    books = st.session_state[
        "books"
    ]

    total = len(result)

    matched = int(
        (
            result["Status"]
            == "Matched"
        ).sum()
    )

    missing_2b = int(
        (
            result["Status"]
            == "Missing in 2B"
        ).sum()
    )

    missing_books = int(
        (
            result["Status"]
            == "Missing in Books"
        ).sum()
    )

    mismatch = int(
        (
            result["Status"]
            == "Value Mismatch"
        ).sum()
    )

    match_rate = round(
        (
            matched
            / total
            * 100
        ),
        1
    ) if total else 0.0

    # ---------------------------------------------------------
    # PORTAL STATS
    # ---------------------------------------------------------

    portal_stats = st.session_state.get(
        "portal_stats",
        {
            "debit_count": 0,
            "credit_count": 0,
            "rcm_count": 0
        }
    )

    debit_count = portal_stats.get(
        "debit_count",
        0
    )

    credit_count = portal_stats.get(
        "credit_count",
        0
    )

    rcm_count = portal_stats.get(
        "rcm_count",
        0
    )

    # ========================================================
    # DASHBOARD
    # ========================================================

    st.markdown(
        '<div class="section-title">📊 Reconciliation Dashboard</div>',
        unsafe_allow_html=True
    )

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    with c1:

        st.markdown(
            f"""
            <div class="kpi blue">
                <div class="kpi-title">
                    TOTAL INVOICES
                </div>
                <div class="kpi-value">
                    {total:,}
                </div>
                <div class="kpi-desc">
                    All analysed B2B records
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:

        st.markdown(
            f"""
            <div class="kpi green">
                <div class="kpi-title">
                    MATCHED
                </div>
                <div class="kpi-value">
                    {matched:,}
                </div>
                <div class="kpi-desc">
                    {match_rate}% match rate
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:

        st.markdown(
            f"""
            <div class="kpi red">
                <div class="kpi-title">
                    MISSING IN 2B
                </div>
                <div class="kpi-value">
                    {missing_2b:,}
                </div>
                <div class="kpi-desc">
                    Books → not in 2B
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c4:

        st.markdown(
            f"""
            <div class="kpi orange">
                <div class="kpi-title">
                    MISSING IN BOOKS
                </div>
                <div class="kpi-value">
                    {missing_books:,}
                </div>
                <div class="kpi-desc">
                    2B → not in Books
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c5:

        st.markdown(
            f"""
            <div class="kpi purple">
                <div class="kpi-title">
                    VALUE MISMATCH
                </div>
                <div class="kpi-value">
                    {mismatch:,}
                </div>
                <div class="kpi-desc">
                    Tax/value difference
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c6:

        itc_at_risk = result.loc[
            result["Status"] != "Matched",
            "ITC Difference"
        ].sum()

        st.markdown(
            f"""
            <div class="kpi teal">
                <div class="kpi-title">
                    ITC AT RISK
                </div>
                <div class="kpi-value">
                    ₹{itc_at_risk:,.0f}
                </div>
                <div class="kpi-desc">
                    Sum of open differences
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # NEW PORTAL DOCUMENT SUMMARY
    # ========================================================

    st.markdown(
        '<div class="section-title">📑 GST Portal Document Analysis</div>',
        unsafe_allow_html=True
    )

    d1, d2, d3 = st.columns(3)

    with d1:

        st.markdown(
            f"""
            <div class="kpi orange">
                <div class="kpi-title">
                    DEBIT NOTES
                </div>
                <div class="kpi-value">
                    {debit_count:,}
                </div>
                <div class="kpi-desc">
                    GST portal CDN records
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with d2:

        st.markdown(
            f"""
            <div class="kpi red">
                <div class="kpi-title">
                    CREDIT NOTES
                </div>
                <div class="kpi-value">
                    {credit_count:,}
                </div>
                <div class="kpi-desc">
                    GST portal CDN records
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with d3:

        st.markdown(
            f"""
            <div class="kpi purple">
                <div class="kpi-title">
                    RCM CASES
                </div>
                <div class="kpi-value">
                    {rcm_count:,}
                </div>
                <div class="kpi-desc">
                    Reverse Charge = Yes
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # PORTAL DOCUMENT DETAILS
    # ========================================================

    p1, p2, p3 = st.columns(3)

    with p1:

        with st.expander(
            f"🟠 Debit Notes ({debit_count})"
        ):

            debit_df = portal_stats.get(
                "debit_notes",
                pd.DataFrame()
            )

            if not debit_df.empty:

                st.dataframe(
                    debit_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "GST portal file mein Debit Note record nahi mila."
                )

    with p2:

        with st.expander(
            f"🔴 Credit Notes ({credit_count})"
        ):

            credit_df = portal_stats.get(
                "credit_notes",
                pd.DataFrame()
            )

            if not credit_df.empty:

                st.dataframe(
                    credit_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "GST portal file mein Credit Note record nahi mila."
                )

    with p3:

        with st.expander(
            f"🟣 RCM Cases ({rcm_count})"
        ):

            rcm_df = portal_stats.get(
                "rcm",
                pd.DataFrame()
            )

            if not rcm_df.empty:

                st.dataframe(
                    rcm_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "B2B mein Reverse Charge = Yes wala record nahi mila."
                )


    # ========================================================
    # INVOICE ANALYSIS
    # ========================================================

    st.markdown(
        '<div class="section-title">🔎 Invoice Analysis</div>',
        unsafe_allow_html=True
    )

    b1, b2, b3, b4 = st.columns(4)

    with b1:

        if st.button(
            f"🔴 Missing in 2B • {missing_2b}",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in 2B"

    with b2:

        if st.button(
            f"🟠 Missing in Books • {missing_books}",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Missing in Books"

    with b3:

        if st.button(
            f"🟣 Value Mismatch • {mismatch}",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Value Mismatch"

    with b4:

        if st.button(
            f"🟢 Matched • {matched}",
            use_container_width=True
        ):

            st.session_state[
                "selected_status"
            ] = "Matched"


    selected_status = st.session_state.get(
        "selected_status",
        "Missing in 2B"
    )

    detail = result[
        result["Status"]
        == selected_status
    ].copy()

    st.subheader(
        f"📋 {selected_status} — {len(detail):,} Invoices"
    )

    explanations = {

        "Missing in 2B":
            "Books mein invoice hai, lekin GST 2B mein nahi mila.",

        "Missing in Books":
            "GST 2B mein invoice hai, lekin Books mein nahi mila.",

        "Value Mismatch":
            "GSTIN aur invoice number match hain, lekin taxable/tax/invoice value different hai.",

        "Matched":
            "Books aur GST 2B mein invoice successfully match hua.",
    }

    st.info(
        explanations[selected_status]
    )

    display = prepare_display(
        detail
    )

    display = apply_match_overrides(
        display
    )

    search_term = st.text_input(
        "🔍 Search within this list (GSTIN, Invoice No, or Party Name)",
        placeholder="e.g. 27AABCU9603R1ZM or INV/2025/1001"
    )

    if not display.empty:

        display["Invoice Date"] = pd.to_datetime(
            display["Invoice Date"],
            errors="coerce"
        ).dt.strftime(
            "%d-%m-%Y"
        )

        display["Invoice Date"] = (
            display["Invoice Date"]
            .fillna("-")
        )

        if search_term.strip():

            term = (
                search_term
                .strip()
                .lower()
            )

            mask = (
                display["GSTIN"]
                .str.lower()
                .str.contains(
                    term,
                    na=False
                )
                |
                display["Invoice Number"]
                .str.lower()
                .str.contains(
                    term,
                    na=False
                )
                |
                display["Party Name"]
                .str.lower()
                .str.contains(
                    term,
                    na=False
                )
            )

            display = display[
                mask
            ]

        st.caption(
            "✏️ 'Match Suggestion' column mein khud bhi type kar sakte ho — remarks ya manual match note."
        )

        edited = st.data_editor(
            display,
            use_container_width=True,
            hide_index=True,
            disabled=[
                c
                for c in display.columns
                if c != "Match Suggestion"
            ],
            column_config={
                "Match Suggestion":
                    st.column_config.TextColumn(
                        "Match Suggestion / Remarks",
                        help=(
                            "Apna remark ya manual match note yahan type karein"
                        ),
                        width="large",
                    )
            },
            key=f"editor_{selected_status}",
        )

        if not edited.empty:

            overrides = st.session_state.setdefault(
                "match_suggestion_overrides",
                {}
            )

            edit_keys = (
                edited["GSTIN"]
                + "|"
                + edited["Invoice Number"]
            )

            for k, v in zip(
                edit_keys,
                edited["Match Suggestion"]
            ):

                overrides[k] = v

        update_col, _ = st.columns(
            [1, 3]
        )

        with update_col:

            if st.button(
                "💾 Update Changes",
                key=f"save_{selected_status}",
                use_container_width=True
            ):

                overrides = st.session_state.setdefault(
                    "match_suggestion_overrides",
                    {}
                )

                edit_keys = (
                    edited["GSTIN"]
                    + "|"
                    + edited["Invoice Number"]
                )

                for k, v in zip(
                    edit_keys,
                    edited["Match Suggestion"]
                ):

                    overrides[k] = v

                st.success(
                    "✅ Changes saved. Ye notes ab tab tak bane rahenge jab tak koi nayi file upload na ho — dobara reconcile karne se bhi nahi udenge."
                )

        csv_bytes = (
            edited
            .to_csv(index=False)
            .encode("utf-8-sig")
        )

        st.download_button(
            f"⬇️ Download '{selected_status}' as CSV",
            data=csv_bytes,
            file_name=(
                f"GST_{selected_status.replace(' ', '_')}.csv"
            ),
            mime="text/csv",
        )

    else:

        st.success(
            "Is category mein koi invoice nahi hai."
        )


    # ========================================================
    # MANUAL SEARCH
    # ========================================================

    if (
        selected_status
        in (
            "Missing in 2B",
            "Missing in Books"
        )
        and
        not detail.empty
    ):

        other_df = (
            two_b
            if selected_status
            == "Missing in 2B"
            else books
        )

        other_label = (
            "GST 2B"
            if selected_status
            == "Missing in 2B"
            else "Books"
        )

        with st.expander(
            f"🔎 Manually search {other_label} for a possible match",
            expanded=False
        ):

            st.caption(
                f"Auto-suggestion ne kuch na pakda ho to yahan khud type karke {other_label} file mein dhoond lo — GSTIN, Invoice Number, ya Party Name, kisi se bhi search ho jayega."
            )

            manual_query = st.text_input(
                f"Type to search in {other_label}",
                key=f"manual_search_{selected_status}",
                placeholder="e.g. GSTIN ka part, invoice number, ya vendor ka naam"
            )

            if manual_query.strip():

                q = (
                    manual_query
                    .strip()
                    .lower()
                )

                other_disp = other_df.copy()

                match_mask = (
                    other_disp["GSTIN"]
                    .str.lower()
                    .str.contains(
                        q,
                        na=False
                    )
                    |
                    other_disp[
                        "Invoice Number"
                    ]
                    .str.lower()
                    .str.contains(
                        q,
                        na=False
                    )
                    |
                    other_disp[
                        "Party Name"
                    ]
                    .str.lower()
                    .str.contains(
                        q,
                        na=False
                    )
                )

                found_cols = [
                    "GSTIN",
                    "Party Name",
                    "Invoice Number",
                    "Invoice Date",
                    "Taxable Value",
                    "IGST",
                    "CGST",
                    "SGST",
                    "Invoice Value"
                ]

                found = other_disp.loc[
                    match_mask,
                    found_cols
                ].copy()

                if not found.empty:

                    found["Invoice Date"] = pd.to_datetime(
                        found["Invoice Date"],
                        errors="coerce"
                    ).dt.strftime(
                        "%d-%m-%Y"
                    )

                    found["Invoice Date"] = (
                        found["Invoice Date"]
                        .fillna("-")
                    )

                    st.dataframe(
                        found,
                        use_container_width=True,
                        hide_index=True
                    )

                    st.caption(
                        f"{len(found)} possible record(s) mile {other_label} mein."
                    )

                else:

                    st.warning(
                        f"Koi record nahi mila jo '{manual_query}' se match kare."
                    )

            else:

                st.caption(
                    "Search shuru karne ke liye upar type karo."
                )


    # ========================================================
    # VENDOR SUMMARY
    # ========================================================

    st.markdown(
        '<div class="section-title">🏢 Vendor-wise Summary</div>',
        unsafe_allow_html=True
    )

    st.caption(
        "Sabse zyada ITC-at-risk wale vendors sabse upar dikhaye gaye hain."
    )

    full_display = prepare_display(
        result
    )

    v_summary = vendor_summary(
        full_display
    )

    st.dataframe(
        v_summary.rename(
            columns={
                "Total_Invoices":
                    "Total Invoices",

                "Missing_in_2B":
                    "Missing in 2B",

                "Missing_in_Books":
                    "Missing in Books",

                "Value_Mismatch":
                    "Value Mismatch",

                "ITC_at_Risk":
                    "ITC at Risk (Rs.)",
            }
        ).style.format(
            {
                "ITC at Risk (Rs.)":
                    "₹{:,.2f}"
            }
        ),
        use_container_width=True,
        hide_index=True,
    )


    # ========================================================
    # ITC ANALYSIS
    # ========================================================

    st.markdown(
        '<div class="section-title">💰 ITC Analysis</div>',
        unsafe_allow_html=True
    )

    books_igst = books[
        "IGST"
    ].sum()

    books_cgst = books[
        "CGST"
    ].sum()

    books_sgst = books[
        "SGST"
    ].sum()

    two_b_igst = two_b[
        "IGST"
    ].sum()

    two_b_cgst = two_b[
        "CGST"
    ].sum()

    two_b_sgst = two_b[
        "SGST"
    ].sum()

    books_itc = (
        books_igst
        + books_cgst
        + books_sgst
    )

    two_b_itc = (
        two_b_igst
        + two_b_cgst
        + two_b_sgst
    )

    itc_difference = (
        books_itc
        - two_b_itc
    )

    i1, i2, i3 = st.columns(3)

    with i1:

        st.metric(
            "📚 Books Total ITC",
            f"₹{books_itc:,.2f}"
        )

    with i2:

        st.metric(
            "📥 GST 2B Total ITC",
            f"₹{two_b_itc:,.2f}"
        )

    with i3:

        st.metric(
            "⚖️ ITC Difference",
            f"₹{itc_difference:,.2f}"
        )

    tax_df = pd.DataFrame({

        "Tax Component": [
            "IGST",
            "CGST",
            "SGST",
            "TOTAL ITC"
        ],

        "Books": [
            books_igst,
            books_cgst,
            books_sgst,
            books_itc
        ],

        "GST 2B": [
            two_b_igst,
            two_b_cgst,
            two_b_sgst,
            two_b_itc
        ],

        "Difference": [
            books_igst - two_b_igst,
            books_cgst - two_b_cgst,
            books_sgst - two_b_sgst,
            itc_difference
        ]
    })

    st.dataframe(
        tax_df.style.format(
            {
                "Books":
                    "₹{:,.2f}",

                "GST 2B":
                    "₹{:,.2f}",

                "Difference":
                    "₹{:,.2f}"
            }
        ),
        use_container_width=True,
        hide_index=True
    )


    # ========================================================
    # CHARTS
    # ========================================================

    st.markdown(
        '<div class="section-title">📈 Visual Analysis</div>',
        unsafe_allow_html=True
    )

    chart1, chart2 = st.columns(2)

    with chart1:

        status_df = pd.DataFrame({

            "Status": [
                "Matched",
                "Missing in 2B",
                "Missing in Books",
                "Value Mismatch"
            ],

            "Count": [
                matched,
                missing_2b,
                missing_books,
                mismatch
            ]
        })

        status_df = status_df[
            status_df["Count"] > 0
        ]

        if not status_df.empty:

            fig = px.pie(
                status_df,
                names="Status",
                values="Count",
                hole=0.55,
                title="Reconciliation Status",
                color="Status",
                color_discrete_map={
                    "Matched":
                        "#16a34a",

                    "Missing in 2B":
                        "#dc2626",

                    "Missing in Books":
                        "#ea580c",

                    "Value Mismatch":
                        "#7c3aed",
                }
            )

            fig.update_traces(
                textinfo="percent+label"
            )

            st.plotly_chart(
                fig,
                use_container_width=True
            )


    with chart2:

        top_vendors = v_summary[
            v_summary["ITC_at_Risk"] > 0
        ].head(10)

        if not top_vendors.empty:

            fig3 = px.bar(
                top_vendors.sort_values(
                    "ITC_at_Risk"
                ),
                x="ITC_at_Risk",
                y="Party Name",
                orientation="h",
                title="Top 10 Vendors by ITC at Risk",
                labels={
                    "ITC_at_Risk":
                        "ITC at Risk (Rs.)",

                    "Party Name":
                        ""
                },
            )

            fig3.update_traces(
                marker_color="#7c3aed"
            )

            st.plotly_chart(
                fig3,
                use_container_width=True
            )

        else:

            itc_df = pd.DataFrame({

                "Tax": [
                    "IGST",
                    "CGST",
                    "SGST"
                ],

                "Amount": [
                    two_b_igst,
                    two_b_cgst,
                    two_b_sgst
                ]
            })

            itc_df = itc_df[
                itc_df["Amount"] > 0
            ]

            if not itc_df.empty:

                fig2 = px.pie(
                    itc_df,
                    names="Tax",
                    values="Amount",
                    hole=0.55,
                    title="GST 2B ITC Distribution"
                )

                fig2.update_traces(
                    textinfo="percent+label"
                )

                st.plotly_chart(
                    fig2,
                    use_container_width=True
                )


    # ========================================================
    # MONTHLY TREND
    # ========================================================

    combined_dates = pd.concat(
        [
            two_b[
                [
                    "Invoice Date",
                    "Taxable Value"
                ]
            ].assign(
                Source="GST 2B"
            ),

            books[
                [
                    "Invoice Date",
                    "Taxable Value"
                ]
            ].assign(
                Source="Books"
            ),
        ]
    )

    combined_dates = combined_dates.dropna(
        subset=[
            "Invoice Date"
        ]
    )

    if not combined_dates.empty:

        combined_dates["Month"] = (
            combined_dates[
                "Invoice Date"
            ]
            .dt.to_period("M")
            .astype(str)
        )

        monthly = (
            combined_dates
            .groupby(
                [
                    "Month",
                    "Source"
                ]
            )["Taxable Value"]
            .sum()
            .reset_index()
        )

        if (
            monthly["Month"]
            .nunique()
            > 1
        ):

            fig4 = px.bar(
                monthly,
                x="Month",
                y="Taxable Value",
                color="Source",
                barmode="group",
                title=(
                    "Month-wise Taxable Value: "
                    "GST 2B vs Books"
                ),
                color_discrete_map={
                    "GST 2B":
                        "#2563eb",

                    "Books":
                        "#16a34a"
                },
            )

            st.plotly_chart(
                fig4,
                use_container_width=True
            )


    # ========================================================
    # PORTAL SHEETS
    # ========================================================

    with st.expander(
        "📂 GST Portal Excel Sheets Automatically Analysed"
    ):

        st.info(
            "Reconciliation ke liye sirf B2B invoices use kiye gaye hain. "
            "Debit/Credit Notes aur RCM ko separately analyse kiya gaya hai."
        )

        col_a, col_b, col_c = st.columns(3)

        with col_a:

            st.subheader(
                "📥 B2B"
            )

            for sheet in st.session_state.get(
                "two_b_sheets",
                []
            ):

                st.write(
                    "•",
                    sheet
                )

        with col_b:

            st.subheader(
                "📝 Debit/Credit Notes"
            )

            for sheet in st.session_state.get(
                "portal_cdn_sheets",
                []
            ):

                st.write(
                    "•",
                    sheet
                )

        with col_c:

            st.subheader(
                "🔄 RCM"
            )

            for sheet in st.session_state.get(
                "portal_rcm_sheets",
                []
            ):

                st.write(
                    "•",
                    sheet
                )

        st.divider()

        st.caption(
            "Other portal sheets:"
        )

        all_sheets = st.session_state.get(
            "portal_all_sheets",
            []
        )

        used_sheets = (
            st.session_state.get(
                "two_b_sheets",
                []
            )
            +
            st.session_state.get(
                "portal_cdn_sheets",
                []
            )
            +
            st.session_state.get(
                "portal_rcm_sheets",
                []
            )
        )

        for sheet in all_sheets:

            if sheet not in used_sheets:

                st.write(
                    "•",
                    sheet
                )


    # ========================================================
    # EXPORT
    # ========================================================

    st.markdown(
        '<div class="section-title">📥 Export Report</div>',
        unsafe_allow_html=True
    )

    excel_report = create_excel(
        result,
        st.session_state[
            "quality_2b"
        ],
        st.session_state[
            "quality_books"
        ],
        st.session_state[
            "tolerance_used"
        ],
    )

    st.download_button(
        "📊 DOWNLOAD COMPLETE EXCEL REPORT",
        data=excel_report,
        file_name="GST_Reconciliation_Report.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        use_container_width=True,
    )

    st.caption(
        "Report mein Summary, Complete Reconciliation, "
        "status-wise sheets, aur Vendor Summary — "
        "sab color-coded aur formatted hain."
    )


# =========================================================
# WELCOME
# =========================================================

else:

    st.markdown("""
    <div class="float-icons">
        <div class="float-icon">📥</div>
        <div class="float-icon">📚</div>
        <div class="float-icon">🔍</div>
        <div class="float-icon">📊</div>
        <div class="float-icon">✅</div>
        <div class="float-icon">💰</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="info-box" style="animation: fadeInUp 0.6s ease;">

    <h3>👋 Welcome to GST Reconciliation Pro</h3>

    Upload your <b>GST 2B</b> and <b>Books</b> Excel files above to get started.

    <br><br>

    The software automatically detects important GST columns even when the
    column names are different — and lets you fine-tune the mapping if needed.

    <br><br>

    <b>GST Portal Support:</b><br>
    B2B invoices are used for reconciliation while Debit Notes,
    Credit Notes and RCM cases are analysed separately.

    <br><br>

    <b>Matching Key:</b><br>
    GSTIN + Invoice Number (duplicates paired in sequence)

    <br><br>

    🔴 Missing in 2B<br>
    🟠 Missing in Books<br>
    🟣 Value Mismatch<br>
    🟢 Matched

    </div>
    """, unsafe_allow_html=True)

    st.markdown(
        '<div class="section-title">✨ What you get</div>',
        unsafe_allow_html=True
    )

    features = [
        (
            "🎯",
            "Smart Status Matching",
            "Matched, Missing in 2B, Missing in Books & Value Mismatch — auto classified."
        ),
        (
            "🧠",
            "Fuzzy Match Suggestions",
            "Typo or formatting mismatches get flagged instead of marked genuinely missing."
        ),
        (
            "🏢",
            "Vendor-wise ITC Risk",
            "See which vendors are holding up your ITC, ranked by risk amount."
        ),
        (
            "📊",
            "Polished Excel Export",
            "Color-coded, multi-sheet, ready-to-share report in one click."
        ),
    ]

    f_cols = st.columns(4)

    for f_col, (icon, title, desc) in zip(f_cols, features):

        with f_col:

            st.markdown(
                f"""
                <div class="feature-card">
                    <div style="font-size:28px;">{icon}</div>
                    <div style="font-weight:800; color:#0f2a5f; margin-top:8px;">{title}</div>
                    <div style="color:#64748b; font-size:12.5px; margin-top:6px;">{desc}</div>
                </div>
                """,
                unsafe_allow_html=True
            )

    st.markdown(
        '<div class="section-title">🛠️ What\'s included</div>',
        unsafe_allow_html=True
    )

    st.markdown("""
    <div class="info-box" style="animation: fadeInUp 0.6s ease .15s backwards;">
    <ul>
        <li>Manual column-mapping override for non-standard exports</li>
        <li>Data quality checks</li>
        <li>Smart match suggestions</li>
        <li>Vendor-wise ITC-at-risk summary</li>
        <li>Search + CSV export</li>
        <li>Professionally formatted Excel report</li>
        <li>GST Portal B2B + CDN + RCM analysis</li>
    </ul>
    </div>
    """, unsafe_allow_html=True)
