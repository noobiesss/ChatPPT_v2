"""Global stylesheet and theme-aware CSS injection.

Edit when: changing the look of the chat UI, composer, or sidebar.
"""

import streamlit as st


APP_CSS = """
<style>
/* Theme-neutral styling: text and page colors come from the active Streamlit theme (light or dark).
   Surfaces use translucent grey so contrast is preserved in both themes. */
header[data-testid="stHeader"] { background: transparent; }
.block-container { max-width: 880px; padding-top: 2rem; padding-bottom: 1rem; }

/* Sidebar */
[data-testid="stSidebar"] .stButton > button {
    justify-content: flex-start; text-align: left; border: none; background: transparent;
    box-shadow: none; font-weight: 400; padding: 0.45rem 0.7rem; border-radius: 10px;
}
[data-testid="stSidebar"] .stButton > button:hover { background: rgba(128, 128, 128, 0.16); }
.st-key-new_chat_btn button {
    background: #C96442 !important; color: #FFFFFF !important; font-weight: 600 !important;
    justify-content: center !important;
}
.st-key-new_chat_btn button:hover { background: #B5573A !important; }
.st-key-settings_btn button {
    border: 1px solid rgba(128, 128, 128, 0.40) !important; justify-content: center !important;
}
.sidebar-label {
    font-size: 0.72rem; text-transform: uppercase; letter-spacing: 0.08em;
    opacity: 0.6; margin: 1.1rem 0 0.3rem 0.4rem;
}

/* Hero (empty chat state) */
.hero { text-align: center; padding: 3rem 0 1.5rem; }
.hero h1 { font-size: 2.1rem; font-weight: 600; letter-spacing: -0.02em; margin-bottom: 0.4rem; }
.hero p { opacity: 0.7; margin: 0; }
.steps { display: grid; grid-template-columns: repeat(3, 1fr); gap: 0.8rem; margin: 1.6rem 0 0.5rem; }
.step {
    background: rgba(128, 128, 128, 0.09); border: 1px solid rgba(128, 128, 128, 0.28);
    border-radius: 14px; padding: 0.9rem 1rem; text-align: left;
}
.step b { display: block; margin-bottom: 0.2rem; }
.step span { opacity: 0.72; font-size: 0.88rem; }
@media (max-width: 700px) { .steps { grid-template-columns: 1fr; } }

/* Primary buttons share the app accent color in light and dark themes */
[data-testid="stBaseButton-primary"] { background: #C96442; border-color: #C96442; color: #FFFFFF; }
[data-testid="stBaseButton-primary"]:hover { background: #B5573A; border-color: #B5573A; color: #FFFFFF; }

/* Composer dock (sticky at the bottom of the main area).
   Theme-independent: a frosted blur strip instead of a hard-coded background color, so it looks
   right in light and dark mode without re-running the script after a theme switch. */
.st-key-composer_dock {
    position: sticky; bottom: 0; z-index: 100; padding: 1.5rem 0 0.5rem; background: transparent;
}
.st-key-composer_dock::before {
    content: ""; position: absolute; inset: 0; z-index: -1; pointer-events: none;
    -webkit-backdrop-filter: blur(12px); backdrop-filter: blur(12px);
    -webkit-mask-image: linear-gradient(to top, #000 70%, transparent);
    mask-image: linear-gradient(to top, #000 70%, transparent);
}
.st-key-composer {
    background: rgba(128, 128, 128, 0.10); border: 1px solid rgba(128, 128, 128, 0.35);
    border-radius: 20px; padding: 0.7rem 1rem 0.6rem;
    -webkit-backdrop-filter: blur(16px); backdrop-filter: blur(16px);
    box-shadow: 0 2px 14px rgba(0, 0, 0, 0.06);
}
.st-key-composer textarea { font-size: 1rem; background: transparent !important; box-shadow: none !important; }
.st-key-composer [data-baseweb="textarea"], .st-key-composer [data-baseweb="base-input"] {
    border: none !important; background: transparent !important; box-shadow: none !important; outline: none !important;
}
.composer-note { text-align: center; font-size: 0.74rem; opacity: 0.6; margin-top: 0.4rem; }
.st-key-composer .stButton > button, .st-key-composer [data-testid="stPopover"] > div > button {
    border-radius: 999px; font-size: 0.82rem; padding: 0.25rem 0.5rem; white-space: nowrap;
    overflow: visible;
}
.st-key-composer button p { white-space: nowrap; overflow: visible; text-overflow: clip; }
.st-key-composer [data-baseweb="select"] > div { border-radius: 999px; min-height: 2rem; font-size: 0.82rem; }

/* Active chat highlight (injected per run) */
.st-key-chat_btn___ACTIVE__ button { background: rgba(128, 128, 128, 0.26) !important; font-weight: 600 !important; }
</style>
"""


def inject_css(active_chat_id: str) -> None:
    """Inject the global stylesheet and highlight the active chat button in the sidebar."""
    css = APP_CSS.replace("__ACTIVE__", active_chat_id)
    st.markdown(css, unsafe_allow_html=True)