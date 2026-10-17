"""Sidebar: New chat, Settings button, Open project, chat list.

Edit when: changing the sidebar.
"""

import streamlit as st

from chatppt.project_io import ProjectFileError, load_project
from chatppt.ui.common import ASSISTANT_AVATAR
from chatppt.ui.settings import open_settings
from chatppt.ui.state import create_chat, delete_chat, restore_project


def _restore_uploaded_project(upload_key: str) -> None:
    """Callback (runs before widgets exist, so composer options can be restored too)."""
    ss = st.session_state
    uploaded = ss.get(upload_key)
    if uploaded is None:
        return
    try:
        loaded = load_project(uploaded.getvalue())
    except ProjectFileError as exc:
        ss["_project_error"] = str(exc)
        return
    restore_project(loaded)
    ss["project_upload_ver"] = ss.get("project_upload_ver", 0) + 1  # new key => empty uploader
    st.toast(f"Opened “{loaded['chat']['title']}”", icon=":material/folder_open:")


def _render_open_project() -> None:
    ss = st.session_state
    with st.popover("Open project", icon=":material/folder_open:", width="stretch"):
        upload_key = f"project_upload_{ss.get('project_upload_ver', 0)}"
        st.file_uploader(
            "Project file (.json)",
            type=["json"],
            key=upload_key,
            help="A project saved with “Save project” above the slide outline. It opens as a new chat.",
        )
        st.button(
            "Open as new chat",
            type="primary",
            icon=":material/upload_file:",
            width="stretch",
            on_click=_restore_uploaded_project,
            args=(upload_key,),
            disabled=ss.get(upload_key) is None,
            key="open_project_btn",
        )
        error = ss.pop("_project_error", None)
        if error:
            st.error(error, icon=":material/error:")


def render_sidebar() -> None:
    ss = st.session_state
    with st.sidebar:
        st.markdown(f"### {ASSISTANT_AVATAR} ChatPPT")
        if st.button("New chat", icon=":material/add:", key="new_chat_btn", width="stretch"):
            create_chat()
            st.rerun()
        if st.button("Settings", icon=":material/settings:", key="settings_btn", width="stretch"):
            open_settings()
        _render_open_project()

        st.markdown('<div class="sidebar-label">Chats</div>', unsafe_allow_html=True)
        for cid in list(ss.chat_order):
            title = ss.chats[cid]["title"]
            title = title if len(title) <= 24 else title[:23] + "…"
            c1, c2 = st.columns([6, 1], vertical_alignment="center")
            if c1.button(title, icon=":material/chat_bubble:", key=f"chat_btn_{cid}", width="stretch"):
                ss.active_chat = cid
                st.rerun()
            if c2.button(":material/close:", key=f"del_{cid}", help="Delete chat"):
                delete_chat(cid)
                st.rerun()
