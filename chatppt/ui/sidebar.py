"""Sidebar: New chat, Settings button, chat list.

Edit when: changing the sidebar.
"""

import streamlit as st

from chatppt.ui.common import ASSISTANT_AVATAR
from chatppt.ui.settings import open_settings
from chatppt.ui.state import create_chat, delete_chat


def render_sidebar() -> None:
    ss = st.session_state
    with st.sidebar:
        st.markdown(f"### {ASSISTANT_AVATAR} ChatPPT")
        if st.button("New chat", icon=":material/add:", key="new_chat_btn", width="stretch"):
            create_chat()
            st.rerun()
        if st.button("Settings", icon=":material/settings:", key="settings_btn", width="stretch"):
            open_settings()

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
