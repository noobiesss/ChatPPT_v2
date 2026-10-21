"""ChatPPT V2 entry point. Run with: streamlit run app.py

Streamlit page layout: wires state, CSS, sidebar, chat area, composer.
Edit when: changing the overall page flow.
"""

import streamlit as st

from chatppt.ui.browser_chats import sync_browser_chats
from chatppt.ui.common import LOGO_ICON_PATH
from chatppt.ui.composer import render_composer
from chatppt.ui.css import inject_css
from chatppt.ui.export import render_deck_block, render_deck_placeholder
from chatppt.ui.generation import run_pending
from chatppt.ui.messages import render_hero, render_messages
from chatppt.ui.sidebar import render_sidebar
from chatppt.ui.state import init_state


def main():
    st.set_page_config(
        page_title="ChatPPT V2",
        page_icon=LOGO_ICON_PATH,
        layout="wide",
        initial_sidebar_state="expanded",
    )
    init_state()
    chat = st.session_state.chats[st.session_state.active_chat]
    inject_css(chat["id"])
    render_sidebar()

    chat_area = st.container()
    with chat_area:
        if not chat["messages"] and chat["pending"] is None:
            render_hero()
        # The outline sits after the message that last changed the deck (not re-appended after every answer)
        deck = None
        if chat["slides"]:
            deck = (lambda: render_deck_block(chat)) if chat["pending"] is None else (lambda: render_deck_placeholder(chat))
        render_messages(chat, deck)

    render_composer(chat)

    # Saves chats to this browser. Must run AFTER the outline editor and the composer, so an edit made in
    # this run is already in the chat when it is fingerprinted (in the sidebar it would see the old data)
    sync_browser_chats()

    # The slow LLM job runs last, after the whole page (including the new user message) is already drawn
    if chat["pending"] is not None:
        run_pending(chat, chat_area)


if __name__ == "__main__":
    main()
