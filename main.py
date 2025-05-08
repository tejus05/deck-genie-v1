import streamlit as st
from ui import render_ui

def main():
    st.set_page_config(
        page_title="Deck Genie - B2B SaaS Presentation Generator",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="collapsed"
    )
    
    render_ui()

if __name__ == "__main__":
    main()