import streamlit as st

def setup_page():
    st.set_page_config(
        page_title="ORION | Operaciones Ropa",
        page_icon="📊",
        layout="wide"
    )

def render_header():
    st.markdown(
        """
        <div style="
            background-color:#3366CC;
            padding:18px;
            border-radius:12px;
            color:white;
            margin-bottom:20px;
        ">
            <h1 style="margin:0;">PRICE SHOES | OPERACIONES ROPA</h1>
            <h2 style="margin:0;">ORION</h2>
            <p style="margin:0;">Plataforma Indicadores de Recuperación de Mercancía</p>
        </div>
        """,
        unsafe_allow_html=True
    )
