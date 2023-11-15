import streamlit as st
from streamlit_option_menu import option_menu
from control.control_dados import *



def pagHome():
    st.title("Conversor de Documentos para PDF")

    pasta_origem = st.text_input("Caminho da Pasta de Origem:", "musicas")
    pasta_destino = st.text_input("Caminho da Pasta de Destino:", "arquivos_pdf")

    if st.button("Converter para PDF"):
        converter_para_pdf(pasta_origem, pasta_destino)
        st.success("Conversão concluída com sucesso!")