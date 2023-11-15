import streamlit as st
import os
from docx2pdf import convert
#import pythoncom

#pythoncom.CoInitialize()

def criar_pasta(pasta):
    # Cria a pasta se não existir
    if not os.path.exists(pasta):
        os.makedirs(pasta)

def importar_arquivo_word(arquivo, pasta_origem):
    # Salva o arquivo Word na pasta de origem
    nome_arquivo_word = os.path.join(pasta_origem, arquivo.name)
    with open(nome_arquivo_word, "wb") as arquivo_destino:
        arquivo_destino.write(arquivo.read())
    return nome_arquivo_word

def converter_para_pdf(nome_arquivo_word, pasta_destino):
    # Converte o arquivo Word para PDF
    nome_arquivo_pdf = os.path.join(pasta_destino, f"{os.path.splitext(os.path.basename(nome_arquivo_word))[0]}.pdf")
    convert(nome_arquivo_word, nome_arquivo_pdf)
    return nome_arquivo_pdf

def converter_arquivos(pasta_origem, pasta_destino):
    criar_pasta(pasta_destino)

    progress_bar = st.progress(0)
    progress_text = st.empty()

    arquivos = os.listdir(pasta_origem)
    total_files = len(arquivos)
    completed_files = 0

    for arquivo_nome in arquivos:
        try:
            nome_arquivo_word = os.path.join(pasta_origem, arquivo_nome)

            # Verifica se o arquivo PDF já existe
            nome_arquivo_pdf = f"{os.path.splitext(arquivo_nome)[0]}.pdf"
            if os.path.exists(os.path.join(pasta_destino, nome_arquivo_pdf)):
                st.warning(f"O arquivo PDF para '{arquivo_nome}' já existe em {os.path.join(pasta_destino, nome_arquivo_pdf)}. Pulando conversão.")
            else:
                converter_para_pdf(nome_arquivo_word, pasta_destino)
                st.success(f"Arquivo Word '{arquivo_nome}' convertido com sucesso para PDF em {os.path.join(pasta_destino, nome_arquivo_pdf)}")

        except Exception as e:
            st.error(f"Erro durante a conversão do arquivo '{arquivo_nome}': {e}")

        completed_files += 1
        progress_value = completed_files / total_files
        progress_bar.progress(progress_value)
        progress_text.text(f"Progresso: {completed_files}/{total_files} arquivos convertidos.")

    # Show success message in a popup at the end
    st.success("Todos os arquivos foram convertidos com sucesso!")

def main():
    st.title("Conversor - Word para PDF")

    # Componente de upload do Streamlit para múltiplos arquivos Word
    arquivos_upload = st.file_uploader("Escolha um ou mais arquivos Word (.docx)", type=["docx"], accept_multiple_files=True)

    # Define a pasta de origem e destino
    pasta_origem = rf"files\arquivos_origem"
    pasta_destino = rf"files\arquivos_pdf"

    # Realiza a importação e conversão ao carregar os arquivos
    if arquivos_upload is not None:
        for arquivo in arquivos_upload:
            try:
                nome_arquivo_word = importar_arquivo_word(arquivo, pasta_origem)
                st.success(f"Arquivo Word '{arquivo.name}' importado com sucesso para {nome_arquivo_word}")
            except Exception as e:
                st.error(f"Erro durante a importação do arquivo '{arquivo.name}': {e}")

        # Botão para converter arquivos
        if st.button("Converter Arquivos"):
            converter_arquivos(pasta_origem, pasta_destino)

if __name__ == "__main__":
    main()
