import streamlit as st
import os
from docx2pdf import convert
import shutil

# Adiciona a chamada explícita do CoInitialize
import pythoncom
pythoncom.CoInitialize()

def importar_arquivos_word(arquivos, pasta_origem, pasta_destino):
    # Cria a pasta de destino se não existir
    if not os.path.exists(pasta_destino):
        os.makedirs(pasta_destino)

    progress_bar = st.progress(0)
    progress_text = st.empty()

    total_files = len(arquivos)
    completed_files = 0

    pdf_paths = []  # List to store the paths of the converted PDF files

    for arquivo in arquivos:
        # Salva cada arquivo Word na pasta de origem
        nome_arquivo_word = os.path.join(pasta_origem, arquivo.name)
        nome_arquivo_pdf = os.path.join(pasta_destino, f"{os.path.splitext(arquivo.name)[0]}.pdf")

        if os.path.exists(nome_arquivo_pdf):
            st.warning(f"O arquivo PDF para '{arquivo.name}' já existe em {nome_arquivo_pdf}. Pulando conversão.")
        else:
            with open(nome_arquivo_word, "wb") as arquivo_destino:
                arquivo_destino.write(arquivo.read())
            
            st.success(f"Arquivo Word '{arquivo.name}' importado com sucesso para {nome_arquivo_word}")

            try:
                # Converte cada arquivo Word para PDF
                convert(nome_arquivo_word, nome_arquivo_pdf)
                st.success(f"Arquivo PDF gerado com sucesso em {nome_arquivo_pdf}")
                pdf_paths.append(nome_arquivo_pdf)  # Store the path for downloading
            except Exception as e:
                st.error(f"Erro durante a conversão do arquivo '{arquivo.name}': {e}")

        completed_files += 1
        progress_value = completed_files / total_files
        progress_bar.progress(progress_value)
        progress_text.text(f"Progresso: {completed_files}/{total_files} arquivos convertidos.")

    # Show success message in a popup at the end
    st.success("Todos os arquivos foram convertidos com sucesso!")

    # Provide a download button for the user
    if st.button("Baixar Todos os Arquivos"):
        zip_filename = "converted_files.zip"
        with st.spinner(f"Preparando para baixar {len(pdf_paths)} arquivos..."):
            # Create a zip file containing all the PDF files
            with shutil.ZipFile(zip_filename, 'w') as zip_file:
                for pdf_path in pdf_paths:
                    zip_file.write(pdf_path, os.path.basename(pdf_path))

        st.success(f"Arquivo zipado '{zip_filename}' pronto para download!")
        st.markdown(f"**[Clique aqui para baixar o arquivo zipado]({zip_filename})**")

def main():
    st.title("Conversor - Word para PDF")

    # Componente de upload do Streamlit para múltiplos arquivos Word
    arquivos_upload = st.file_uploader("Escolha um ou mais arquivos Word (.docx)", type=["docx"], accept_multiple_files=True)

    if arquivos_upload is not None:
        # Botão para importar os arquivos
        if st.button("Converter Arquivos"):
            importar_arquivos_word(arquivos_upload, rf"files\arquivos_origem", rf"files\arquivos_pdf")

if __name__ == "__main__":
    main()
