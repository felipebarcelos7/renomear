import os
import pandas as pd
import logging
import fitz
from tkinter import Tk, filedialog

def configure_logging():
    log_file = "rename_log.txt"
    logging.basicConfig(
        filename=log_file,
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

def extract_text_from_pdf(pdf_file):
    try:
        with fitz.open(pdf_file) as pdf_document:
            page_text = ''
            for page_num in range(pdf_document.page_count):
                page = pdf_document.load_page(page_num)
                page_text += page.get_text()
        return page_text.lower()
    except Exception as e:
        logging.warning(f"Erro ao extrair texto do PDF: {e}")
        return ''

def select_excel_file():
    root = Tk()
    root.withdraw()
    excel_file = filedialog.askopenfilename(title="Selecione o arquivo Excel")
    return excel_file

def select_pdf_folder():
    root = Tk()
    root.withdraw()
    pdf_folder = filedialog.askdirectory(title="Selecione a pasta de destino dos PDFs")
    return pdf_folder

def rename_pdf_from_excel(excel_file, pdf_folder):
    # Carregar a planilha Excel usando pandas
    df = pd.read_excel(excel_file)

    # Configurar o logging
    configure_logging()

    # Criar listas para arquivos renomeados e não renomeados
    renamed_files = []
    not_renamed_files = []
    found_but_not_renamed = []

    # Iterar sobre os nomes na planilha
    for index, row in df.iterrows():
        name_to_search = row['Nome'].strip().lower()  # Certifique-se de substituir 'Nome' pelo nome correto da coluna na planilha

        # Variável para rastrear se encontramos o nome no PDF
        found_match = False

        # Iterar sobre os arquivos na pasta PDF
        for filename in os.listdir(pdf_folder):
            if filename.lower().endswith(".pdf"):
                # Ler o conteúdo do PDF e buscar o nome
                pdf_content = extract_text_from_pdf(os.path.join(pdf_folder, filename))
                if name_to_search in pdf_content:
                    new_filename = os.path.join(pdf_folder, f"{name_to_search}.pdf")
                    original_file = os.path.join(pdf_folder, filename)
                    try:
                        os.rename(original_file, new_filename)
                        logging.info(f"Arquivo renomeado: {filename} -> {name_to_search}")
                        renamed_files.append((filename, f"{name_to_search}.pdf"))
                        found_match = True
                        break
                    except Exception as e:
                        logging.warning(f"Erro ao renomear arquivo: {filename} - {e}")

        # Verificar se não encontramos uma correspondência e informar ao usuário
        if not found_match:
            logging.warning(f"Nome encontrado no PDF, mas arquivo não renomeado: {name_to_search}")
            found_but_not_renamed.append(name_to_search)

    print("Concluído!")
    print("Arquivos renomeados:", renamed_files)
    print("Arquivos não renomeados:", not_renamed_files)
    print("Nomes encontrados no PDF, mas não renomeados:", found_but_not_renamed)

# Exemplo de uso
if __name__ == "__main__":
    excel_file = select_excel_file()
    pdf_folder = select_pdf_folder()

    if excel_file and pdf_folder:
        rename_pdf_from_excel(excel_file, pdf_folder)
