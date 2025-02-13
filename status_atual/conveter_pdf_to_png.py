import fitz  # PyMuPDF
from dotenv import load_dotenv
import os

load_dotenv()
ROOT = os.getenv('ROOT')

def pdf_para_imagem(nome_arquivo):
    pdf_path = os.path.join(ROOT, 'arquivo_pdf', nome_arquivo)
    output_path = os.path.join(ROOT, 'arquivo_pdf', nome_arquivo.replace('.pdf', '.png'))

    try:
        # Abre o PDF e extrai a primeira página como imagem
        doc = fitz.open(pdf_path)
        page = doc.load_page(0)  # Carrega a primeira página (índice 0)
        pix = page.get_pixmap(dpi=300)  # Converte para imagem com DPI 300

        # Salva a imagem como PNG
        pix.save(output_path)
        print(f"PDF convertido e salvo como: {output_path}")

    except Exception as e:
        print(f"Erro ao converter PDF: {e}")

    finally:
        doc.close()


