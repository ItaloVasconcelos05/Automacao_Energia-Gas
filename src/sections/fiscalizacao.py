from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import *

def gerar_secao_fiscalizacao(doc, row, BASE_DIR):
    """
    Gera a seção '4. FISCALIZAÇÃO DAS INSTALAÇÕES DE GÁS' do relatório com base nos dados da PLANILHA DE FISCALIZAÇÕES.

    Parâmetros:
    - doc: objeto Document do python-docx.
    - row: linha da planilha contendo os dados da fiscalização.
    """
    
    # Adiciona o título da seção
    adicionar_paragrafo(
        doc,
        "4. FISCALIZAÇÃO DAS INSTALAÇÕES DE GÁS",
        tamanho_fonte=14,
        alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
        cor=(0, 0, 0),
        bold=True,
        estilo='Heading 2',
        espaco_depois=12,
        espaco_antes=12
    )
    
    