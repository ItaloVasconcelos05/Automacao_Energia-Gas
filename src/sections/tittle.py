import os
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches
from utils import *

def gerar_titulo(doc, BASE_DIR):
    """
    Gera a página de título do relatório.
    """
    
    # Adiciona o cabeçalho e os títulos principais (se a imagem existir)
    logo_path = os.path.join(BASE_DIR, "assets/logo_arpe.png")
    if os.path.exists(logo_path):
        adicionar_imagem_no_cabecalho(
            doc,
            logo_path,
            largura=Inches(2),
            alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER
        )
    
    adicionar_paragrafo(
        doc,
        "Relatório de Fiscalização",
        tamanho_fonte=16,
        alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER,
        estilo='Title',
        espaco_antes=12,
        espaco_depois=6,
        cor=(0, 0, 0)
    )
    
    adicionar_paragrafo(
        doc,
        "Agência Reguladora de Pernambuco - ARPE",
        tamanho_fonte=12,
        alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER,
        estilo='Subtitle',
        espaco_antes=6,
        espaco_depois=12,
        cor=(0, 0, 0)
    )
    
    # Adiciona a imagem principal e centraliza (se existir)
    gas_path = os.path.join(BASE_DIR, "assets/gas.png")
    if os.path.exists(gas_path):
        doc.add_picture(gas_path, width=Inches(6))
        doc.paragraphs[-1].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    # Adiciona o título do relatório
    adicionar_paragrafo(
        doc,
        "FISCALIZAÇÃO DAS INSTALAÇÕES DE GÁS NOS MUNICÍPIOS DE CABO DE SANTO AGOSTINHO E JABOATÃO DOS GUARARAPES",
        bold=True,
        tamanho_fonte=10,
        estilo='Heading 1',
        alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER,
        espaco_antes=12,
        cor=(0, 0, 0)
    )
    
    # Agrupa e adiciona os nomes dos autores e a data
    autores_e_data = [
        "Argemiro Rivas",
        "Marta Rejane",
        "João Paulo Costa",
        "ABRIL/2025"
    ]
    
    for texto in autores_e_data:
        adicionar_paragrafo(
            doc,
            texto,
            tamanho_fonte=10,
            alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER,
            estilo='Heading 2',
            cor=(0, 0, 0)
        )
    
    # Adiciona as informações do processo administrativo
    informacoes_processo = [
        "RELATÓRIO DE FISCALIZAÇÃO DIRETA PROCESSO ADMINISTRATIVO",
        "PA-007/2025-CEEGC- GAS PROCESSOS",
        "SEI N° 0030200024.001385/2025-99"
    ]

    for texto in informacoes_processo:
        adicionar_paragrafo(
            doc,
            texto,
            tamanho_fonte=10,
            alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER,
            estilo='Heading 1',
            bold=True,
            cor=(0, 0, 0)
        )