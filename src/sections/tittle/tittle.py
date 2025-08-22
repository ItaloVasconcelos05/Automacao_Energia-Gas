import os
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches
from docx.shared import Pt # Import Pt for point units
from utils import *

def gerar_titulo(doc, BASE_DIR):
    adicionar_imagem_no_cabecalho(doc, os.path.join(BASE_DIR, "assets/logo_arpe.png"), largura=Inches(2), alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER)

    adicionar_paragrafo(doc, "Relatório de Fiscalização", tamanho_fonte=16, alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER, estilo='Title' ,espaco_antes=12, espaco_depois=6, cor=(0,0,0))
    adicionar_paragrafo(doc, "Agência Reguladora de Pernambuco - ARPE", tamanho_fonte=12, alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER, estilo='Subtitle', espaco_antes=6, espaco_depois=12, cor=(0,0,0))
    
    doc.add_picture(os.path.join(BASE_DIR, "assets/gas.png"), width=Inches(6))
    logo_gas = doc.paragraphs[-1]
    logo_gas.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    
    adicionar_paragrafo(doc, "FISCALIZAÇÃO DAS INSTALAÇÕES DE GÁS NOS MUNICÍPIOS DE CABO DE SANTO AGOSTINHO E JABOATÃO DOS GUARARAPES", bold=True, tamanho_fonte=10, estilo='Heading 1', alinhamento=WD_PARAGRAPH_ALIGNMENT.CENTER, espaco_antes=12, cor=(0,0,0))
    adicionar_paragrafo(doc, "Argemiro Rivas", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 2', cor=(0,0,0))
    adicionar_paragrafo(doc, "Marta Rejane", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 2', cor=(0,0,0))
    adicionar_paragrafo(doc, "João Paulo Costa", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 2', cor=(0,0,0))
    adicionar_paragrafo(doc, "Alexandre Almeida", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 2', cor=(0,0,0))
    adicionar_paragrafo(doc, "ABRIL/2025", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 2', espaco_depois=20, cor=(0,0,0))

    adicionar_paragrafo(doc, "RELATÓRIO DE FISCALIZAÇÃO DIRETA PROCESSO ADMINISTRATIVO", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 1', bold=True, espaco_antes=20, cor=(0,0,0))
    adicionar_paragrafo(doc, "PA-007/2025-CEEGC- GAS PROCESSOS", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 1', bold=True, cor=(0,0,0))
    adicionar_paragrafo(doc, "SEI N° 0030200024.001385/2025-99", 10, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, estilo='Heading 1', bold=True, cor=(0,0,0))