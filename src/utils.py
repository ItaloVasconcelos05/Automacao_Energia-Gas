from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
import os

# Funções auxiliares de formatação: 
def adicionar_paragrafo_justificado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto justificado."""

    paragrafo = doc.add_paragraph(texto)
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Ajustar fonte se necessário (o padrão do python-docx é Calibri)
    # for run in paragraph.runs:
    #     run.font.name = 'Arial'
    #     run.font.size = Pt(tamanho_fonte)

def adicionar_texto_centralizado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto centralizado."""

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run(texto)
    run.bold = True
    
    # run.font.name = 'Arial'
    # run.font.size = Pt(tamanho_fonte)

def adicionar_titulo_secao(doc, texto):
    """Adiciona um título de seção formatado."""

    secao = doc.add_paragraph()
    secao.add_run(texto).bold = True

# Função para ajustar a largura das colunas 
def ajustar_largura_colunas(caminho_planilha):
    wb = load_workbook(caminho_planilha)
    ws = wb.active

    for coluna in ws.columns:
        max_length = 0
        coluna_letra = coluna[0].column_letter

        for celula in coluna: 
            try:
                if celula.value:
                    max_length = max(max_length, len(str(celula.value)))
            except:
                pass

        # Define largura da coluna com margem extra
        ajuste = max_length + 2
        ws.column_dimensions[coluna_letra].width = ajuste

    wb.save(caminho_planilha)

# Função para verificar se arquivo está em uso
def arquivo_em_uso(caminho):
    try:
        os.rename(caminho, caminho)
        return False
    except PermissionError:
        return True
    