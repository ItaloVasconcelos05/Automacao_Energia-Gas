from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook
import os
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image
import io


# Funções auxiliares de formatação:
def adicionar_paragrafo_justificado(doc, texto, tamanho_fonte=12):
    """Adiciona um parágrafo com texto justificado."""

    paragrafo = doc.add_paragraph(texto)
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW

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


def aplicar_estilo_texto(
    run, tamanho=12, negrito=False, fonte="Arial", cor_rgb=(0, 0, 0)
):
    run.font.name = fonte
    run._element.rPr.rFonts.set(qn("w:eastAsia"), fonte)
    run.font.size = Pt(tamanho)
    run.bold = negrito
    run.font.color.rgb = RGBColor(*cor_rgb)


def aplicar_borda_paragrafo(paragraph):
    p = paragraph._element
    pPr = p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    for border_name in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{border_name}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "2")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    pPr.append(borders)


def adicionar_legenda_formatada(doc, texto):
    par = doc.add_paragraph()
    run = par.add_run(texto)
    aplicar_estilo_texto(run, tamanho=10, fonte="Arial", cor_rgb=(90, 90, 90))
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    aplicar_borda_paragrafo(par)


def processar_imagem_para_relatorio(caminho_imagem, largura_max=1024, qualidade=80):
    # Abre a imagem
    img = Image.open(caminho_imagem)
    # Converte para RGB se necessário (evita problemas com PNG/transparência)
    if img.mode != "RGB":
        img = img.convert("RGB")
    # Redimensiona mantendo proporção
    if img.width > largura_max:
        proporcao = largura_max / float(img.width)
        altura_nova = int(float(img.height) * proporcao)
        img = img.resize((largura_max, altura_nova), Image.LANCZOS)
    # Salva em memória, sem metadados, com compressão JPEG
    buffer = io.BytesIO()
    img.save(buffer, format="JPEG", quality=qualidade, optimize=True)
    buffer.seek(0)
    return buffer
