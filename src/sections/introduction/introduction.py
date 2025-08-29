from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import *
import os


def gerar_secao_introducao(doc: Document, row, BASE_DIR):
    """
    Gera a seção '1. INTRODUÇÃO' do relatório com base nos dados da PLANILHA DE FISCALIZAÇÕES.

    Parâmetros:
    - doc: objeto Document do python-docx.
    - row: linha da planilha contendo os dados da fiscalização.
    """
    
    #ABREVIATURAS E SIGLAS
    adicionar_paragrafo(doc, "LISTA DE ABREVIATURAS E SIGLAS", tamanho_fonte=13, alinhamento=WD_ALIGN_PARAGRAPH.CENTER,cor=(0,0,0), bold=True, estilo='Heading 3', espaco_depois=6, espaco_antes=12)
    
    dados_produto = [
    ('Sigla', 'Definição'),
    ('CRM', 'Conjunto de Regulagem de Pressão e Medição'),
    ('ERP', 'Estação de Regulagem de Pressão'),
    ('ERPM', 'Estação de Regulagem, Pressão e Medição'),
    ('ETC','Estação de Transferência de Custódia'),
    ('GNV','Gás Natural Veicular')
    ]

    larguras_produto = [1, 4]
    alinhamentos_produto = [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]

    adicionar_tabela(doc, 
                    dados=dados_produto, 
                    largura_colunas=larguras_produto,
                    alinhamento_colunas=alinhamentos_produto
                    )
            
            
    # Adiciona o título "SUMÁRIO"
    adicionar_paragrafo(doc, "SUMÁRIO", tamanho_fonte=14, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
    # SUMÁRIO
    adicionar_paragrafo(doc, "SUMÁRIO", tamanho_fonte=14, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, bold=True, estilo='Heading 3', espaco_antes=25, espaco_depois=6, cor=(0,0,0))
    adicionar_paragrafo(doc, "1. INTRODUÇÃO", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "2. OBJETIVOS", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "3. METODOLOGIA", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "4. FISCALIZAÇÃO", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "   4.1. PREPARAÇÃO E PLANEJAMENTO", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "   4.2. EXECUÇÃO DA FISCALIZAÇÃO", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "   4.3. MONITORAMENTO E AVALIAÇÃO", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "5. DETERMINAÇÕES GERAIS", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "APÊNDICE 1 - FOTOS DAS NÃO CONFORMIDADES", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    adicionar_paragrafo(doc, "APÊNDICE 2 - ANÁLISE DAS FISCALIZAÇÕES", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)
    
