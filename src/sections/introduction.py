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
    
    # Adiciona a lista de abreviaturas e siglas
    adicionar_paragrafo(
        doc,
        "LISTA DE ABREVIATURAS E SIGLAS",
        tamanho_fonte=13,
        alinhamento=WD_ALIGN_PARAGRAPH.CENTER,
        cor=(0, 0, 0),
        bold=True,
        estilo='Heading 3',
        espaco_depois=6,
        espaco_antes=12
    )

    dados_siglas = [
        ('Sigla', 'Definição'),
        ('CRM', 'Conjunto de Regulagem de Pressão e Medição'),
        ('ERP', 'Estação de Regulagem de Pressão'),
        ('ERPM', 'Estação de Regulagem, Pressão e Medição'),
        ('ETC', 'Estação de Transferência de Custódia'),
        ('GNV', 'Gás Natural Veicular')
    ]
    adicionar_tabela(
        doc,
        dados=dados_siglas,
        largura_colunas=[1, 4],
        alinhamento_colunas=[WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.CENTER]
    )
    
    # Adiciona o sumário
    adicionar_paragrafo(
        doc,
        "SUMÁRIO",
        tamanho_fonte=14,
        alinhamento=WD_ALIGN_PARAGRAPH.CENTER,
        bold=True,
        estilo='Heading 3',
        espaco_antes=25,
        espaco_depois=6,
        cor=(0, 0, 0)
    )

    # Lista de tópicos do sumário
    topicos_sumario = [
        "1. INTRODUÇÃO",
        "2. OBJETIVOS",
        "3. METODOLOGIA",
        "4. FISCALIZAÇÃO",
        "    4.1. PREPARAÇÃO E PLANEJAMENTO",
        "    4.2. EXECUÇÃO DA FISCALIZAÇÃO",
        "    4.3. MONITORAMENTO E AVALIAÇÃO",
        "5. DETERMINAÇÕES GERAIS",
        "APÊNDICE 1 - FOTOS DAS NÃO CONFORMIDADES",
        "APÊNDICE 2 - ANÁLISE DAS FISCALIZAÇÕES"
    ]
    for topico in topicos_sumario:
        adicionar_paragrafo(doc, topico, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.LEFT)

    # Adiciona o título e os parágrafos da introdução
    adicionar_paragrafo(
        doc,
        "1. INTRODUÇÃO",
        tamanho_fonte=14,
        alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
        bold=True,
        estilo='Heading 3',
        espaco_antes=25,
        espaco_depois=6,
        cor=(0, 0, 0)
    )
    
    texto_introducao = (
        "Atualmente, a prestação dos serviços públicos de odorização, canalização e distribuição de gás natural em Pernambuco é realizada pela Companhia Pernambucana de Gás (Copergás). Diante das transformações regulatórias e desafios operacionais, a Agência de Regulação dos Serviços Públicos Delegados do Estado de Pernambuco (Arpe), por meio da Coordenadoria de Energia Elétrica e Gás Canalizado (CEEGC), conduz fiscalizações e procedimentos administrativos voltados à regulação técnico-operacional dos serviços prestados pela Copergás. O objetivo dessas atividades é avaliar as condições operacionais, a conservação e a manutenção das instalações de gás, além de verificar a conformidade com a legislação vigente, a qualidade do serviço prestado e a satisfação dos usuários."
    )
    
    segundo_paragrafo = (
        "Nesse contexto, as Fiscalizações Periódicas, organizadas dentro da Agenda Regulatória da CEEGC, têm como propósito inspecionar se as instalações do sistema de distribuição de gás natural atendem às normas legais, garantindo a adequação e a padronização dos serviços prestados. Este relatório apresenta os resultados das mais recentes fiscalizações realizadas in loco nos municípios de Cabo de Santo Agostinho, Camaragibe, Goiana, Igarassu, Ipojuca, Itapissuma e São Lourenço Da Mata, durante o mês de julho de 2025."
    )
    
    adicionar_paragrafo(doc, texto_introducao, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=6)
    adicionar_paragrafo(doc, segundo_paragrafo, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=12)