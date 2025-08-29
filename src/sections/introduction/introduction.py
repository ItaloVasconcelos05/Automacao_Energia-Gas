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
    adicionar_paragrafo(doc, "SUMÁRIO", tamanho_fonte=14, espaco_antes=25, alinhamento=WD_ALIGN_PARAGRAPH.CENTER, bold=True)
    
<<<<<<< HEAD
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
=======
    # Adiciona os itens principais
    adicionar_paragrafo(doc, "1. INTRODUÇÃO", estilo='Heading 1', cor=(0,0,0))
    adicionar_paragrafo(doc, "2. OBJETIVOS", estilo='Heading 1', cor=(0,0,0))
    adicionar_paragrafo(doc, "3. METODOLOGIA", estilo='Heading 1', cor=(0,0,0))
    adicionar_paragrafo(doc, "4. FISCALIZAÇÃO", estilo='Heading 1', cor=(0,0,0))
    
    # Adiciona os subitens com indentação
    # Para indentar, você precisa definir um estilo de parágrafo diferente ou
    # ajustar a indentação do parágrafo após criá-lo.
    
    p_4_1 = adicionar_paragrafo(doc, "4.1 Preparação e Planejamento", estilo='Heading 2', cor=(0,0,0))
    p_4_2 = adicionar_paragrafo(doc, "4.2 Execução da Fiscalização", estilo='Heading 2', cor=(0,0,0))
    p_4_3 = adicionar_paragrafo(doc, "4.3 Monitoramento e Avaliação", estilo='Heading 2', cor=(0,0,0))
    
    # Adiciona os itens finais
    adicionar_paragrafo(doc, "5. DETERMINAÇÕES GERAIS", estilo='Heading 1', cor=(0,0,0))
    adicionar_paragrafo(doc, "APÊNDICE 1 - FOTOS DAS NÃO CONFORMIDADES", estilo='Heading 1', cor=(0,0,0))
    adicionar_paragrafo(doc, "APÊNDICE 2 - ANÁLISE DAS FISCALIZAÇÕES", estilo='Heading 1', espaco_depois=12, cor=(0,0,0))

    
    # Introdução
    adicionar_paragrafo(doc, "1. INTRODUÇÃO", tamanho_fonte=14, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, bold=True, estilo='Heading 3', espaco_antes=25, espaco_depois=6, cor=(0,0,0))
    texto_introducao = "Atualmente, a prestação dos serviços públicos de odorização, canalização e distribuição de gás natural em Pernambuco é realizada pela Companhia Pernambucana de Gás (Copergás). Diante das transformações regulatórias e desafios operacionais, a Agência de Regulação dos Serviços Públicos Delegados do Estado de Pernambuco (Arpe), por meio da Coordenadoria de Energia Elétrica e Gás Canalizado (CEEGC), conduz fiscalizações e procedimentos administrativos voltados à regulação técnico-operacional dos serviços prestados pela Copergás. O objetivo dessas atividades é avaliar as condições operacionais, a conservação e a manutenção das instalações de gás, além de verificar a conformidade com a legislação vigente, a qualidade do serviço prestado e a satisfação dos usuários."
    
    
    segundo_paragrafo = "Nesse contexto, as Fiscalizações Periódicas, organizadas dentro da Agenda Regulatória da CEEGC, têm como propósito inspecionar se as instalações do sistema de distribuição de gás natural atendem às normas legais, garantindo a adequação e a padronização dos serviços prestados. Este relatório apresenta os resultados das mais recentes fiscalizações realizadas in loco nos municípios de Cabo de Santo Agostinho, Camaragibe, Goiana, Igarassu, Ipojuca, Itapissuma e São Lourenço Da Mata, durante o mês de julho de 2025."

    # Parágrafos da introdução
    adicionar_paragrafo(doc, texto_introducao, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=6)
    adicionar_paragrafo(doc, segundo_paragrafo, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=12)

>>>>>>> 6f6eba5994234751fff6f3aa806da66c4d28a03c
    
