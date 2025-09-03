from docx import Document
from utils import *
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
import os


def gerar_secao_fiscalizacoes(doc: Document, row, BASE_DIR):
    """
    Gera a seção '4. FISCALIZAÇÕES' do relatório.
    """

    # Título da seção
    adicionar_paragrafo(
        doc,
        "4. FISCALIZAÇÕES",
        tamanho_fonte=14,
        alinhamento=WD_ALIGN_PARAGRAPH.CENTER,
        bold=True,
        estilo='Heading 3',
        espaco_antes=25,
        espaco_depois=6,
        cor=(0, 0, 0),
    )
    adicionar_paragrafo(doc, "O processo de fiscalização pela Coordenadoria de Energia Elétrica e Gás Canalizado da Arpe é detalhado e sistemático, e neste item estão consolidadas as principais informações do processo, abordando os seguintes subitens: Preparação e Planejamento, Execução da Fiscalização e Monitoramento e Avaliação.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, bold =["Preparação e Planejamento", "Execução da Fiscalização", "Monitoramento e Avaliação"], espaco_depois=6)

    # Preparação e Planejamento
    adicionar_paragrafo(doc, "4.1. PREPARAÇÃO E PLANEJAMENTO", estilo='Heading 2', espaco_antes=12, espaco_depois=6)
    
    adicionar_paragrafo(doc, "Conforme destacado no Item 3, esta é uma etapa preliminar a execução da fiscalização, onde foram desenvolvidas as seguintes ações: Levantamento e análise de Fiscalizações anteriores e Definição dos municípios a serem fiscalizados e Solicitação de um funcionário da Copergás para acompanhar a fiscalização.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=12)
    
    adicionar_paragrafo(doc, "Análise de Fiscalizações anteriores: De acordo com o Planejamento Regulatório da ARPE para o setor de Gás Canalizado, foi estabelecida a meta de 372 fiscalizações. A partir dessa diretriz, a equipe da CEEGC definiu como objetivo secundário inspecionar todos os municípios do estado que possuem instalações da Copergás.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=["Análise de Fiscalizações anteriores:"], espaco_depois=12, estilo = "List Bullet")
    
    adicionar_paragrafo(doc, "Para isso, a meta foi distribuída de forma ponderada entre 26 municípios, considerando que aqueles com maior número de clientes deveriam ser fiscalizados com maior frequência. No entanto, essa distribuição passou por ajustes, já que em alguns municípios há um número muito reduzido de clientes (entre 1 e 3) e, nessas localidades, não faria sentido fiscalizar apenas um cliente, sendo necessário abranger todos. Além disso, três dos 26 municípios concentram 84,19% dos clientes da Copergás, o que reforça a necessidade de uma estratégia de fiscalização mais direcionada e eficiente.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=12)
    
    adicionar_paragrafo(doc, "Assim, a Tabela 01 apresenta a distribuição da meta de fiscalização por município. Na Figura 01, é exibido o mapa do estado de Pernambuco, destacando os municípios que devem ser fiscalizados, aqueles que já atingiram a quantidade necessária de fiscalizações e os que já foram inspecionados, mas ainda necessitam de novas fiscalizações.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=12)
    
    

    adicionar_paragrafo(doc, "4.2. EXECUÇÃO DA FISCALIZAÇÃO", estilo='Heading 2')
    
    adicionar_paragrafo(doc, "As Tabelas 02, 03, 04, 05 e 06 apresentam os clientes fiscalizados em cada um dos dias que foram realizadas fiscalizações no mês, junto ao membros da equipe da CEEGC que realizaram a fiscalização e o funcionário da Copergás que os acompanharam. As Não Conformidades (NC) constatadas in loco estão relacionadas na Tabela 07, e os seus registros fotográficos estão no Apêndice 1.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY)
    
    
    caminho_arquivo = os.path.join(BASE_DIR, "assets", "Acompanhamento_fiscaliz_2025.xlsx")

    try:
        df_tabela = pd.read_excel(
            caminho_arquivo,
            sheet_name="acompanhamento meta cidade",
            usecols="L:N",
            header=0,
        )
        print(df_tabela.head())
    except FileNotFoundError:
        print(f"Erro: O arquivo '{caminho_arquivo}' não foi encontrado.")
    except ValueError as e:
        print(f"Erro ao ler o arquivo: {e}")


    adicionar_paragrafo(doc, "4.3. MONITORAMENTO E AVALIAÇÃO", estilo='Heading 2')
    adicionar_paragrafo(doc, "Foram consolidadas evidências, registros e análises de conformidade para emissão deste relatório.", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY)


