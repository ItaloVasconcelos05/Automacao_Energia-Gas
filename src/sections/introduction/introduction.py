from docx import Document
from docx.enum.section import WD_SECTION
from utils import (
    adicionar_paragrafo_justificado,
    adicionar_titulo_secao,
)


def gerar_secao_introducao(doc: Document, row):
    """
    Gera a seção '1. INTRODUÇÃO' do relatório com base nos dados da PLANILHA DE FISCALIZAÇÕES.

    Parâmetros:
    - doc: objeto Document do python-docx.
    - row: linha da planilha contendo os dados da fiscalização.
    """

    adicionar_titulo_secao(doc, "1. INTRODUÇÃO")

    texto = (
        f"A Coordenadoria de Transportes e Rodovias da ARPE realizou vistoria no Terminal Rodoviário Intermunicipal "
        f"localizado no município de {row['Local']} no dia {row['Data']}, com o objetivo de verificar as condições operacionais, "
        f"de conservação, de manutenção e de segurança do referido terminal, em conformidade com o Contrato de Concessão "
        f"de Serviço Público nº {row['Contrato']}, firmado entre o Governo do Estado de Pernambuco e a empresa SOCICAM. "
        f"\n\nA visita técnica foi realizada pela equipe composta por {row['Pessoal Responsável']}, que analisou os aspectos físicos e funcionais "
        f"do terminal, incluindo áreas de embarque e desembarque, sanitários, vias internas, sinalização, segurança, infraestrutura "
        f"e atendimento ao usuário. Durante a fiscalização, foram registradas as não conformidades observadas, "
        f"as quais serão detalhadas nas seções seguintes deste relatório."
    )

    adicionar_paragrafo_justificado(doc, texto)
