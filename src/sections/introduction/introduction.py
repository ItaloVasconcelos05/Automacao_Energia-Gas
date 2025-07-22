from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run(
        "A Coordenadoria de Transportes e Rodovias da Arpe realizou vistoria nos Terminais Rodoviários Intermunicipais concedidos com o objetivo de verificar as condições operacionais, de conservação, de manutenção e de segurança dos referidos terminais, conforme Contrato de Serviço Público Nº "
    )
    par.add_run(str(row["Contrato"])).bold = True
    par.add_run(
        ", firmado entre o Governo do Estado, representado pela Secretaria de Transportes (SETRA) e a SOCICAM - Administração, Projetos e Representações Ltda. A ação foi no dia "
    )
    par.add_run(str(row["Data"])).bold = True
    par.add_run(", exclusivamente no Terminal de ")
    par.add_run(str(row["Local"])).bold = True
    par.add_run(
        " e nos dias 24 a 28 de março de 2025, nas cidades de Caruaru, Garanhuns, Arcoverde, Serra Talhada e Petrolina. As visitas técnicas foram realizadas pela equipe formada por "
    )
    par.add_run(str(row["Pessoal Responsável"])).bold = True
    par.add_run(
        ".\n\nNeste Relatório de Fiscalização foram observadas as condições de conservação, limpeza e higiene das áreas de embarque e desembarque, dos sanitários, as condições do pavimento das vias de circulação interna, a infraestrutura oferecida, os locais de estocagem de veículos, a segurança e o atendimento ao usuário, bem como toda estrutura para funcionamento dos terminais. A equipe da Arpe conversou com os responsáveis pelos seis terminais que forneceram informações complementares à fiscalização, principalmente sobre a implantação dos sistemas contra incêndio."
    )
