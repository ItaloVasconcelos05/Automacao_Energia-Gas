from docx import Document
from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado


def gerar_secao_fundamentacao_legal(doc: Document):
    """
    Gera a seção '2. FUNDAMENTAÇÃO LEGAL' no relatório.
    Esta seção contém texto fixo, baseado na legislação e regulamentação da ARPE.
    """

    adicionar_titulo_secao(doc, "2. FUNDAMENTAÇÃO LEGAL")

    texto = (
        "A presente fiscalização encontra fundamento nas seguintes normas legais e regulamentares:\n\n"
        "- Lei nº 12.524, de 30 de dezembro de 2003 – Altera e consolida as disposições da Lei nº 12.126, de 12 de dezembro de 2001, "
        "que cria a Agência de Regulação dos Serviços Públicos Delegados do Estado de Pernambuco – ARPE, regulamentada pelo "
        "Decreto nº 30.200, de 09 de fevereiro de 2007.\n\n"
        "- Lei nº 13.254, de 21 de junho de 2007 e alterações, em especial a Lei Estadual nº 15.200, de 17 de dezembro de 2013 – "
        "Estrutura o Sistema de Transporte Coletivo Intermunicipal de Passageiros do Estado de Pernambuco, regulamentada pelo "
        "Decreto nº 40.559, de 31 de março de 2014.\n\n"
        "- Resolução Arpe nº 46, de 07 de abril de 2008 (Antiga nº 06/2008) – Aprova o Regulamento dos Terminais Rodoviários do "
        "Estado de Pernambuco, alterada parcialmente pela Resolução ARPE nº 53, de 26 de janeiro de 2009 (Antiga 003/2009).\n\n"
        "- Resolução Arpe nº 083, de 30 de julho de 2013 – Dispõe sobre os procedimentos de fiscalização, autuação e aplicação de "
        "penalidades aos prestadores de serviços públicos delegados no Estado de Pernambuco fiscalizados pela ARPE mediante delegação.\n\n"
        "- Contrato de Concessão de Serviço Público nº 1.041.080/08, de 19 de setembro de 2008, e seus aditivos, especialmente o "
        "Segundo Termo Aditivo de 29 de setembro de 2017 – contrato celebrado entre o Estado de Pernambuco, representado pela "
        "Secretaria de Transportes – SETRA, e a SOCICAM – Administração, Projetos e Representações Ltda."
    )

    adicionar_paragrafo_justificado(doc, texto)
