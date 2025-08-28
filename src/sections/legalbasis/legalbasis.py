from docx import Document
from utils import adicionar_titulo_secao, adicionar_paragrafo
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt


def gerar_secao_fundamentacao_legal(doc: Document):
    """
    Gera a seção '3. METODOLOGIA' no relatório.
    Esta seção contém texto fixo, baseado na legislação e regulamentação da ARPE.
    """

    adicionar_titulo_secao(doc, "3. METODOLOGIA")

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run(
        "A fiscalização direta e periódica do(s) município(s) de "
    )
    par.add_run("Caruaru").bold = True
    par.add_run(", ")
    par.add_run("Glória do Goitá").bold = True
    par.add_run(", ")
    par.add_run("Moreno").bold = True
    par.add_run(", ")
    par.add_run("Vitória de Santo Antão").bold = True
    par.add_run(" e ")
    par.add_run("Recife").bold = True
    par.add_run(
        " realizada por analistas da Coordenadoria de Energia Elétrica e Gás Canalizado da Arpe é submetida a uma metodologia que promova a qualidade e eficiência dos serviços prestados. Ela é organizada em três etapas: "
    )
    par.add_run("Preparação e Planejamento").bold = True
    par.add_run(", ")
    par.add_run("Execução da Fiscalização").bold = True
    par.add_run(" e ")
    par.add_run("Monitoramento e Avaliação").bold = True
    par.add_run(".\n")

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run("Preparação e Planejamento").bold = True
    par.add_run(
        ": Compreende a organização e estruturação das atividades preliminares à execução da fiscalização, perpassando pelos seguintes pontos: Levantamento e análise de Fiscalizações anteriores, Definição dos municípios a serem fiscalizados e Solicitação de um funcionário da Copergás para acompanhar a fiscalização."
    )

    par.add_run("Execução da Fiscalização").bold = True
    par.add_run(
        ": A execução da fiscalização é pautada por um arcabouço de normas e diretrizes, possibilitando que todos os equipamentos estejam em conformidade aos padrões estabelecidos:"
    )
    
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run("- ")
    par.add_run("Lei Federal n° 14.134, de 8 de abril de 2021").bold = True
    par.add_run(
        ", que dispõe sobre as atividades relativas ao transporte de gás natural, de que trata o art. 177 da Constituição Federal, e sobre as atividades de escoamento, tratamento, processamento, estocagem subterrânea, acondicionamento, liquefação, regaseificação e comercialização de gás natural; altera as Leis nºs 9.478, de 6 de agosto de 1997, e 9.847, de 26 de outubro de 1999; e revoga a Lei nº 11.909, de 4 de março de 2009, e dispositivo da Lei nº 10.438, de 26 de abril de 2002;"
    )
    
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run("- ")
    par.add_run("Lei Estadual nº 17.641, de 5 de janeiro de 2022").bold = True
    par.add_run(", que altera a ")
    par.add_run("Lei Estadual nº 15.900, de 11 de outubro de 2016").bold = True
    par.add_run(
        ", que estabelece as normas relativas à exploração direta, ou mediante concessão, dos serviços locais de gás canalizado no Estado de Pernambuco, com vistas ao desenvolvimento e expansão dos serviços de gás canalizado no Estado de Pernambuco;" 
    )

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run("- ")
    par.add_run("Lei Estadual nº 12.524, de 30 de dezembro de 2003").bold = True
    par.add_run(
        " que altera e consolida as disposições da Lei nº 12.126, de 12/12/2001, que criou a Agência de Regulação dos Serviços Públicos Delegados do Estado de Pernambuco – Arpe: "
    )

    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = par.add_run("Art. 3º")
    run.font.size = Pt(8)
    run = par.add_run(
        " Compete à ARPE a regulação de todos os serviços públicos delegados pelo Estado de Pernambuco, ou por ele diretamente prestados, embora sujeitos à delegação, quer de sua competência ou a ele delegados por outros entes federados, em decorrência de norma legal ou regulamentar, disposição convenial ou contratual.\n\n")
    run.font.size = Pt(8)
    run = par.add_run("§ 1º A atividade reguladora da ARPE deverá ser exercida, em especial, nas seguintes áreas:\n")
    run.font.size = Pt(8)
    run = par.add_run("(...);\n")
    run.font.size = Pt(8)
    run = par.add_run("II - energia elétrica;")
    run.bold = True
    run.font.size = Pt(8)
    run = par.add_run("(...);\n")
    run.font.size = Pt(8)
    run = par.add_run("VI - distribuição de gás canalizado;")
    run.bold = True
    run.font.size = Pt(8)
    run = par.add_run("(...);\n")
    run.font.size = Pt(8)
    run = par.add_run("Art. 4º Compete ainda à Arpe:")
    run.font.size = Pt(8)
    run = par.add_run("(...);\n")
    run.font.size = Pt(8)
    run = par.add_run("X - fiscalizar diretamente ou mediante convênio com o Estado de Pernambuco, através de seus órgãos ou entidades vinculadas, com sua supervisão, os aspectos técnico, econômico, contábil, financeiro, operacional e jurídico dos serviços públicos delegados, valendo-se inclusive, de indicadores e procedimentos amostrais.;\n")
    run.font.size = Pt(8)
    


    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run("- ")
    par.add_run(
        "Resolução da Arpe nº 034, de 10 de agosto de 2006"
    ).bold = True
    par.add_run(
        " - Dispõe sobre a prestação do serviço de fornecimento de gás canalizado no Estado de Pernambuco, estabelecendo procedimentos e indicadores de segurança e qualidade a serem adotados pela Companhia Pernambucana de Gás - COPERGÁS, estabelece penalidades e dá outras providências;"
    )
    
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    par.add_run("- ")
    par.add_run(
        "Resolução nº 083, de 30 de julho de 2013"
    ).bold = True
    par.add_run(
        " , que dispõe sobre os procedimentos de fiscalização, autuação e aplicação de penalidades aos prestadores de serviços públicos delegados no Estado de Pernambuco e aos serviços públicos fiscalizados pela Arpe mediante delegação:\n\n"
    )
    run = par.add_run("Art. 1º. Regulamentar os procedimentos de fiscalização, autuação e aplicação de penalidades aos prestadores de serviços públicos delegados no Estado de Pernambuco.\n")
    run.font.size = Pt(8)
    run = par.add_run("(...);\n")
    run.font.size = Pt(8)
    
    par.add_run("- ")
    par.add_run(
        "Norma técnica da ABNT NBR 12.712"
    ).bold = True
    par.add_run(
        " – Projeto de sistemas de transmissão e distribuição de gás combustível;\n\n"
    )
    
    par.add_run("- ")
    par.add_run(
        "Norma técnica da ABNT NBR 15.526"
    ).bold = True
    par.add_run(
        " Redes de distribuição interna para gases combustíveis em instalações residenciais - Projeto e execução.\n\n"
    )