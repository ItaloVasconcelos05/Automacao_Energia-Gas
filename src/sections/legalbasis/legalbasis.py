from docx import Document
from utils import *
from docx.enum.text import WD_ALIGN_PARAGRAPH


def gerar_secao_fundamentacao_legal(doc: Document):
    """
    Gera a seção '3. METODOLOGIA' no relatório.
    Esta seção contém texto fixo, baseado na legislação e regulamentação da ARPE.
    """

    adicionar_titulo_secao(doc, "3. METODOLOGIA")

    adicionar_paragrafo(
        doc,
        "A fiscalização direta e periódica do(s) município(s) de "
        "Caruaru, Glória do Goitá, Moreno, Vitória de Santo Antão e Recife "
        "realizada por analistas da Coordenadoria de Energia Elétrica e Gás Canalizado da Arpe é submetida a uma metodologia que promova a qualidade e eficiência dos serviços prestados. Ela é organizada em três etapas: "
        "Preparação e Planejamento, Execução da Fiscalização e Monitoramento e Avaliação.",
      bold=["Caruaru", "Glória do Goitá", "Moreno", "Vitória de Santo Antão", "Recife", "Preparação e Planejamento", "Execução da Fiscalização", "Monitoramento e Avaliação"],
    alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY,
    espaco_antes=12, espaco_depois=12
    )

    adicionar_paragrafo(
        doc,
        "Preparação e Planejamento: Compreende a organização e estruturação das atividades preliminares à execução da fiscalização, perpassando pelos seguintes pontos: Levantamento e análise de Fiscalizações anteriores, Definição dos municípios a serem fiscalizados e Solicitação de um funcionário da Copergás para acompanhar a fiscalização.",
    bold=["Preparação e Planejamento"],
    alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY,
    estilo = "List Number",
    espaco_depois=12
    )

    adicionar_paragrafo(
        doc,
        "Execução da Fiscalização: A execução da fiscalização é pautada por um arcabouço de normas e diretrizes, possibilitando que todos os equipamentos estejam em conformidade aos padrões estabelecidos:",
    bold=["Execução da Fiscalização"],
    alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY,
    estilo = "List Number",
    espaco_depois=12
    )

    adicionar_paragrafo(
        doc,
        "Lei Federal n° 14.134, de 8 de abril de 2021, que dispõe sobre as atividades relativas ao transporte de gás natural, de que trata o art. 177 da Constituição Federal, e sobre as atividades de escoamento, tratamento, processamento, estocagem subterrânea, acondicionamento, liquefação, regaseificação e comercialização de gás natural; altera as Leis nºs 9.478, de 6 de agosto de 1997, e 9.847, de 26 de outubro de 1999; e revoga a Lei nº 11.909, de 4 de março de 2009, e dispositivo da Lei nº 10.438, de 26 de abril de 2002;\n",
    bold=["Lei Federal n° 14.134, de 8 de abril de 2021"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo = "List Bullet",
    espaco_depois=8
    )

    adicionar_paragrafo(
        doc,
        "Lei Estadual nº 17.641, de 5 de janeiro de 2022, que altera a Lei Estadual nº 15.900, de 11 de outubro de 2016, que estabelece as normas relativas à exploração direta, ou mediante concessão, dos serviços locais de gás canalizado no Estado de Pernambuco, com vistas ao desenvolvimento e expansão dos serviços de gás canalizado no Estado de Pernambuco;\n",
    bold=["Lei Estadual nº 17.641, de 5 de janeiro de 2022", "Lei Estadual nº 15.900, de 11 de outubro de 2016"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo = "List Bullet",
    espaco_depois=8
    )

    adicionar_paragrafo(
        doc,
        "Lei Estadual nº 12.524, de 30 de dezembro de 2003 que altera e consolida as disposições da Lei nº 12.126, de 12/12/2001, que criou a Agência de Regulação dos Serviços Públicos Delegados do Estado de Pernambuco – Arpe: ",
    bold=["Lei Estadual nº 12.524, de 30 de dezembro de 2003", "Lei nº 12.126, de 12/12/2001"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo = "List Bullet",
    espaco_depois=8
    )

    adicionar_paragrafo(
        doc,
        "Art. 3º Compete à ARPE a regulação de todos os serviços públicos delegados pelo Estado de Pernambuco, ou por ele diretamente prestados, embora sujeitos à delegação, quer de sua competência ou a ele delegados por outros entes federados, em decorrência de norma legal ou regulamentar, disposição convenial ou contratual.\n\n"
        "§ 1º A atividade reguladora da ARPE deverá ser exercida, em especial, nas seguintes áreas:\n"
        "(...);\n"
        "II - energia elétrica;\n(...);\n"
        "VI - distribuição de gás canalizado;\n(...);\n"
        "Art. 4º Compete ainda à Arpe:\n(...);\n"
        "X - fiscalizar diretamente ou mediante convênio com o Estado de Pernambuco, através de seus órgãos ou entidades vinculadas, com sua supervisão, os aspectos técnico, econômico, contábil, financeiro, operacional e jurídico dos serviços públicos delegados, valendo-se inclusive, de indicadores e procedimentos amostrais.;\n",
    bold=["II - energia elétrica", "VI - distribuição de gás canalizado",],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo="Quote",
    espaco_depois=12,
    tamanho_fonte=10
    )

    adicionar_paragrafo(
        doc,
        "Resolução da Arpe nº 034, de 10 de agosto de 2006 - Dispõe sobre a prestação do serviço de fornecimento de gás canalizado no Estado de Pernambuco, estabelecendo procedimentos e indicadores de segurança e qualidade a serem adotados pela Companhia Pernambucana de Gás - COPERGÁS, estabelece penalidades e dá outras providências;\n",
    bold=["Resolução da Arpe nº 034, de 10 de agosto de 2006"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo = "List Bullet",
    espaco_depois=8
    )
    
    adicionar_paragrafo(
        doc,
        "Resolução nº 083, de 30 de julho de 2013, que dispõe sobre os procedimentos de fiscalização, autuação e aplicação de penalidades aos prestadores de serviços públicos delegados no Estado de Pernambuco e aos serviços públicos fiscalizados pela Arpe mediante delegação:",
    bold=["Resolução nº 083, de 30 de julho de 2013,"],
    alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY,
    estilo = "List Bullet",
    espaco_depois=6
    )
    adicionar_paragrafo(
        doc,
        "Art. 1º. Regulamentar os procedimentos de fiscalização, autuação e aplicação de penalidades aos prestadores de serviços públicos delegados no Estado de Pernambuco.Art. 1º. Regulamentar os procedimentos de fiscalização, autuação e aplicação de penalidades aos prestadores de serviços públicos delegados no Estado de Pernambuco",
    alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY,
    estilo= "Quote",
    espaco_depois=12,
    tamanho_fonte=10
    )

    adicionar_paragrafo(
        doc,
        "Norma técnica da ABNT NBR 12.712 – Projeto de sistemas de transmissão e distribuição de gás combustível;",
    bold=["Norma técnica da ABNT NBR 12.712"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo= "List Bullet",
    espaco_depois= 8
    )
    adicionar_paragrafo(
        doc,
        "Norma técnica da ABNT NBR 15.526 Redes de distribuição interna para gases combustíveis em instalações residenciais - Projeto e execução.",
    bold=["Norma técnica da ABNT NBR 15.526"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo= "List Bullet",
    espaco_depois=8
    )
    adicionar_paragrafo(
        doc,
        "Monitoramento e Avaliação: Após a execução da fiscalização, seguem os trâmites pertinentes as Resoluções Arpe. Esta etapa é fundamental para garantir a eficácia das ações corretivas e a melhoria contínua dos serviços prestados. Os principais pontos do Monitoramento e Avaliação são: Termo de Notificação e Relatório de Fiscalização, Plano de Ação e Análise de Indicadores e Relatórios de Acompanhamento e Avaliação Final.\n",
    bold=["Monitoramento e Avaliação:"],
    alinhamento=WD_ALIGN_PARAGRAPH.LEFT,
    estilo= "List Number",
    espaco_depois=6
    )

 