from utils import adicionar_titulo_secao, adicionar_paragrafo_justificado, adicionar_texto_centralizado

def gerar_secao_consideracoes_finais(doc, row):
    """
    Gera a seção '5. CONSIDERAÇÕES FINAIS' do relatório.

    Parâmetros:
    - doc: objeto Document.
    - row: linha da fiscalização (Series do DataFrame fiscalizacoes_df).
    """

    adicionar_titulo_secao(doc, "5. CONSIDERAÇÕES FINAIS")

    texto1 = (
        "Diante das constatações apontadas no presente Relatório de Fiscalização, solicita-se o envio deste documento à "
        "concessionária responsável, para que sejam adotadas as providências necessárias à regularização das não conformidades "
        "identificadas, bem como a apresentação dos respectivos prazos de conclusão dos serviços e/ou obras."
    )

    texto2 = (
        "Recomenda-se que a concessionária mantenha esta Agência informada acerca da adequação dos sistemas de segurança obrigatórios, "
        "como o Sistema de Proteção contra Descargas Atmosféricas (SPDA) e o Sistema de Combate a Incêndio, além da correta manutenção "
        "das áreas de uso comum, equipamentos e estruturas prediais dos terminais concedidos."
    )

    texto3 = (
        "Sugere-se o encaminhamento deste relatório à Empresa Pernambucana de Transporte Intermunicipal – EPTI, na qualidade de "
        "Poder Concedente e gestora do Sistema de Transporte Coletivo Intermunicipal de Passageiros do Estado de Pernambuco (STCIP-PE)."
    )

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)
    adicionar_paragrafo_justificado(doc, texto3)

    adicionar_texto_centralizado(doc, f"\n\nRecife, {row['Data']}.")
    adicionar_texto_centralizado(doc, "\n\n_______________________________________")
    adicionar_texto_centralizado(doc, "Enildo Manoel da Silva Junior")
    adicionar_texto_centralizado(doc, "Analista de Regulação")
