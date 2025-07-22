from utils import (
    adicionar_titulo_secao,
    adicionar_paragrafo_justificado,
    adicionar_texto_centralizado,
)


def gerar_secao_consideracoes_finais(doc, row):
    """
    Gera a seção '5. CONSIDERAÇÕES FINAIS' do relatório.
    Parâmetros:
    - doc: objeto Document.
    - row: linha da fiscalização (Series do DataFrame fiscalizacoes_df).
    """

    adicionar_titulo_secao(doc, "5. CONSIDERAÇÕES FINAIS")

    texto1 = "Diante das constatações apontadas no presente Relatório de Fiscalização, solicita-se o seu envio à SOCICAM para que esta apresente as providências para sanar as Não Conformidades evidenciadas, bem como para estabelecer os respectivos prazos de conclusão dos serviços e obras que forem necessários."

    texto2 = "Conforme sinalizado nas Observações Importantes, a Arpe vai continuar monitorando a montagem do elevador para PCD e a adequação do Sistema contra Incêndio e do Sistema de Proteção contra Descargas Atmosféricas (SPDA) no Terminal de Recife. Deve-se levar em consideração que a Concessionária, em carta enviada à Arpe, deu prazo até abril de 2025 para realizar a adequação desses Sistemas nos terminais de Caruaru, Garanhuns e Petrolina, mas não apresentou prazo para os terminais de Arcoverde e Serra Talhada. Dessa forma, recomenda-se que a SOCICAM mantenha a Arpe atualizada com documentos e informações referentes à adequação tanto do Sistema contra Incêndio quanto do Sistema de Proteção contra Descargas Atmosféricas (SPDA) de todos os terminais concedidos."

    texto3 = "É recomendável ainda que a SOCICAM observe a inadequação do armazenamento de objetos nas áreas destinadas aos extintores e que mantenha os equipamentos de auxílio a pessoas com necessidades específicas como cadeiras de rodas e de transbordo de maneira a facilitar o seu uso em caso de necessidade."

    texto4 = "Por fim, indicamos o encaminhamento deste Relatório de Fiscalização para conhecimento da EPTI, na qualidade de Poder Concedente do Contrato de Concessão e gestora do Sistema de Transporte Coletivo Intermunicipal de Passageiros (STCIP-PE)."

    adicionar_paragrafo_justificado(doc, texto1)
    adicionar_paragrafo_justificado(doc, texto2)
    adicionar_paragrafo_justificado(doc, texto3)
    adicionar_paragrafo_justificado(doc, texto4)

    adicionar_texto_centralizado(doc, f"\n\nRecife, {row['Data']}.")
    adicionar_texto_centralizado(doc, "\n\n")

    # Assinaturas dos responsáveis (pode ser uma string separada por ";" ou ",")
    assinantes = str(row.get("Assinatura", "")).split(";")
    for assinante in assinantes:
        nome = assinante.strip()
        if nome:
            adicionar_texto_centralizado(doc, "_______________________")
            adicionar_texto_centralizado(doc, nome)
            adicionar_texto_centralizado(doc, "Analista de Regulação")
            adicionar_texto_centralizado(doc, "")  # Espaço em branco entre assinaturas

    adicionar_texto_centralizado(doc, "\nCiente e de acordo:\n")
    coordenador = str(row.get("Coordenador", "")).strip()
    if coordenador:
        adicionar_texto_centralizado(doc, "_______________________")
        adicionar_texto_centralizado(doc, coordenador)
