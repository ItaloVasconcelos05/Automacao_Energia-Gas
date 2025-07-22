from docx import Document
from docx2pdf import convert
from docx.shared import Inches
import pandas as pd
from docx2pdf import convert
from tqdm import tqdm
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import sys
import os
from sections.introduction.introduction import gerar_secao_introducao
from sections.legalbasis.legalbasis import gerar_secao_fundamentacao_legal
from sections.nonconformity.nonconformity import (
    gerar_secao_nao_conformidades_constatadas,
)
from sections.nonconformityresume.nonconformityresume import (
    gerar_secao_resumo_nao_conformidades,
)
from sections.finalconsiderations.finalconsiderations import (
    gerar_secao_consideracoes_finais,
)
from utils import (
    adicionar_texto_centralizado,
    ajustar_largura_colunas,
    arquivo_em_uso,
)


def gerar_relatorio():
    """
    Gera o relatório completo (docx + pdf) com base nos dados da fiscalização.

    Parâmetros:
    - row: linha da planilha (Series).
    - nao_conformidades_df: DataFrame da aba 'Não-conformidades'.
    - fotos_dir: pasta com as imagens.
    - pasta_saida: onde salvar o .docx e .pdf.
    """

    if getattr(sys, "frozen", False):
        BASE_DIR = os.path.dirname(sys.executable)
    else:
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    FOTOS_DIR = os.path.join(BASE_DIR, "assets")
    RELATORIOS_DIR = os.path.join(BASE_DIR, "reports")
    CAMINHO_PLANILHA = os.path.join(BASE_DIR, "planilha_fiscalizacao.xlsx")
    COLUNA_STATUS = "Relatório Gerado"

    os.makedirs(RELATORIOS_DIR, exist_ok=True)
    os.makedirs(FOTOS_DIR, exist_ok=True)

    if arquivo_em_uso(CAMINHO_PLANILHA):
        print("⚠️ A planilha está em uso. Feche-a antes de executar o script.")
        exit(1)

    fiscalizacoes_df = pd.read_excel(CAMINHO_PLANILHA, sheet_name="Fiscalizações")
    nao_conformidades_df = pd.read_excel(
        CAMINHO_PLANILHA, sheet_name="Não-conformidades "
    )

    observacoes_df = pd.read_excel(
        CAMINHO_PLANILHA, sheet_name="Observações Importantes"
    )

    if COLUNA_STATUS not in fiscalizacoes_df.columns:
        fiscalizacoes_df[COLUNA_STATUS] = False
    fiscalizacoes_df[COLUNA_STATUS] = (
        fiscalizacoes_df[COLUNA_STATUS].fillna(False).astype(bool)
    )

    pendentes = fiscalizacoes_df[~fiscalizacoes_df[COLUNA_STATUS]]

    if pendentes.empty:
        print("✅ Nenhum relatório pendente.")
        return

    for idx in tqdm(pendentes.index, desc="Gerando relatórios"):
        row = fiscalizacoes_df.loc[idx]
        id_fisc = row["ID da Fiscalização"]
        doc = Document()

        doc.add_picture(os.path.join(BASE_DIR, "assets/logo_arpe.png"), width=Inches(2))
        logo_arpe = doc.paragraphs[-1]
        logo_arpe.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        adicionar_texto_centralizado(doc, "DIRETORIA DE REGULAÇÃO TÉCNICO-OPERACIONAL")
        adicionar_texto_centralizado(doc, "COORDENADORIA DE TRANSPORTES E RODOVIAS")
        adicionar_texto_centralizado(
            doc, "RELATÓRIO DE FISCALIZAÇÃO TÉCNICO-OPERACIONAL CTR 01/2025"
        )
        adicionar_texto_centralizado(
            doc, "TERMINAIS RODOVIÁRIOS INTERMUNICIPAIS CONCEDIDOS À EMPRESA SOCICAM"
        )
        adicionar_texto_centralizado(
            doc, "CONTRATO DE CONCESSÃO DE SERVIÇO PÚBLICO Nº 1.041.080/08"
        )

        doc.add_section(WD_SECTION.NEW_PAGE)

        gerar_secao_introducao(doc, row)
        gerar_secao_fundamentacao_legal(doc)
        gerar_secao_nao_conformidades_constatadas(
            doc, row, nao_conformidades_df, FOTOS_DIR, observacoes_df
        )
        gerar_secao_resumo_nao_conformidades(doc, row, nao_conformidades_df)
        gerar_secao_consideracoes_finais(doc, row)

        nome_arquivo = f"relatorio_{id_fisc}"
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        convert(caminho_docx, caminho_pdf)
        fiscalizacoes_df.at[idx, COLUNA_STATUS] = True

    if not arquivo_em_uso(CAMINHO_PLANILHA):
        with pd.ExcelWriter(
            CAMINHO_PLANILHA, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
            # Substitui somente as abas que foram manipuladas
            fiscalizacoes_df.to_excel(writer, sheet_name="Fiscalizações", index=False)
            nao_conformidades_df.to_excel(
                writer, sheet_name="Não-conformidades ", index=False
            )
        ajustar_largura_colunas(CAMINHO_PLANILHA)

    print("🎉 Relatórios gerados e planilha atualizada com sucesso.")

    return caminho_docx, caminho_pdf
