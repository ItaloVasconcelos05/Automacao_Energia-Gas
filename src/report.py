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
from sections import (
    gerar_secao_consideracoes_finais,
    gerar_secao_introducao,
    gerar_secao_fundamentacao_legal,
    gerar_secao_nao_conformidades_constatadas,
    gerar_secao_resumo_nao_conformidades,
    gerar_secao_objetivos,
    gerar_titulo,
)
from utils import (adicionar_paragrafo,ajustar_largura_colunas,arquivo_em_uso,)

Inches
def gerar_relatorio():

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
        # fiscalizacoes_df.at[idx, COLUNA_STATUS] = True  # Comentado para manter o status como False após gerar o relatório
        doc = Document()
        gerar_titulo(doc, BASE_DIR)
        doc.add_section(WD_SECTION.NEW_PAGE)
        gerar_secao_introducao(doc, row, BASE_DIR)
        gerar_secao_objetivos(doc)
        gerar_secao_fundamentacao_legal(doc)
        gerar_secao_nao_conformidades_constatadas(doc, row, nao_conformidades_df, FOTOS_DIR, observacoes_df)
        
        gerar_secao_resumo_nao_conformidades(doc, row, nao_conformidades_df)
        
        gerar_secao_consideracoes_finais(doc, row)

        nome_arquivo = f"relatorio_{id_fisc}"
        caminho_docx = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.docx")
        caminho_pdf = os.path.join(RELATORIOS_DIR, f"{nome_arquivo}.pdf")

        doc.save(caminho_docx)
        # Garante sobrescrita limpa do PDF e captura erros de conversão
        try:
            if os.path.exists(caminho_pdf): 
                os.remove(caminho_pdf)
            convert(caminho_docx, caminho_pdf)
        except Exception as e:
            print(f"Erro ao converter para PDF: {e}")
            print("Verifique se o Microsoft Word está instalado e fechado durante a conversão.")
        # fiscalizacoes_df.at[idx, COLUNA_STATUS] = True

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
