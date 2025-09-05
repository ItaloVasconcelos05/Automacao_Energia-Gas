from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import *


def gerar_secao_objetivos(doc: Document):
    """
    Gera a seção '2. OBJETIVOS' do relatório.
    Parâmetros:
    - doc: objeto Document do python-docx.
    """
    adicionar_paragrafo(doc, "2. OBJETIVOS", tamanho_fonte=14, alinhamento=WD_ALIGN_PARAGRAPH.LEFT, bold=True, estilo='Heading 3', espaco_antes=25, espaco_depois=6, cor=(0,0,0))
    
    texto_objetivos = "A fiscalização direta e periódica tem por objetivo verificar o grau de conformidade das unidades operacionais dos com as legislações e normas vigentes dos serviços de distribuição de gás natural e determinar e/ou recomendar medidas corretivas. Os objetivos específicos são:"
    adicionar_paragrafo(doc, texto_objetivos, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_depois=12)

    # Objetivos em formato de lista com marcadores
    objetivos = [
        ("Conformidade Legal: ", "Verificar e assegurar o cumprimento das normas legais e regulamentares aplicáveis ao setor de gás canalizado, especificamente para os equipamentos encontrados nos sistemas de distribuição, a exemplo de medidores, lacres, tubos, canos e placas de identificação;"),
        ("Condições Operacionais, de Conservação e Manutenção: ", "Analisar as condições técnico-operacionais com foco na eficiência do sistema, atendando-se para o estado de conservação das unidades, de suas condições de manutenção e de segurança.")
    ]
    for titulo, texto in objetivos:
        adicionar_paragrafo(doc, f"• {titulo}{texto}", tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, bold=[titulo], espaco_depois=6)

    # Texto final com espaçamento
    texto_final = "De acordo com a Agenda Regulatória da Coordenadoria de Energia Elétrica e Gás Canalizado, a meta de fiscalizações a serem realizadas no ano de 2025 é de 372. Considerando essa meta, a equipe tenta abranger todos os municípios em que há instalações da Copergás e os diferentes nichos existentes."
    adicionar_paragrafo(doc, texto_final, tamanho_fonte=12, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY, espaco_antes=12, espaco_depois=12)
    
    
