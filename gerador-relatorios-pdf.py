import os
import re
from collections import defaultdict
from fpdf import FPDF
import openpyxl
from datetime import datetime

# ============================================================
# CONFIGURAÇÕES - AJUSTE AQUI
# ============================================================
ARQUIVO_EXCEL = r"C:\Documentos\Projects\pyCharm\gera-relatorios\atendimentos-cbmdf.xlsx"
LOGO_PATH = r"C:\Documentos\Projects\pyCharm\gera-relatorios\logo-lavorato.png"
PASTA_SAIDA = r"C:\Documentos\Projects\pyCharm\gera-relatorios\relatorios"

# Nomes das colunas na planilha
COL_PACIENTE = "Paciente"
COL_PLANO = "Tipo Atendimento" # convênio (ex: CBMDF)
COL_TIPO_ATENDIMENTO = "Plano"  # procedimento/especialidade
COL_DATA = "Data"
COL_STATUS = "Status" # filtra só "Agendada" (ou todas)

# Meses por extenso
MESES = {
    1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril",
    5: "maio", 6: "junho", 7: "julho", 8: "agosto",
    9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro",
}

# Data do documento gerada automaticamente (data de hoje)
hoje = datetime.now()
DATA_DOCUMENTO = f"{hoje.day} de {MESES[hoje.month]} de {hoje.year}"

# Mapeamento: Tipo Atendimento -> nome que aparece no PDF
MAPA_ESPECIALIDADE = {
    "TERAPIA ABA - SESSAO": "PSICOLOGIA (TERAPIA ABA)",
    "TERAPIA ABA - ATENDIMENTO SEMANAL CONFORME ESPECIFICACAO MEDICA": "PSICOLOGIA (TERAPIA ABA)",
    "PSICOTERAPIA INDIVIDUAL": "PSICOLOGIA",
    "AVALIACAO PSICOLOGICA": "PSICOLOGIA",
    "PSICOPEDAGOGIA INDIVIDUAL": "PSICOPEDAGOGIA",
    "PSICOMOTRICIDADE INDIVIDUAL": "PSICOMOTRICIDADE",
    "SESSOES DE FONOTERAPIA/FONOAUDIOLOGIA": "FONOAUDIOLOGIA",
}

# Números por extenso
EXTENSO = {
    1: "uma", 2: "duas", 3: "três", 4: "quatro", 5: "cinco",
    6: "seis", 7: "sete", 8: "oito", 9: "nove", 10: "dez",
    11: "onze", 12: "doze", 13: "treze", 14: "quatorze", 15: "quinze",
    16: "dezesseis", 17: "dezessete", 18: "dezoito", 19: "dezenove", 20: "vinte",
    21: "vinte e uma", 22: "vinte e duas", 23: "vinte e três", 24: "vinte e quatro",
    25: "vinte e cinco", 26: "vinte e seis", 27: "vinte e sete", 28: "vinte e oito",
    29: "vinte e nove", 30: "trinta", 31: "trinta e um"
}


# ============================================================
# CLASSE DO PDF
# ============================================================
class RelatorioPDF(FPDF):
    def __init__(self, logo_path=None):
        super().__init__()
        self.logo_path = logo_path

    def header(self):
        if self.logo_path and os.path.exists(self.logo_path):
            logo_w = 50
            x_pos = (210 - logo_w) / 2
            self.image(self.logo_path, x=x_pos, y=10, w=logo_w)
            self.ln(35)
        else:
            self.ln(10)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(128, 128, 128)
        self.cell(0, 10, f"Página {self.page_no()} de {{nb}}", 0, 0, "R")


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================
def write_mixed(pdf, parts):
    """Escreve texto com trechos alternando negrito/normal."""
    for part in parts:
        style = "B" if part.get("bold") else ""
        pdf.set_font("Helvetica", style, 12)
        pdf.write(8, part["text"])


def sanitize(text):
    """Remove caracteres problemáticos para nomes de arquivo."""
    text = text.strip()
    text = re.sub(r'[\\/*?:"<>|]', "", text)
    text = text.replace(" ", "_")
    return text


def get_convenio_info(plano):
    """Retorna (saudê, nome completo, subtipo) baseado no plano."""
    p = plano.strip().upper()
    if p == "CBMDF":
        return (
            "SAÚDE CBMDF",
            "CORPO DE BOMBEIROS MILITAR DO DISTRITO FEDERAL (CBMDF)",
            "CBMDF TÍPICO",
        )
    elif p == "PMDF":
        return (
            "SAÚDE PMDF",
            "POLÍCIA MILITAR DO DISTRITO FEDERAL (PMDF)",
            "PMDF TÍPICO",
        )
    else:
        return (f"SAÚDE {p}", p, f"{p} TÍPICO")


def gerar_pdf(nome, plano, especialidade_label, sessoes, mes_ref, ano_ref):
    """Gera um PDF de solicitação para um paciente/especialidade."""

    pdf = RelatorioPDF(logo_path=LOGO_PATH)
    pdf.alias_nb_pages()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=25)

    # --- Data ---
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, f"Brasília, {DATA_DOCUMENTO}", 0, 1, "R")
    pdf.ln(10)

    # --- Destinatário ---
    saude, nome_completo, subtipo = get_convenio_info(plano)

    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 6, "Ao", 0, 1, "L")

    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 6, saude, 0, 1, "L")

    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 6, nome_completo, 0, 1, "L")
    pdf.cell(0, 6, subtipo, 0, 1, "L")

    pdf.set_font("Helvetica", "U", 12)
    pdf.cell(0, 6, "BRASÍLIA - DF", 0, 1, "L")
    pdf.ln(12)

    # --- Saudação ---
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 8, "Prezados(as) Senhores(as)", 0, 1, "L")
    pdf.ln(9)

    # --- Parágrafo 1 ---
    nome_upper = nome.strip().upper()
    write_mixed(pdf, [
        {"text": "Solicitamos autorização para realização de sessões de "},
        {"text": especialidade_label, "bold": True},
        {"text": " para a paciente "},
        {"text": nome_upper, "bold": True},
        {"text": ". O(a) paciente necessita de acompanhamento constante na especialidade mencionada para ter condições de desenvolvimento."},
    ])
    pdf.ln(10)

    # --- Parágrafo 2 ---
    sessoes_int = int(sessoes)
    extenso = EXTENSO.get(sessoes_int, str(sessoes_int))
    mes_nome = MESES.get(mes_ref, str(mes_ref))

    write_mixed(pdf, [
        {"text": "Esclarecemos que a intervenção na especialidade requerida exige acompanhamento constante para obtenção de bom resultado terapêutico. Por esta razão, solicitamos "},
        {"text": f"{sessoes_int} ({extenso})", "bold": True},
        {"text": " sessões para o mês de "},
        {"text": mes_nome, "bold": True},
        {"text": f" de {ano_ref}.", "bold": False},
    ])
    pdf.ln(20)

    # --- Despedida ---
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 8, "Atenciosamente,", 0, 1, "L")

    # --- Salvar ---
    os.makedirs(PASTA_SAIDA, exist_ok=True)
    nome_arquivo = f"{sanitize(nome)}_{sanitize(especialidade_label)}.pdf"
    caminho = os.path.join(PASTA_SAIDA, nome_arquivo)
    pdf.output(caminho)
    return caminho


# ============================================================
# LEITURA E AGRUPAMENTO
# ============================================================
def main():
    wb = openpyxl.load_workbook(ARQUIVO_EXCEL)
    ws = wb.active

    # Lê cabeçalhos
    headers = [cell.value.strip() if cell.value else "" for cell in ws[1]]
    idx = {h: i for i, h in enumerate(headers)}

    # Agrupa: (paciente, plano, especialidade_mapeada) -> contagem de sessões
    # Também guarda mês/ano de referência
    contagem = defaultdict(int)
    meses_encontrados = set()

    for row in ws.iter_rows(min_row=2, values_only=True):
        paciente = row[idx[COL_PACIENTE]]
        plano = row[idx[COL_PLANO]]
        tipo_atend = row[idx[COL_TIPO_ATENDIMENTO]]
        data = row[idx[COL_DATA]]

        if not paciente or not tipo_atend:
            continue

        # Mapeia o tipo de atendimento para o nome da especialidade no PDF
        tipo_upper = tipo_atend.strip().upper()
        especialidade_label = MAPA_ESPECIALIDADE.get(tipo_upper, tipo_upper)

        # Extrai mês/ano da data
        if hasattr(data, "month"):
            mes, ano = data.month, data.year
        else:
            # Tenta parsear string dd/mm/yyyy
            partes = str(data).split("/")
            mes, ano = int(partes[1]), int(partes[2])

        meses_encontrados.add((mes, ano))
        chave = (paciente.strip(), plano.strip(), especialidade_label)
        contagem[chave] += 1

    # Determina mês/ano de referência (pega o mais frequente)
    if meses_encontrados:
        mes_ref, ano_ref = max(meses_encontrados)
    else:
        mes_ref, ano_ref = 1, 2026

    # Gera PDFs
    print(f"Mês de referência detectado: {MESES[mes_ref]}/{ano_ref}")
    print(f"Total de combinações (paciente + especialidade): {len(contagem)}")
    print(f"{'='*60}\n")

    total = 0
    for (paciente, plano, especialidade), sessoes in sorted(contagem.items()):
        caminho = gerar_pdf(paciente, plano, especialidade, sessoes, mes_ref, ano_ref)
        print(f"✓ {paciente:<45} | {especialidade:<30} | {sessoes} sessões")
        total += 1

    print(f"\n{'='*60}")
    print(f"Total de relatórios gerados: {total}")
    print(f"Pasta de saída: {os.path.abspath(PASTA_SAIDA)}")


if __name__ == "__main__":
    main()