import os
import pandas as pd
from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
from num2words import num2words
import zipfile
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PyPDF2 import PdfWriter

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# ==========================
# FORMATAR REAL BR
# ==========================
def formatar_real(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def data_por_extenso():
    hoje = datetime.now()
    return f"Porto Alegre, {hoje.day} de {hoje.strftime('%B')} de {hoje.year}"

def competencia_mes(mes):
    ano = datetime.now().year
    return f"{mes}/{ano}"

# ==========================
# GERAR RECIBO
# ==========================
def gerar_recibo(vendedor, dados, mes, total):

    doc = Document()

    # Cabeçalho
    header = doc.sections[0].header
    p = header.paragraphs[0]
    p.text = data_por_extenso()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_heading("RECIBO DE PAGAMENTO", level=1)

    total_extenso = num2words(total, lang='pt_BR').capitalize()

    doc.add_paragraph(
        f"\nInformo que recebi da 3i Importação e Exportação Ltda, CNPJ 20.783.843/0001-19, a quantia de  {formatar_real(total)} "
        f"({total_extenso})), a título de serviços prestados."
    )

    doc.add_paragraph(f"\nCompetência: {competencia_mes(mes)}")

    # Separar proventos e descontos
    proventos = {k: v for k, v in dados.items() if v > 0}
    descontos = {k: v for k, v in dados.items() if v < 0}

    doc.add_heading("PROVENTOS", level=2)
    for desc, valor in proventos.items():
        doc.add_paragraph(f"{desc}: {formatar_real(valor)}")

    doc.add_heading("DESCONTOS", level=2)
    for desc, valor in descontos.items():
        doc.add_paragraph(f"{desc}: {formatar_real(valor)}")

    doc.add_heading(f"TOTAL LÍQUIDO: {formatar_real(total)}", level=2)

    nome_arquivo = os.path.join(UPLOAD_FOLDER, f"Recibo_{vendedor}.docx")
    doc.save(nome_arquivo)

    return nome_arquivo

# ==========================
# ROTA PRINCIPAL
# ==========================
@app.route("/", methods=["GET", "POST"])
def index():

    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril",
        "Maio", "Junho", "Julho", "Agosto",
        "Setembro", "Outubro", "Novembro", "Dezembro"
    ]

    mes_atual = meses[datetime.now().month - 1]

    if request.method == "POST":

        file = request.files["file"]
        caminho = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(caminho)

        df_full = pd.read_excel(caminho, header=None)

        mes_escolhido = request.form["mes"]
        mes_lower = mes_escolhido.lower()

        linha_mes = None
        col_inicio = None
        col_fim = None

        # ==========================
        # ACHAR MÊS
        # ==========================
        for i in range(len(df_full)):
            linha = df_full.iloc[i].fillna("").astype(str).str.lower()

            if any(mes_lower in cel for cel in linha):
                linha_mes = i

                for j, cel in enumerate(linha):
                    if mes_lower in cel:
                        col_inicio = j
                        break
                break

        if linha_mes is None:
            return f"Mês '{mes_escolhido}' não encontrado."

        # ==========================
        # ACHAR FIM DO BLOCO
        # ==========================
        meses_lista = [m.lower() for m in meses]

        for j in range(col_inicio + 1, len(df_full.columns)):
            cel = str(df_full.iloc[linha_mes, j] or "").lower()

            if any(m in cel for m in meses_lista) or "total" in cel:
                col_fim = j
                break

        if col_fim is None:
            col_fim = len(df_full.columns)

        # ==========================
        # LER BASE COMPLETA
        # ==========================
        header_row = linha_mes + 1
        df_base = pd.read_excel(caminho, header=header_row)

        # 🔥 PEGA DESCRIÇÃO (COLUNA B)
        descricao_coluna = df_base.columns[1]

        # 🔥 PEGA SOMENTE VALORES DO MÊS
        df_valores = df_base.iloc[:, col_inicio:col_fim]

        # 🔥 JUNTA DESCRIÇÃO + VALORES
        df = pd.concat([df_base[[descricao_coluna]], df_valores], axis=1)
        df = df.dropna(how="all")

        colunas = list(df.columns)
        vendedores = colunas[1:]

        arquivos_gerados = []

        for vendedor in vendedores:

            dados_vendedor = {}
            total_planilha = 0

            for i in range(len(df)):
                descricao = str(df.iloc[i][descricao_coluna]).strip()
                valor = df.iloc[i][vendedor]

                if descricao.upper() == "TOTAL":
                    try:
                        total_planilha = float(valor)
                    except:
                        total_planilha = 0
                    continue

                if pd.notna(valor):
                    try:
                        valor = float(valor)
                        if valor != 0:
                            dados_vendedor[descricao] = valor
                    except:
                        continue

            if dados_vendedor:
                arquivo = gerar_recibo(
                    vendedor,
                    dados_vendedor,
                    mes_escolhido,
                    total_planilha
                )
                arquivos_gerados.append(arquivo)

        if not arquivos_gerados:
            return "Nenhum recibo foi gerado."

        zip_path = os.path.join(UPLOAD_FOLDER, "Recibos_Gerados.zip")

        with zipfile.ZipFile(zip_path, "w") as zipf:
            for arquivo in arquivos_gerados:
                zipf.write(arquivo, os.path.basename(arquivo))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html", meses=meses, mes_atual=mes_atual)

if __name__ == "__main__":
    app.run(debug=True)