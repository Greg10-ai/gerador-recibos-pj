import os
import pandas as pd
import matplotlib
matplotlib.use('Agg')

import matplotlib.pyplot as plt
from flask import Flask, render_template, request, send_file, redirect, url_for, session
from docx import Document
from docx.shared import Inches
from datetime import datetime
from num2words import num2words
import zipfile
from docx.enum.text import WD_ALIGN_PARAGRAPH
import unicodedata

app = Flask(__name__)
app.secret_key = "segredo123"

UPLOAD_FOLDER = "/tmp/uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

USUARIO = "Evelyn"
SENHA = "Monique"

# ==========================
# NORMALIZAR NOMES
# ==========================
def normalizar_nome(nome):
    nome = str(nome).strip().lower()
    nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
    nome = " ".join(nome.split())
    nome = nome.replace(" ", "")
    return nome

# ==========================
# DADOS BANCÁRIOS (CORRIGIDO)
# ==========================
def carregar_dados_bancarios(caminho):
    df = pd.read_excel(caminho)

    # normaliza colunas
    df.columns = [normalizar_nome(col) for col in df.columns]

    mapa = {}

    for _, row in df.iterrows():
        nome = normalizar_nome(row.get("vendedores", ""))

        info = {
            "nome": row.get("vendedores", ""),
            "cnpj": row.get("cnpj", ""),
            "banco": row.get("banco", ""),
            "agencia": row.get("agencia", ""),
            "conta": row.get("conta", ""),
            "pix": row.get("pix", "")
        }

        mapa[nome] = info

    return mapa

# MATCH INTELIGENTE
def encontrar_dados_bancarios(nome_vendedor, mapa_banco):
    chave = normalizar_nome(nome_vendedor)

    if chave in mapa_banco:
        return mapa_banco[chave]

    for k in mapa_banco:
        if chave in k or k in chave:
            print(f"⚠️ Banco match aproximado: {nome_vendedor} -> {k}")
            return mapa_banco[k]

    print(f"❌ Sem dados bancários para: {nome_vendedor}")
    return {}

def encontrar_imagem(nome_vendedor, mapa_imagens):
    chave = normalizar_nome(nome_vendedor)

    if chave in mapa_imagens:
        return mapa_imagens[chave]

    for k in mapa_imagens:
        if chave in k or k in chave:
            return mapa_imagens[k]

    return None

# ==========================
def tratar_valor(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, str):
        valor = valor.replace("%", "").replace(".", "").replace(",", ".").strip()

    try:
        return round(float(valor), 2)
    except:
        return None

def formatar_real(valor):
    valor = round(float(valor), 2)
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_data_extenso(data_str):
    meses = ["janeiro","fevereiro","março","abril","maio","junho",
             "julho","agosto","setembro","outubro","novembro","dezembro"]

    data = datetime.strptime(data_str, "%Y-%m-%d")
    return f"Porto Alegre, {data.day} de {meses[data.month - 1]} de {data.year}"

def competencia_mes(mes):
    return f"{mes}/{datetime.now().year}"

# ==========================
def gerar_imagens_abas(caminho_excel, mes_escolhido):
    imagens = {}
    xls = pd.ExcelFile(caminho_excel)

    for aba in xls.sheet_names:
        try:
            df_full = pd.read_excel(xls, sheet_name=aba, header=None)

            mes_lower = mes_escolhido.lower()
            linha_mes = None
            col_mes = None

            for i in range(len(df_full)):
                linha = df_full.iloc[i].fillna("").astype(str).str.lower()

                if any(mes_lower in cel for cel in linha):
                    linha_mes = i
                    for j, cel in enumerate(linha):
                        if mes_lower in cel:
                            col_mes = j
                            break
                    break

            if linha_mes is None or col_mes is None:
                continue

            header_row = linha_mes + 1
            df = pd.read_excel(xls, sheet_name=aba, header=header_row)

            descricao_col = df.columns[1]
            valor_col = df.columns[col_mes]

            df_ab = df[[descricao_col, valor_col]].dropna(how="all").head(50)

            fig, ax = plt.subplots(figsize=(8, len(df_ab) * 0.4 + 1))
            ax.axis('off')

            tabela = ax.table(
                cellText=df_ab.values,
                colLabels=df_ab.columns,
                loc='center'
            )

            tabela.auto_set_font_size(False)
            tabela.set_fontsize(8)

            caminho_img = os.path.join(UPLOAD_FOLDER, f"{aba}.png")

            plt.savefig(caminho_img, bbox_inches='tight')
            plt.close(fig)

            imagens[normalizar_nome(aba)] = caminho_img

        except:
            continue

    return imagens

# ==========================
def gerar_recibo(vendedor, dados, mes, total, data_recibo, imagem=None, dados_bancarios=None):

    doc = Document()

    header = doc.sections[0].header
    p = header.paragraphs[0]
    p.text = formatar_data_extenso(data_recibo)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_heading("RECIBO DE PAGAMENTO", level=1)

    total_extenso = num2words(total, lang='pt_BR').capitalize()

    doc.add_paragraph(
        f"\nInformo que recebi da empresa 3i Importação e Exportação Ltda, CNPJ 20.783.843/0001-19, "
        f"o valor de {formatar_real(total)} "
        f"({total_extenso}),"
        f"referente a serviços prestados de gestão comercial."
    )

    doc.add_paragraph(f"\nCompetência: {competencia_mes(mes)}")

    doc.add_heading("PROVENTOS", level=2)
    for desc, valor in dados.items():
        if valor > 0:
            doc.add_paragraph(f"{desc}: {formatar_real(valor)}")

    doc.add_heading("DESCONTOS", level=2)
    for desc, valor in dados.items():
        if valor < 0:
            doc.add_paragraph(f"{desc}: {formatar_real(valor)}")

    doc.add_heading(f"TOTAL LÍQUIDO: {formatar_real(total)}", level=2)

    # 🔥 DADOS BANCÁRIOS NO FINAL
    if dados_bancarios:
        doc.add_paragraph("")
        doc.add_heading("DADOS BANCÁRIOS", level=2)
        doc.add_paragraph(f"Nome: {dados_bancarios.get('nome','')}")
        doc.add_paragraph(f"CNPJ: {dados_bancarios.get('cnpj','')}")
        doc.add_paragraph(f"Banco: {dados_bancarios.get('banco','')}")
        doc.add_paragraph(f"Agência: {dados_bancarios.get('agencia','')}")
        doc.add_paragraph(f"Conta: {dados_bancarios.get('conta','')}")
        doc.add_paragraph(f"PIX: {dados_bancarios.get('pix','')}")

    if imagem and os.path.exists(imagem):
        doc.add_page_break()

        titulo = doc.add_paragraph(f"Apuração Mês {mes}/{datetime.now().year}")
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p_img.add_run()
        run.add_picture(imagem, width=Inches(4))

    caminho = os.path.join(UPLOAD_FOLDER, f"Recibo_{vendedor}.docx")
    doc.save(caminho)

    return caminho

# ==========================
@app.route("/")
def home():
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def login():
    if request.form["usuario"] == USUARIO and request.form["senha"] == SENHA:
        session["logado"] = True
        return redirect(url_for("sistema"))
    return "Login inválido"

@app.route("/sistema", methods=["GET", "POST"])
def sistema():

    if not session.get("logado"):
        return redirect(url_for("home"))

    meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
             "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

    mes_atual = meses[datetime.now().month - 1]

    if request.method == "POST":

        data_recibo = request.form.get("data_recibo")

        file = request.files["file"]
        caminho = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(caminho)

        mapa_banco = {}
        file_banco = request.files.get("file_banco")

        if file_banco and file_banco.filename != "":
            caminho_banco = os.path.join(UPLOAD_FOLDER, file_banco.filename)
            file_banco.save(caminho_banco)
            mapa_banco = carregar_dados_bancarios(caminho_banco)

        mapa_imagens = {}
        file_imagens = request.files.get("file_imagens")

        if file_imagens and file_imagens.filename != "":
            caminho_img_excel = os.path.join(UPLOAD_FOLDER, file_imagens.filename)
            file_imagens.save(caminho_img_excel)

            mapa_imagens = gerar_imagens_abas(
                caminho_img_excel,
                request.form["mes"]
            )

        df_full = pd.read_excel(caminho, header=None)

        mes_escolhido = request.form["mes"].lower()

        linha_mes = None
        col_inicio = None
        col_fim = None

        for i in range(len(df_full)):
            linha = df_full.iloc[i].fillna("").astype(str).str.lower()

            if any(mes_escolhido in cel for cel in linha):
                linha_mes = i
                for j, cel in enumerate(linha):
                    if mes_escolhido in cel:
                        col_inicio = j
                        break
                break

        if linha_mes is None:
            return "Mês não encontrado."

        meses_lista = [m.lower() for m in meses]

        for j in range(col_inicio + 1, len(df_full.columns)):
            cel = str(df_full.iloc[linha_mes, j] or "").lower()
            if any(m in cel for m in meses_lista) or "total" in cel:
                col_fim = j
                break

        if col_fim is None:
            col_fim = len(df_full.columns)

        header_row = linha_mes + 1
        df_base = pd.read_excel(caminho, header=header_row)

        descricao_coluna = df_base.columns[1]
        df_valores = df_base.iloc[:, col_inicio:col_fim]

        df = pd.concat([df_base[[descricao_coluna]], df_valores], axis=1)
        df = df.dropna(how="all")

        vendedores = df.columns[1:]
        arquivos = []

        for vendedor in vendedores:

            dados = {}
            total_planilha = 0

            for i in range(len(df)):
                descricao = str(df.iloc[i][descricao_coluna]).strip()
                valor_bruto = df.iloc[i][vendedor]

                valor = tratar_valor(valor_bruto)

                if descricao.upper() == "TOTAL":
                    if valor is not None:
                        total_planilha = valor
                    continue

                if valor is not None and valor != 0:
                    dados[descricao] = valor

            if dados:
                imagem_vendedor = encontrar_imagem(vendedor, mapa_imagens)

                dados_bancarios = encontrar_dados_bancarios(vendedor, mapa_banco)

                arquivo = gerar_recibo(
                    vendedor,
                    dados,
                    request.form["mes"],
                    total_planilha,
                    data_recibo,
                    imagem_vendedor,
                    dados_bancarios
                )

                arquivos.append(arquivo)

        zip_path = os.path.join(UPLOAD_FOLDER, "Recibos.zip")

        with zipfile.ZipFile(zip_path, "w") as zipf:
            for arq in arquivos:
                zipf.write(arq, os.path.basename(arq))

        return send_file(zip_path, as_attachment=True)

    return render_template("index.html", meses=meses, mes_atual=mes_atual)

if __name__ == "__main__":
    app.run(debug=True)