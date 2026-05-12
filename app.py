import os
import pandas as pd
import unicodedata
import zipfile
from flask import Flask, render_template, request, send_file, redirect, url_for, session
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
from num2words import num2words
from playwright.sync_api import sync_playwright

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
# DADOS BANCÁRIOS
# ==========================
def carregar_dados_bancarios(caminho):
    df = pd.read_excel(caminho)
    df.columns = [normalizar_nome(col) for col in df.columns]

    mapa = {}
    for _, row in df.iterrows():
        nome = normalizar_nome(row.get("vendedores", ""))
        info = {
            "nome":    row.get("vendedores", ""),
            "cnpj":    row.get("cnpj", ""),
            "banco":   row.get("banco", ""),
            "agencia": row.get("agencia", ""),
            "conta":   row.get("conta", ""),
            "pix":     row.get("pix", "")
        }
        mapa[nome] = info

    return mapa

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
# UTILITÁRIOS
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
# ROTA: APURAÇÃO (usada pelo Playwright para screenshot)
# ==========================
@app.route("/apuracao")
def apuracao():
    linhas      = request.args.get("linhas", "[]")
    mes         = request.args.get("mes", "")
    ano         = request.args.get("ano", "")
    col_desc    = request.args.get("col_descricao", "Descrição")
    col_valor   = request.args.get("col_valor", "Valor")

    import json
    linhas = json.loads(linhas)

    return render_template(
        "apuracao.html",
        linhas=linhas,
        mes=mes,
        ano=ano,
        col_descricao=col_desc,
        col_valor=col_valor
    )

# ==========================
# GERAR IMAGEM VIA PLAYWRIGHT
# ==========================
def gerar_imagens_abas(caminho_excel, mes_escolhido, base_url="http://localhost:5000"):
    import json
    imagens = {}
    xls = pd.ExcelFile(caminho_excel)

    CAMPOS_PERCENTUAL = [
        "icm meta", "icm novos", "meta margem", "real",
        "icm", "icm meta base ativa", "% carteira ativa", "% liquidado"
    ]

    def formatar_celula(v, descricao=""):
        desc_lower = str(descricao).strip().lower()
        eh_percentual = any(c.lower() == desc_lower for c in CAMPOS_PERCENTUAL)

        try:
            f = float(v)
            if eh_percentual:
                if abs(f) <= 1.5:
                    return f"{f * 100:.2f}%"
                return f"{f:.2f}%"
        except:
            return str(v) if pd.notna(v) else "-"

        try:
            f = float(v)
            if f == int(f):
                return f"{int(f):,}".replace(",", ".")
            return f"R$ {f:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return str(v) if pd.notna(v) else "-"

    for aba in xls.sheet_names:
        try:
            df_full = pd.read_excel(xls, sheet_name=aba, header=None)

            mes_lower = mes_escolhido.lower()
            linha_mes = None
            col_mes   = None

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
                print(f"⚠️ Mês '{mes_escolhido}' não encontrado na aba '{aba}'")
                continue

            header_row = linha_mes + 1
            df = pd.read_excel(xls, sheet_name=aba, header=header_row)

            descricao_col = df.columns[1]
            valor_col     = df.columns[col_mes]

            df_ab = df[[descricao_col, valor_col]].dropna(how="all").head(60)

            df_formatado = df_ab.copy()
            df_formatado[valor_col] = [
                formatar_celula(row[valor_col], row[descricao_col])
                for _, row in df_ab.iterrows()
            ]

            # Monta lista de linhas como dicts para o template
            linhas = []
            for _, row in df_formatado.iterrows():
                desc  = row[descricao_col]
                valor = row[valor_col]
                linhas.append({
                    "descricao": str(desc),
                    "valor":     str(valor),
                    "is_total":  "total" in str(desc).lower()
                })

            # Monta URL para a rota /apuracao
            params = {
                "mes":          mes_escolhido.capitalize(),
                "ano":          str(datetime.now().year),
                "col_descricao": str(descricao_col),
                "col_valor":    str(valor_col),
                "linhas":       json.dumps(linhas, ensure_ascii=False)
            }

            from urllib.parse import urlencode
            url = f"{base_url}/apuracao?{urlencode(params)}"

            caminho_img = os.path.join(UPLOAD_FOLDER, f"{aba}.png")

            # Screenshot via Playwright apontando para a rota Flask
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page    = browser.new_page()
                page.goto(url)
                page.wait_for_timeout(500)
                page.screenshot(path=caminho_img, full_page=True)
                browser.close()

            imagens[normalizar_nome(aba)] = caminho_img
            print(f"✅ Imagem gerada: {aba} -> {caminho_img}")

        except Exception as e:
            print(f"❌ Erro na aba '{aba}': {e}")
            continue

    return imagens

# ==========================
# GERAR RECIBO
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
        f"({total_extenso}), "
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

    if dados_bancarios:
        doc.add_paragraph("")
        doc.add_heading("DADOS BANCÁRIOS", level=2)
        doc.add_paragraph(f"Nome: {dados_bancarios.get('nome', '')}")
        doc.add_paragraph(f"CNPJ: {dados_bancarios.get('cnpj', '')}")
        doc.add_paragraph(f"Banco: {dados_bancarios.get('banco', '')}")
        doc.add_paragraph(f"Agência: {dados_bancarios.get('agencia', '')}")
        doc.add_paragraph(f"Conta: {dados_bancarios.get('conta', '')}")
        doc.add_paragraph(f"PIX: {dados_bancarios.get('pix', '')}")

    if imagem and os.path.exists(imagem):
        doc.add_page_break()

        titulo = doc.add_paragraph(f"Apuração Mês {mes}/{datetime.now().year}")
        titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

        run = p_img.add_run()
        run.add_picture(imagem, width=Inches(5))

    caminho = os.path.join(UPLOAD_FOLDER, f"Recibo_{vendedor}.docx")
    doc.save(caminho)

    return caminho

# ==========================
# ROTAS
# ==========================
@app.route("/")
def home():
    return render_template("login.html")

@app.route("/login", methods=["POST"])
def login():
    if request.form["usuario"] == USUARIO and request.form["senha"] == SENHA:
        session["logado"] = True
        return redirect(url_for("sistema"))
    return render_template("login.html", erro="Usuário ou senha inválidos."), 401

@app.route("/sistema", methods=["GET", "POST"])
def sistema():

    if not session.get("logado"):
        return redirect(url_for("home"))

    meses = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho",
             "Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"]

    mes_atual = meses[datetime.now().month - 1]

    if request.method == "POST":

        try:
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

                # Pega a URL base do Railway automaticamente
                base_url = request.host_url.rstrip("/")

                mapa_imagens = gerar_imagens_abas(
                    caminho_img_excel,
                    request.form["mes"],
                    base_url=base_url
                )

            df_full = pd.read_excel(caminho, header=None)
            mes_escolhido = request.form["mes"].lower()

            linha_mes  = None
            col_inicio = None
            col_fim    = None

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
                return render_template("index.html", meses=meses, mes_atual=mes_atual,
                                       erro="❌ Mês não encontrado no arquivo.")

            meses_lista = [m.lower() for m in meses]

            for j in range(col_inicio + 1, len(df_full.columns)):
                cel = str(df_full.iloc[linha_mes, j] or "").lower()
                if any(m in cel for m in meses_lista) or "total" in cel:
                    col_fim = j
                    break

            if col_fim is None:
                col_fim = len(df_full.columns)

            header_row    = linha_mes + 1
            df_base       = pd.read_excel(caminho, header=header_row)
            descricao_col = df_base.columns[1]
            df_valores    = df_base.iloc[:, col_inicio:col_fim]
            df            = pd.concat([df_base[[descricao_col]], df_valores], axis=1)
            df            = df.dropna(how="all")

            vendedores = df.columns[1:]
            arquivos   = []

            for vendedor in vendedores:
                dados          = {}
                total_planilha = 0

                for i in range(len(df)):
                    descricao   = str(df.iloc[i][descricao_col]).strip()
                    valor_bruto = df.iloc[i][vendedor]
                    valor       = tratar_valor(valor_bruto)

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

        except Exception as e:
            import traceback
            erro = traceback.format_exc()
            print("❌ ERRO:", erro)
            return render_template("index.html", meses=meses, mes_atual=mes_atual,
                                   erro=f"Erro interno: {e}"), 500

    return render_template("index.html", meses=meses, mes_atual=mes_atual)


if __name__ == "__main__":
    app.run(debug=True)