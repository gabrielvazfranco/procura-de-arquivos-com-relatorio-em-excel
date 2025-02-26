import os
import pandas as pd
from datetime import datetime
from flask import Flask, render_template, request, send_file

# Definir o diretório base do projeto
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
STATIC_DIR = os.path.join(BASE_DIR, "static")

# Criar o app Flask
app = Flask(__name__, template_folder=TEMPLATES_DIR, static_folder=STATIC_DIR)

# Criar a pasta de relatórios se não existir
os.makedirs("relatorios", exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def index():
    print("Arquivos na pasta templates:", os.listdir(TEMPLATES_DIR))  # Debugging
    if request.method == "POST":
        diretorio = request.form.get("diretorio")
        ano_limite = request.form.get("ano_limite")

        if not diretorio or not ano_limite:
            return render_template("index.html", erro="Por favor, selecione um diretório e informe o ano.")

        try:
            ano_limite = int(ano_limite)
            data_limite = datetime(ano_limite + 1, 1, 1)
        except ValueError:
            return render_template("index.html", erro="Ano inválido.")

        dados = []
        for root, _, files in os.walk(diretorio):
            for file in files:
                try:
                    caminho_arquivo = os.path.join(root, file)
                    ultima_modificacao = datetime.utcfromtimestamp(os.path.getmtime(caminho_arquivo))
                    ultimo_acesso = datetime.utcfromtimestamp(os.path.getatime(caminho_arquivo))
                    tamanho_arquivo = os.path.getsize(caminho_arquivo) / (1024 * 1024)  # MB

                    if ultima_modificacao < data_limite or ultimo_acesso < data_limite:
                        dados.append({
                            "Nome": file,
                            "Pasta": root,
                            "Última Modificação": ultima_modificacao,
                            "Último Acesso": ultimo_acesso,
                            "Tamanho (MB)": round(tamanho_arquivo, 2)
                        })
                except Exception as e:
                    print(f"Erro ao acessar {caminho_arquivo}: {e}")

        if dados:
            df = pd.DataFrame(dados)
            nome_arquivo = f"relatorios/relatorio_{ano_limite}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df.to_excel(nome_arquivo, index=False)
            return render_template("index.html", download_link=nome_arquivo)

        return render_template("index.html", erro="Nenhum arquivo encontrado.")

    return render_template("index.html")

@app.route("/download/<path:filename>")
def download(filename):
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)