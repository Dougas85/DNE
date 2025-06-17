import os
import requests
from flask import Flask, render_template, request, flash, redirect, url_for
from werkzeug.utils import secure_filename

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.secret_key = 'segredo'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Endereço da máquina local onde o robô está rodando
ROBO_LOCAL_URL = "http://SEU_IP_LOCAL:5001/processar"  # Substitua pelo IP da sua máquina local

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('excel')
        if not file:
            flash('Nenhum arquivo enviado.')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            caminho = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(caminho)

            flash(f'Arquivo {filename} enviado com sucesso!')

            try:
                # Envia o nome do arquivo para o robô local processar
                resposta = requests.post(ROBO_LOCAL_URL, json={"arquivo": filename})
                if resposta.status_code == 200:
                    flash("Arquivo enviado para processamento local.")
                else:
                    flash(f"Falha ao comunicar com o robô local: {resposta.text}")
            except Exception as e:
                flash(f"Erro ao tentar se comunicar com o robô local: {e}")

            return redirect(url_for('index'))
        else:
            flash('Formato inválido. Envie um arquivo .xlsx')
            return redirect(request.url)

    return render_template('upload.html')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
