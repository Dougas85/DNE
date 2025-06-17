import os
from flask import Flask, render_template, request, flash, redirect, url_for
from werkzeug.utils import secure_filename
from robo import processar_arquivo

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.secret_key = 'segredo'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    print("Templates disponíveis:", os.listdir('templates'))
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
                resultado = processar_arquivo(caminho)
                flash(resultado)
            except Exception as e:
                flash(f'Ocorreu um erro durante o processamento: {e}')

            return redirect(url_for('index'))
        else:
            flash('Formato inválido. Envie um arquivo .xlsx')
            return redirect(request.url)

    return render_template('upload.html')

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
