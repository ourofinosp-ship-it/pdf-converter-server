from flask import Flask, request, send_file, after_this_request
import subprocess
import os
import tempfile
import time

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # Limite de 500MB

# Caminho para o LibreOffice no Docker
LIBREOFFICE_PATH = "libreoffice"

@app.route('/')
def index():
    return '''
    <h2>Conversor de Documentos para PDF</h2>
    <form method="post" enctype="multipart/form-data" action="/convert">
        <input type="file" name="file" accept=".doc,.docx,.ppt,.pptx,.xls,.xlsx" required>
        <input type="submit" value="Converter para PDF">
    </form>
    '''

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return 'Nenhum arquivo enviado!', 400

    file = request.files['file']
    if file.filename == '':
        return 'Nenhum arquivo selecionado!', 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)
    output_path = input_path.rsplit('.', 1)[0] + '.pdf'

    try:
        subprocess.run([
            LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf',
            '--outdir', UPLOAD_FOLDER, input_path
        ], check=True, capture_output=True, timeout=240)  # Timeout de 240s para conversão

        @after_this_request
        def cleanup(response):
            for _ in range(5):  # Tenta 5 vezes para limpar arquivos
                try:
                    if os.path.exists(input_path): os.remove(input_path)
                    if os.path.exists(output_path): os.remove(output_path)
                    break
                except Exception as e:
                    print(f"Erro ao limpar arquivos: {e}")
                    time.sleep(1)  # Espera 1s antes de tentar novamente
            return response

        return send_file(output_path, as_attachment=True, download_name='converted.pdf')

    except subprocess.TimeoutExpired:
        return 'Erro: Conversão demorou demais (timeout). Tente um arquivo menor.', 500
    except subprocess.CalledProcessError as e:
        return f'Erro na conversão: {e.stderr.decode()}', 500
    except FileNotFoundError:
        return 'Erro: LibreOffice não foi encontrado no caminho especificado.', 500

if __name__ == '__main__':
    app.run(debug=True)
