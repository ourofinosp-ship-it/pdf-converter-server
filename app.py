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
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Conversor de Documentos para PDF</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                max-width: 800px;
                margin: 0 auto;
                padding: 20px;
                background-color: #f5f5f5;
            }
            .container {
                background: white;
                padding: 30px;
                border-radius: 10px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.1);
                text-align: center;
            }
            h2 {
                color: #2c3e50;
                margin-bottom: 30px;
            }
            .file-input {
                margin: 20px 0;
                padding: 15px;
                border: 2px dashed #3498db;
                border-radius: 5px;
                background-color: #f8f9fa;
            }
            .submit-btn {
                background-color: #3498db;
                color: white;
                padding: 12px 24px;
                border: none;
                border-radius: 5px;
                cursor: pointer;
                font-size: 16px;
                transition: background-color 0.3s;
            }
            .submit-btn:hover {
                background-color: #2980b9;
            }
            .supported-formats {
                margin-top: 20px;
                color: #7f8c8d;
                font-size: 14px;
            }
            .emoji {
                font-size: 24px;
                margin: 0 5px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h2>📄 Conversor de Documentos para PDF 📄</h2>
            <form method="post" enctype="multipart/form-data" action="/convert">
                <div class="file-input">
                    <input type="file" name="file" accept=".doc,.docx,.ppt,.pptx,.xls,.xlsx" required>
                </div>
                <input type="submit" value="🔄 Converter para PDF" class="submit-btn">
            </form>
            <div class="supported-formats">
                <p>📝 Formatos suportados: .doc, .docx, .ppt, .pptx, .xls, .xlsx</p>
                <p>⚡ Limite máximo: 500MB</p>
            </div>
        </div>
    </body>
    </html>
    '''

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return '❌ Nenhum arquivo enviado!', 400

    file = request.files['file']
    if file.filename == '':
        return '❌ Nenhum arquivo selecionado!', 400

    # Gera um nome de arquivo único para evitar conflitos
    timestamp = str(int(time.time()))
    original_filename = file.filename
    safe_filename = f"{timestamp}_{original_filename.replace(' ', '_')}"
    
    input_path = os.path.join(UPLOAD_FOLDER, safe_filename)
    file.save(input_path)
    output_filename = safe_filename.rsplit('.', 1)[0] + '.pdf'
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)

    try:
        # Executa a conversão com LibreOffice
        result = subprocess.run([
            LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf',
            '--outdir', UPLOAD_FOLDER, input_path
        ], check=True, capture_output=True, timeout=240)

        if not os.path.exists(output_path):
            return '❌ Erro: Arquivo PDF não foi gerado.', 500

        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(input_path):
                    os.remove(input_path)
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception as e:
                print(f"⚠️ Erro ao limpar arquivos: {e}")
            return response

        return send_file(
            output_path, 
            as_attachment=True, 
            download_name=f'converted_{original_filename.rsplit(".", 1)[0]}.pdf',
            mimetype='application/pdf'
        )

    except subprocess.TimeoutExpired:
        return '⏰ Erro: Conversão demorou demais (timeout). Tente um arquivo menor.', 500
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr.decode() if e.stderr else str(e)
        return f'❌ Erro na conversão: {error_msg}', 500
    except FileNotFoundError:
        return '❌ Erro: LibreOffice não foi encontrado.', 500
    except Exception as e:
        return f'❌ Erro inesperado: {str(e)}', 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
