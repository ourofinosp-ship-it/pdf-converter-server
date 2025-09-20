from flask import Flask, request, send_file, after_this_request, render_template_string
import subprocess
import os
import tempfile
import time
import uuid

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # Limite de 500MB

# Caminho para o LibreOffice
LIBREOFFICE_PATH = "libreoffice"

# Template HTML com interface moderna e emojis Twemoji
HTML_TEMPLATE = '''
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Conversor de Documentos para PDF</title>
    
    <!-- Twemoji CDN para emojis coloridos -->
    <script src="https://twemoji.maxcdn.com/v/latest/twemoji.min.js" crossorigin="anonymous"></script>
    
    <link href="https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Poppins', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: #333;
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
            width: 100%;
            max-width: 600px;
            padding: 40px;
            text-align: center;
        }
        
        .logo {
            font-size: 3.5rem;
            margin-bottom: 20px;
        }
        
        h1 {
            color: #333;
            margin-bottom: 10px;
            font-weight: 600;
        }
        
        .description {
            color: #666;
            margin-bottom: 30px;
            font-size: 1.1rem;
            line-height: 1.6;
        }
        
        .upload-container {
            background: #f8f9fa;
            border: 2px dashed #ced4da;
            border-radius: 10px;
            padding: 30px;
            margin-bottom: 25px;
            transition: all 0.3s ease;
        }
        
        .upload-container:hover {
            border-color: #667eea;
            background: #f0f4ff;
        }
        
        .file-input {
            display: none;
        }
        
        .file-label {
            display: flex;
            flex-direction: column;
            align-items: center;
            cursor: pointer;
        }
        
        .file-icon {
            font-size: 3rem;
            margin-bottom: 15px;
        }
        
        .file-text {
            font-size: 1.2rem;
            color: #495057;
            margin-bottom: 10px;
        }
        
        .file-hint {
            color: #6c757d;
            font-size: 0.9rem;
        }
        
        .submit-btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            border-radius: 50px;
            padding: 15px 30px;
            font-size: 1.1rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.3s ease;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 100%;
        }
        
        .submit-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        
        .submit-btn:active {
            transform: translateY(0);
        }
        
        .btn-icon {
            margin-right: 10px;
        }
        
        .supported-formats {
            margin-top: 30px;
            padding-top: 20px;
            border-top: 1px solid #e9ecef;
        }
        
        .formats-title {
            font-size: 1rem;
            color: #495057;
            margin-bottom: 15px;
        }
        
        .format-icons {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin-bottom: 15px;
        }
        
        .format-icon {
            font-size: 2rem;
        }
        
        .format-names {
            display: flex;
            justify-content: center;
            flex-wrap: wrap;
            gap: 15px;
            font-size: 0.9rem;
            color: #6c757d;
        }
        
        .format-badge {
            background: #f8f9fa;
            padding: 5px 12px;
            border-radius: 20px;
            border: 1px solid #e9ecef;
        }
        
        .footer {
            margin-top: 30px;
            color: #fff;
            text-align: center;
            font-size: 0.9rem;
        }
        
        .footer a {
            color: #fff;
            text-decoration: none;
        }
        
        .alert {
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            display: flex;
            align-items: center;
        }
        
        .alert-error {
            background: #ffe6e6;
            color: #d9534f;
            border: 1px solid #f5c6cb;
        }
        
        .alert-success {
            background: #e6f7ee;
            color: #28a745;
            border: 1px solid #c3e6cb;
        }
        
        .alert-icon {
            margin-right: 10px;
            font-size: 1.2rem;
        }
        
        .twemoji {
            height: 1em;
            width: 1em;
            margin: 0 .05em 0 .1em;
            vertical-align: -0.1em;
        }
        
        @media (max-width: 768px) {
            .container {
                padding: 25px;
            }
            
            .logo {
                font-size: 2.5rem;
            }
            
            h1 {
                font-size: 1.5rem;
            }
            
            .upload-container {
                padding: 20px;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="logo">
            📄
        </div>
        
        <h1>Conversor de Documentos para PDF</h1>
        
        <p class="description">Converta seus documentos para PDF de forma rápida, simples e gratuita.</p>
        
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                {% for category, message in messages %}
                    <div class="alert alert-{{ category }}">
                        <span class="alert-icon">
                            {% if category == 'error' %}
                                ⚠️
                            {% else %}
                                ✅
                            {% endif %}
                        </span>
                        {{ message }}
                    </div>
                {% endfor %}
            {% endif %}
        {% endwith %}
        
        <form method="post" enctype="multipart/form-data" action="/convert">
            <div class="upload-container">
                <input type="file" name="file" id="file" class="file-input" accept=".doc,.docx,.ppt,.pptx,.xls,.xlsx" required>
                <label for="file" class="file-label">
                    <span class="file-icon">📤</span>
                    <span class="file-text">Selecione seu documento</span>
                    <span class="file-hint">Clique ou arraste o arquivo até aqui</span>
                </label>
            </div>
            
            <button type="submit" class="submit-btn">
                <span class="btn-icon">🔄</span> Converter para PDF
            </button>
        </form>
        
        <div class="supported-formats">
            <p class="formats-title">Formatos suportados:</p>
            <div class="format-icons">
                <span class="format-icon" title="Word">📝</span>
                <span class="format-icon" title="PowerPoint">📊</span>
                <span class="format-icon" title="Excel">📈</span>
            </div>
            <div class="format-names">
                <span class="format-badge">.doc</span>
                <span class="format-badge">.docx</span>
                <span class="format-badge">.ppt</span>
                <span class="format-badge">.pptx</span>
                <span class="format-badge">.xls</span>
                <span class="format-badge">.xlsx</span>
            </div>
        </div>
    </div>
    
    <div class="footer">
        <p>Desenvolvido com ❤️ usando Flask e LibreOffice</p>
    </div>

    <script>
        // Inicializar Twemoji para converter todos os emojis para versões coloridas
        twemoji.parse(document.body, {
            folder: 'svg',
            ext: '.svg'
        });
        
        // Drag and drop functionality
        const fileInput = document.getElementById('file');
        const fileLabel = document.querySelector('.file-label');
        const fileText = document.querySelector('.file-text');
        const fileHint = document.querySelector('.file-hint');
        
        fileLabel.addEventListener('dragover', (e) => {
            e.preventDefault();
            fileLabel.style.borderColor = '#667eea';
            fileLabel.style.backgroundColor = '#f0f4ff';
        });
        
        fileLabel.addEventListener('dragleave', () => {
            fileLabel.style.borderColor = '#ced4da';
            fileLabel.style.backgroundColor = '#f8f9fa';
        });
        
        fileLabel.addEventListener('drop', (e) => {
            e.preventDefault();
            fileLabel.style.borderColor = '#ced4da';
            fileLabel.style.backgroundColor = '#f8f9fa';
            
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                updateFileName();
            }
        });
        
        fileInput.addEventListener('change', updateFileName);
        
        function updateFileName() {
            if (fileInput.files.length) {
                fileText.textContent = fileInput.files[0].name;
                fileHint.textContent = 'Clique ou arraste um arquivo diferente para alterar';
            } else {
                fileText.textContent = 'Selecione seu documento';
                fileHint.textContent = 'Clique ou arraste o arquivo até aqui';
            }
        }
        
        // Form submission feedback
        const form = document.querySelector('form');
        const submitBtn = document.querySelector('.submit-btn');
        
        form.addEventListener('submit', () => {
            if (fileInput.files.length) {
                submitBtn.innerHTML = '<span class="btn-icon">⏳</span> Convertendo...';
                submitBtn.disabled = true;
                
                // Atualizar os emojis após mudar o conteúdo do botão
                setTimeout(() => {
                    twemoji.parse(submitBtn, {
                        folder: 'svg',
                        ext: '.svg'
                    });
                }, 100);
            }
        });
    </script>
</body>
</html>
'''

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return render_template_string(HTML_TEMPLATE, error='Nenhum arquivo enviado!'), 400

    file = request.files['file']
    if file.filename == '':
        return render_template_string(HTML_TEMPLATE, error='Nenhum arquivo selecionado!'), 400

    # Generate unique filename to avoid conflicts
    unique_id = str(uuid.uuid4())[:8]
    original_filename = file.filename
    safe_filename = f"{unique_id}_{original_filename.replace(' ', '_')}"
    
    input_path = os.path.join(UPLOAD_FOLDER, safe_filename)
    file.save(input_path)
    output_filename = safe_filename.rsplit('.', 1)[0] + '.pdf'
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)

    try:
        # Execute conversion with LibreOffice
        result = subprocess.run([
            LIBREOFFICE_PATH, '--headless', '--convert-to', 'pdf',
            '--outdir', UPLOAD_FOLDER, input_path
        ], check=True, capture_output=True, timeout=240)

        if not os.path.exists(output_path):
            return render_template_string(HTML_TEMPLATE, error='Erro: Arquivo PDF não foi gerado.'), 500

        @after_this_request
        def cleanup(response):
            try:
                if os.path.exists(input_path):
                    os.remove(input_path)
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception as e:
                print(f"Erro ao limpar arquivos: {e}")
            return response

        return send_file(
            output_path, 
            as_attachment=True, 
            download_name=f'converted_{original_filename.rsplit(".", 1)[0]}.pdf',
            mimetype='application/pdf'
        )

    except subprocess.TimeoutExpired:
        return render_template_string(HTML_TEMPLATE, error='Erro: Conversão demorou demais (timeout). Tente um arquivo menor.'), 500
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr.decode() if e.stderr else str(e)
        return render_template_string(HTML_TEMPLATE, error=f'Erro na conversão: {error_msg}'), 500
    except FileNotFoundError:
        return render_template_string(HTML_TEMPLATE, error='Erro: LibreOffice não foi encontrado.'), 500
    except Exception as e:
        return render_template_string(HTML_TEMPLATE, error=f'Erro inesperado: {str(e)}'), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
