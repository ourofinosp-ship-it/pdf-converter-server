from flask import Flask, request, send_file, after_this_request, render_template_string, jsonify
import subprocess
import os
import tempfile
import time
import uuid
import threading
import logging
from werkzeug.utils import secure_filename

# Configurar logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = Flask(__name__)
UPLOAD_FOLDER = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # Limite de 500MB

# Caminho para o LibreOffice
LIBREOFFICE_PATH = "libreoffice"

# Dicionário para acompanhar conversões em andamento
conversion_status = {}

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
    
    <!-- Bootstrap para melhorar o visual -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    
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
            max-width: 700px;
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
        
        .file-info {
            margin-top: 15px;
            padding: 10px;
            background: #e9ecef;
            border-radius: 5px;
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
            margin-top: 20px;
        }
        
        .submit-btn:hover:not(:disabled) {
            transform: translateY(-2px);
            box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        
        .submit-btn:disabled {
            background: #6c757d;
            cursor: not-allowed;
        }
        
        .submit-btn:active {
            transform: translateY(0);
        }
        
        .btn-icon {
            margin-right: 10px;
        }
        
        .progress-container {
            margin-top: 20px;
            display: none;
        }
        
        .progress-info {
            display: flex;
            justify-content: space-between;
            margin-bottom: 5px;
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
        
        .time-estimate {
            font-size: 0.85rem;
            color: #6c757d;
            margin-top: 5px;
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
        
        <div id="alerts">
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
        </div>
        
        <form id="conversion-form" method="post" enctype="multipart/form-data" action="/convert">
            <div class="upload-container">
                <input type="file" name="file" id="file" class="file-input" accept=".doc,.docx,.ppt,.pptx,.xls,.xlsx,.odt,.ods,.odp" required>
                <label for="file" class="file-label">
                    <span class="file-icon">📤</span>
                    <span class="file-text">Selecione seu documento</span>
                    <span class="file-hint">Clique ou arraste o arquivo até aqui</span>
                </label>
                <div id="file-info" class="file-info" style="display: none;">
                    <span id="file-name"></span> • <span id="file-size"></span>
                </div>
            </div>
            
            <div class="progress-container" id="progress-container">
                <div class="progress-info">
                    <span>Convertendo...</span>
                    <span id="progress-percentage">0%</span>
                </div>
                <div class="progress">
                    <div id="progress-bar" class="progress-bar progress-bar-striped progress-bar-animated" 
                         role="progressbar" style="width: 0%" aria-valuenow="0" aria-valuemin="0" aria-valuemax="100"></div>
                </div>
                <div class="time-estimate" id="time-estimate">
                    Tempo estimado: calculando...
                </div>
            </div>
            
            <button type="submit" class="submit-btn" id="submit-btn">
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
                <span class="format-badge">.odt</span>
                <span class="format-badge">.odp</span>
                <span class="format-badge">.ods</span>
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
        
        // Elementos DOM
        const fileInput = document.getElementById('file');
        const fileLabel = document.querySelector('.file-label');
        const fileText = document.querySelector('.file-text');
        const fileHint = document.querySelector('.file-hint');
        const fileInfo = document.getElementById('file-info');
        const fileName = document.getElementById('file-name');
        const fileSize = document.getElementById('file-size');
        const form = document.getElementById('conversion-form');
        const submitBtn = document.getElementById('submit-btn');
        const progressContainer = document.getElementById('progress-container');
        const progressBar = document.getElementById('progress-bar');
        const progressPercentage = document.getElementById('progress-percentage');
        const timeEstimate = document.getElementById('time-estimate');
        
        // Variáveis para controle de progresso
        let progressInterval;
        let startTime;
        let conversionId = null;
        
        // Formatador de tamanho de arquivo
        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }
        
        // Atualizar informações do arquivo
        function updateFileInfo() {
            if (fileInput.files.length) {
                const file = fileInput.files[0];
                fileName.textContent = file.name;
                fileSize.textContent = formatFileSize(file.size);
                fileInfo.style.display = 'block';
                
                // Estimar tempo baseado no tamanho do arquivo (aproximadamente 4MB por segundo)
                const estimatedTime = Math.max(10, Math.min(300, Math.round(file.size / (4 * 1024 * 1024))));
                timeEstimate.textContent = `Tempo estimado: ${estimatedTime} segundos para arquivos deste tamanho`;
            } else {
                fileInfo.style.display = 'none';
            }
        }
        
        // Drag and drop functionality
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
                updateFileInfo();
            }
        });
        
        fileInput.addEventListener('change', () => {
            updateFileName();
            updateFileInfo();
        });
        
        function updateFileName() {
            if (fileInput.files.length) {
                fileText.textContent = fileInput.files[0].name;
                fileHint.textContent = 'Clique ou arraste um arquivo diferente para alterar';
            } else {
                fileText.textContent = 'Selecione seu documento';
                fileHint.textContent = 'Clique ou arraste o arquivo até aqui';
            }
        }
        
        // Simular progresso da conversão
        function simulateProgress() {
            startTime = Date.now();
            let progress = 0;
            
            progressInterval = setInterval(() => {
                if (progress < 90) { // Para em 90% e espera pela conclusão real
                    progress += 1;
                    const elapsedSeconds = ((Date.now() - startTime) / 1000).toFixed(1);
                    const remaining = Math.max(1, Math.round((100 - progress) * elapsedSeconds / progress));
                    
                    progressBar.style.width = `${progress}%`;
                    progressBar.setAttribute('aria-valuenow', progress);
                    progressPercentage.textContent = `${progress}%`;
                    timeEstimate.textContent = `Tempo estimado: ${remaining} segundos restantes`;
                }
            }, 500);
        }
        
        // Verificar status da conversão
        function checkConversionStatus() {
            if (!conversionId) return;
            
            fetch(`/conversion-status/${conversionId}`)
                .then(response => response.json())
                .then(data => {
                    if (data.status === 'completed') {
                        clearInterval(progressInterval);
                        progressBar.style.width = '100%';
                        progressBar.setAttribute('aria-valuenow', 100);
                        progressPercentage.textContent = '100%';
                        timeEstimate.textContent = 'Conversão concluída!';
                        
                        // Redirecionar para download após breve delay
                        setTimeout(() => {
                            window.location.href = `/download/${conversionId}`;
                        }, 1000);
                    } else if (data.status === 'error') {
                        clearInterval(progressInterval);
                        showError(data.message || 'Erro na conversão');
                        submitBtn.disabled = false;
                        submitBtn.innerHTML = '<span class="btn-icon">🔄</span> Tentar novamente';
                        twemoji.parse(submitBtn);
                    }
                    // Se ainda estiver processando, continuamos verificando
                    else if (data.status === 'processing') {
                        setTimeout(checkConversionStatus, 2000);
                    }
                })
                .catch(error => {
                    console.error('Erro ao verificar status:', error);
                    setTimeout(checkConversionStatus, 2000);
                });
        }
        
        // Mostrar erro
        function showError(message) {
            const alertsDiv = document.getElementById('alerts');
            alertsDiv.innerHTML = `
                <div class="alert alert-error">
                    <span class="alert-icon">⚠️</span>
                    ${message}
                </div>
            `;
            twemoji.parse(alertsDiv);
        }
        
        // Form submission
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            
            if (!fileInput.files.length) {
                showError('Por favor, selecione um arquivo primeiro.');
                return;
            }
            
            const formData = new FormData(form);
            
            // Mostrar progresso
            progressContainer.style.display = 'block';
            submitBtn.disabled = true;
            submitBtn.innerHTML = '<span class="btn-icon">⏳</span> Processando...';
            twemoji.parse(submitBtn);
            
            // Iniciar simulação de progresso
            simulateProgress();
            
            try {
                const response = await fetch('/convert', {
                    method: 'POST',
                    body: formData
                });
                
                if (response.ok) {
                    const data = await response.json();
                    conversionId = data.conversion_id;
                    
                    // Iniciar verificação de status
                    setTimeout(checkConversionStatus, 2000);
                } else {
                    const errorData = await response.json();
                    clearInterval(progressInterval);
                    showError(errorData.error || 'Erro no servidor');
                    submitBtn.disabled = false;
                    submitBtn.innerHTML = '<span class="btn-icon">🔄</span> Tentar novamente';
                    twemoji.parse(submitBtn);
                    progressContainer.style.display = 'none';
                }
            } catch (error) {
                clearInterval(progressInterval);
                showError('Erro de conexão. Verifique sua internet e tente novamente.');
                submitBtn.disabled = false;
                submitBtn.innerHTML = '<span class="btn-icon">🔄</span> Tentar novamente';
                twemoji.parse(submitBtn);
                progressContainer.style.display = 'none';
            }
        });
        
        // Inicializar informações da página
        updateFileInfo();
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
        return jsonify({'error': 'Nenhum arquivo enviado!'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado!'}), 400

    # Validar tipo de arquivo
    allowed_extensions = {'.doc', '.docx', '.ppt', '.pptx', '.xls', '.xlsx', '.odt', '.ods', '.odp'}
    file_ext = os.path.splitext(file.filename)[1].lower()
    if file_ext not in allowed_extensions:
        return jsonify({'error': 'Tipo de arquivo não suportado!'}), 400

    # Gerar ID único para a conversão
    conversion_id = str(uuid.uuid4())
    original_filename = secure_filename(file.filename)
    safe_filename = f"{conversion_id}_{original_filename}"
    
    input_path = os.path.join(UPLOAD_FOLDER, safe_filename)
    file.save(input_path)
    output_filename = safe_filename.rsplit('.', 1)[0] + '.pdf'
    output_path = os.path.join(UPLOAD_FOLDER, output_filename)

    # Registrar conversão em andamento
    conversion_status[conversion_id] = {
        'status': 'processing',
        'input_path': input_path,
        'output_path': output_path,
        'original_filename': original_filename,
        'start_time': time.time()
    }

    # Executar conversão em uma thread separada
    def convert_document():
        try:
            # Comando otimizado para conversão
            result = subprocess.run([
                LIBREOFFICE_PATH, 
                '--headless',
                '--convert-to', 'pdf:writer_pdf_Export',
                '--outdir', UPLOAD_FOLDER, 
                input_path
            ], check=True, capture_output=True, timeout=300)

            if os.path.exists(output_path):
                conversion_status[conversion_id]['status'] = 'completed'
                conversion_status[conversion_id]['end_time'] = time.time()
                logger.info(f"Conversão concluída: {original_filename} -> {output_filename}")
            else:
                conversion_status[conversion_id]['status'] = 'error'
                conversion_status[conversion_id]['error'] = 'Arquivo PDF não foi gerado.'
                logger.error(f"Erro na conversão: PDF não gerado para {original_filename}")

        except subprocess.TimeoutExpired:
            conversion_status[conversion_id]['status'] = 'error'
            conversion_status[conversion_id]['error'] = 'Conversão excedeu o tempo limite.'
            logger.error(f"Timeout na conversão: {original_filename}")
        except subprocess.CalledProcessError as e:
            error_msg = e.stderr.decode() if e.stderr else str(e)
            conversion_status[conversion_id]['status'] = 'error'
            conversion_status[conversion_id]['error'] = f'Erro na conversão: {error_msg}'
            logger.error(f"Erro na conversão: {error_msg}")
        except Exception as e:
            conversion_status[conversion_id]['status'] = 'error'
            conversion_status[conversion_id]['error'] = f'Erro inesperado: {str(e)}'
            logger.error(f"Erro inesperado: {str(e)}")

    # Iniciar a conversão em thread separada
    thread = threading.Thread(target=convert_document)
    thread.start()

    return jsonify({'conversion_id': conversion_id})

@app.route('/conversion-status/<conversion_id>')
def get_conversion_status(conversion_id):
    if conversion_id not in conversion_status:
        return jsonify({'error': 'ID de conversão não encontrado'}), 404
    
    status_info = conversion_status[conversion_id]
    return jsonify(status_info)

@app.route('/download/<conversion_id>')
def download_file(conversion_id):
    if conversion_id not in conversion_status:
        return jsonify({'error': 'ID de conversão não encontrado'}), 404
    
    status_info = conversion_status[conversion_id]
    
    if status_info['status'] != 'completed':
        return jsonify({'error': 'Conversão não concluída'}), 400
    
    output_path = status_info['output_path']
    original_filename = status_info['original_filename']
    download_name = f'converted_{original_filename.rsplit(".", 1)[0]}.pdf'

    @after_this_request
    def cleanup(response):
        try:
            # Limpar arquivos após o download
            if os.path.exists(status_info['input_path']):
                os.remove(status_info['input_path'])
            if os.path.exists(status_info['output_path']):
                os.remove(status_info['output_path'])
            # Remover do registro de status
            if conversion_id in conversion_status:
                del conversion_status[conversion_id]
        except Exception as e:
            logger.error(f"Erro ao limpar arquivos: {e}")
        return response

    return send_file(
        output_path, 
        as_attachment=True, 
        download_name=download_name,
        mimetype='application/pdf'
    )

# Rota para limpar conversões antigas (executar periodicamente)
@app.route('/cleanup-old')
def cleanup_old():
    current_time = time.time()
    removed_count = 0
    
    for conv_id, info in list(conversion_status.items()):
        # Remover conversões com mais de 1 hora
        if current_time - info.get('start_time', 0) > 3600:
            try:
                if os.path.exists(info['input_path']):
                    os.remove(info['input_path'])
                if os.path.exists(info['output_path']):
                    os.remove(info['output_path'])
                del conversion_status[conv_id]
                removed_count += 1
            except Exception as e:
                logger.error(f"Erro ao limpar conversão antiga {conv_id}: {e}")
    
    return jsonify({'removed': removed_count})

if __name__ == '__main__':
    # Iniciar limpeza periódica em thread separada
    def periodic_cleanup():
        while True:
            time.sleep(3600)  # Executar a cada hora
            with app.app_context():
                cleanup_old()
    
    cleanup_thread = threading.Thread(target=periodic_cleanup)
    cleanup_thread.daemon = True
    cleanup_thread.start()
    
    app.run(debug=True, host='0.0.0.0', port=5000)
