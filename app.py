from flask import Flask, render_template, request, send_file, jsonify
import os
from werkzeug.utils import secure_filename
from excel_processor import ExcelProcessor
import json

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Tạo thư mục nếu chưa có
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
os.makedirs('static/charts', exist_ok=True)

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        if 'files[]' not in request.files:
            return jsonify({'error': 'Không có file nào được chọn'}), 400
        
        files = request.files.getlist('files[]')
        
        if not files or files[0].filename == '':
            return jsonify({'error': 'Không có file nào được chọn'}), 400
        
        uploaded_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                uploaded_files.append(filepath)
        
        if not uploaded_files:
            return jsonify({'error': 'Không có file Excel hợp lệ'}), 400
        
        # Xử lý file Excel
        processor = ExcelProcessor(uploaded_files)
        result = processor.process()
        
        if 'error' in result:
            return jsonify(result), 400
        
        # Tạo file Excel output (có biểu đồ bên trong)
        output_file = processor.create_output_excel(result['data'])
        
        return jsonify({
            'success': True,
            'message': f'Đã xử lý thành công {len(result["skus"])} mã SKU',
            'skus': result['skus'],
            'output_file': output_file
        })
    
    except Exception as e:
        return jsonify({'error': f'Lỗi xử lý: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        filepath = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({'error': f'Lỗi tải file: {str(e)}'}), 404

@app.route('/view_charts/<sku>')
def view_charts(sku):
    return render_template('charts.html', sku=sku)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)




