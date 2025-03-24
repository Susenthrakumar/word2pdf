from flask import Flask, request, render_template, send_file, jsonify
import os
import uuid
import time
import subprocess

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# Create necessary directories if they don't exist
for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
    if not os.path.exists(folder):
        os.makedirs(folder)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if file and file.filename.endswith(('.docx', '.doc')):
        unique_id = str(uuid.uuid4())
        timestamp = int(time.time())
        file_id = f"{timestamp}_{unique_id}"
        
        input_filename = f"{file_id}_{file.filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        
        output_filename = f"{file_id}_{os.path.splitext(file.filename)[0]}.pdf"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        try:
            # Convert DOCX to PDF using unoconv (LibreOffice)
            subprocess.run(["unoconv", "-f", "pdf", "-o", output_path, input_path], check=True)

            os.remove(input_path)  # Delete the uploaded DOCX file
            
            return jsonify({
                'success': True,
                'filename': os.path.splitext(file.filename)[0] + '.pdf',
                'download_url': f'/download/{output_filename}'
            })
            
        except subprocess.CalledProcessError as e:
            os.remove(input_path)  # Cleanup if failed
            return jsonify({'error': 'Conversion failed: ' + str(e)}), 500
    else:
        return jsonify({'error': 'Invalid file format. Please upload a Word document (.doc or .docx)'}), 400

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename), 
                     as_attachment=True, 
                     download_name=filename.split('_', 2)[2])  # Remove the unique ID prefix

@app.route('/cleanup', methods=['POST'])
def cleanup():
    threshold_time = time.time() - (24 * 60 * 60)  # 24 hours ago
    
    deleted_count = 0
    for folder in [app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER']]:
        for filename in os.listdir(folder):
            file_path = os.path.join(folder, filename)
            try:
                # Extract timestamp from filename
                timestamp = int(filename.split('_')[0])
                if timestamp < threshold_time:
                    os.remove(file_path)
                    deleted_count += 1
            except (ValueError, IndexError, OSError):
                continue
    
    return jsonify({'success': True, 'deleted_count': deleted_count})

if __name__ == '__main__':
    app.run(debug=True)
