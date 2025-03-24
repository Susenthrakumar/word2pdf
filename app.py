# app.py
from flask import Flask, request, render_template, send_file, jsonify
import os
import subprocess
import uuid
import time
import shutil

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
        # Generate unique filename to prevent collisions
        unique_id = str(uuid.uuid4())
        timestamp = int(time.time())
        file_id = f"{timestamp}_{unique_id}"
        
        # Save uploaded file with the unique ID
        input_filename = f"{file_id}_{file.filename}"
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        file.save(input_path)
        
        # Create output filename and path
        output_filename = f"{file_id}_{os.path.splitext(file.filename)[0]}.pdf"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        
        try:
            # Try multiple conversion methods until one succeeds
            conversion_success = False
            error_messages = []
            
            # Method 1: Try with LibreOffice
            try:
                conversion_success = convert_with_libreoffice(input_path, output_path)
            except Exception as e:
                error_messages.append(f"LibreOffice conversion failed: {str(e)}")
            
            # Method 2: Try with unoconv (if Method 1 failed)
            if not conversion_success:
                try:
                    conversion_success = convert_with_unoconv(input_path, output_path)
                except Exception as e:
                    error_messages.append(f"Unoconv conversion failed: {str(e)}")
            
            # Method 3: Try with python-docx and reportlab (if Methods 1 & 2 failed)
            if not conversion_success:
                try:
                    conversion_success = convert_with_python_docx(input_path, output_path)
                except Exception as e:
                    error_messages.append(f"Python-docx conversion failed: {str(e)}")
            
            # If no conversion method succeeded
            if not conversion_success:
                raise Exception("All conversion methods failed: " + " | ".join(error_messages))
            
            # Clean up the uploaded file
            os.remove(input_path)
            
            # Return the download URL
            return jsonify({
                'success': True,
                'filename': os.path.splitext(file.filename)[0] + '.pdf',
                'download_url': f'/download/{output_filename}'
            })
            
        except Exception as e:
            # If conversion fails, clean up and return error
            if os.path.exists(input_path):
                os.remove(input_path)
            return jsonify({'error': str(e)}), 500
    else:
        return jsonify({'error': 'Invalid file format. Please upload a Word document (.doc or .docx)'}), 400

def find_libreoffice_executable():
    """Find the LibreOffice executable on the system"""
    possible_executables = [
        'libreoffice',
        'soffice',
        '/usr/bin/libreoffice',
        '/usr/bin/soffice',
        '/usr/lib/libreoffice/program/soffice',
        '/opt/libreoffice*/program/soffice'
    ]
    
    for exe in possible_executables:
        # Handle wildcard paths
        if '*' in exe:
            import glob
            matches = glob.glob(exe)
            for match in matches:
                if os.path.isfile(match) and os.access(match, os.X_OK):
                    return match
        # Direct path or command name
        elif shutil.which(exe):
            return shutil.which(exe)
    
    raise FileNotFoundError("LibreOffice executable not found. Please install LibreOffice.")

def convert_with_libreoffice(input_path, output_path):
    """Convert a Word document to PDF using LibreOffice"""
    output_dir = os.path.dirname(output_path)
    
    try:
        # Find LibreOffice executable
        libreoffice_exec = find_libreoffice_executable()
        
        # Command to convert using LibreOffice
        cmd = [
            libreoffice_exec, 
            '--headless', 
            '--convert-to', 
            'pdf',
            '--outdir', 
            output_dir,
            input_path
        ]
        
        process = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if process.returncode != 0:
            raise Exception(f"LibreOffice returned error code {process.returncode}: {process.stderr}")
        
        # LibreOffice puts the output in the output_dir with original name but .pdf extension
        base_filename = os.path.basename(input_path)
        base_name_no_ext = os.path.splitext(base_filename)[0]
        temp_output = os.path.join(output_dir, f"{base_name_no_ext}.pdf")
        
        if os.path.exists(temp_output) and temp_output != output_path:
            os.rename(temp_output, output_path)
            return True
        elif os.path.exists(output_path):
            return True
        else:
            raise FileNotFoundError("Conversion completed but output file not found")
            
    except FileNotFoundError as e:
        raise e
    except Exception as e:
        raise Exception(f"LibreOffice conversion error: {str(e)}")

def convert_with_unoconv(input_path, output_path):
    """Convert a Word document to PDF using unoconv (another LibreOffice-based tool)"""
    try:
        # Check if unoconv is installed
        if not shutil.which('unoconv'):
            raise FileNotFoundError("unoconv not found. Install with 'pip install unoconv' or 'apt-get install unoconv'")
        
        # Run the conversion
        cmd = ['unoconv', '-f', 'pdf', '-o', output_path, input_path]
        process = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
        
        if process.returncode != 0:
            raise Exception(f"unoconv returned error code {process.returncode}: {process.stderr}")
        
        return os.path.exists(output_path)
        
    except Exception as e:
        raise Exception(f"unoconv conversion error: {str(e)}")

def convert_with_python_docx(input_path, output_path):
    """
    Basic conversion using python-docx and reportlab.
    Note: This will have limited formatting support compared to LibreOffice.
    """
    try:
        # Import necessary libraries - these should be added to requirements.txt
        from docx import Document
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        
        # Open the Word document
        doc = Document(input_path)
        
        # Create a new PDF
        c = canvas.Canvas(output_path, pagesize=letter)
        width, height = letter
        
        # Set initial position
        y = height - 40
        
        # Process each paragraph
        for para in doc.paragraphs:
            if not para.text.strip():
                y -= 12  # Skip some space for empty paragraphs
                continue
                
            # Wrap text to fit page width
            text = para.text
            c.drawString(40, y, text[:80])  # Simplified: just show first 80 chars
            
            y -= 12  # Move down for next line
            
            # Check if we need a new page
            if y < 40:
                c.showPage()
                y = height - 40
        
        # Save the PDF
        c.save()
        
        return os.path.exists(output_path)
        
    except ImportError:
        raise Exception("Required libraries not installed. Run 'pip install python-docx reportlab'")
    except Exception as e:
        raise Exception(f"python-docx conversion error: {str(e)}")

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['OUTPUT_FOLDER'], filename), 
                     as_attachment=True, 
                     download_name=filename.split('_', 2)[2])  # Remove the unique ID prefix

# Clean up old files periodically (you might want to implement this as a background task)
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
                # Skip files that don't follow the naming convention or can't be deleted
                continue
    
    return jsonify({'success': True, 'deleted_count': deleted_count})

if __name__ == '__main__':
    app.run(debug=True)
