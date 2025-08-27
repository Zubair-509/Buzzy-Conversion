import os
import logging
import uuid
import tempfile
from flask import Flask, render_template, request, send_file, jsonify, flash, redirect, url_for
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
import pdf2docx
from pathlib import Path
import PyPDF2

# Configure logging
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key-change-in-production")

# Configuration
UPLOAD_FOLDER = 'uploads'
CONVERTED_FOLDER = 'converted'
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB max file size
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['CONVERTED_FOLDER'] = CONVERTED_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# Ensure upload and converted directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(CONVERTED_FOLDER, exist_ok=True)

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def validate_pdf(file_path):
    """Validate that the uploaded file is a proper PDF."""
    try:
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            # Check if PDF has pages and is readable
            if len(pdf_reader.pages) == 0:
                return False, "PDF file appears to be empty"
            
            # Try to read first page to ensure it's not corrupted
            first_page = pdf_reader.pages[0]
            first_page.extract_text()  # This will raise an exception if corrupted
            
            return True, "Valid PDF file"
    except Exception as e:
        logging.error(f"PDF validation error: {str(e)}")
        return False, f"Invalid or corrupted PDF file: {str(e)}"

def convert_pdf_to_docx(pdf_path, docx_path):
    """Convert PDF to DOCX using pdf2docx library."""
    try:
        logging.info(f"Starting conversion: {pdf_path} -> {docx_path}")
        
        # Use pdf2docx converter
        cv = pdf2docx.Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        
        logging.info(f"Conversion completed successfully: {docx_path}")
        return True, "Conversion completed successfully"
        
    except Exception as e:
        logging.error(f"Conversion error: {str(e)}")
        return False, f"Conversion failed: {str(e)}"

@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Handle file upload and conversion."""
    try:
        # Check if file is in request
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'No file selected'})
        
        file = request.files['file']
        
        # Check if file is selected
        if file.filename == '':
            return jsonify({'success': False, 'error': 'No file selected'})
        
        # Check file extension
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': 'Only PDF files are allowed'})
        
        # Generate unique filename
        unique_id = str(uuid.uuid4())
        original_filename = secure_filename(file.filename or "document.pdf")
        pdf_filename = f"{unique_id}_{original_filename}"
        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], pdf_filename)
        
        # Save uploaded file
        file.save(pdf_path)
        logging.info(f"File saved: {pdf_path}")
        
        # Validate PDF file
        is_valid, validation_message = validate_pdf(pdf_path)
        if not is_valid:
            # Clean up uploaded file
            os.remove(pdf_path)
            return jsonify({'success': False, 'error': validation_message})
        
        # Generate output filename
        docx_filename = f"{unique_id}_{Path(original_filename).stem}.docx"
        docx_path = os.path.join(app.config['CONVERTED_FOLDER'], docx_filename)
        
        # Convert PDF to DOCX
        conversion_success, conversion_message = convert_pdf_to_docx(pdf_path, docx_path)
        
        # Clean up uploaded PDF file
        os.remove(pdf_path)
        
        if not conversion_success:
            return jsonify({'success': False, 'error': conversion_message})
        
        # Return success with download URL
        return jsonify({
            'success': True, 
            'message': 'File converted successfully',
            'download_url': url_for('download_file', filename=docx_filename),
            'filename': f"{Path(original_filename).stem}.docx"
        })
        
    except RequestEntityTooLarge:
        return jsonify({'success': False, 'error': 'File too large. Maximum size is 50MB.'})
    except Exception as e:
        logging.error(f"Upload error: {str(e)}")
        # Clean up files if they exist
        try:
            if 'pdf_path' in locals() and 'pdf_path' in vars():
                if os.path.exists(pdf_path):
                    os.remove(pdf_path)
        except:
            pass
        return jsonify({'success': False, 'error': f'An error occurred: {str(e)}'})

@app.route('/download/<filename>')
def download_file(filename):
    """Handle file download."""
    try:
        file_path = os.path.join(app.config['CONVERTED_FOLDER'], filename)
        
        if not os.path.exists(file_path):
            flash('File not found or expired', 'error')
            return redirect(url_for('index'))
        
        def cleanup_file():
            """Clean up the file after download."""
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                    logging.info(f"Cleaned up file: {file_path}")
            except Exception as e:
                logging.error(f"Error cleaning up file: {str(e)}")
        
        # Schedule cleanup after response
        @app.after_request
        def remove_file(response):
            try:
                cleanup_file()
            except Exception as e:
                logging.error(f"Error in cleanup: {str(e)}")
            return response
        
        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename.split('_', 1)[1] if '_' in filename else filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        
    except Exception as e:
        logging.error(f"Download error: {str(e)}")
        flash('Error downloading file', 'error')
        return redirect(url_for('index'))

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error."""
    return jsonify({'success': False, 'error': 'File too large. Maximum size is 50MB.'}), 413

@app.errorhandler(500)
def internal_error(e):
    """Handle internal server errors."""
    logging.error(f"Internal server error: {str(e)}")
    return jsonify({'success': False, 'error': 'Internal server error occurred'}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
