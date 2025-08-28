# app.py
from flask import Flask, request, render_template, send_file, jsonify, flash
import os
import tempfile
import uuid
from werkzeug.utils import secure_filename
from tip import create_final_payroll_report
import threading
import time

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'  # Change this to a random secret key
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Create uploads directory if it doesn't exist
UPLOAD_FOLDER = 'temp_uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def cleanup_file(filepath, delay=300):
    """Delete file after delay (default 5 minutes)"""
    def delete_file():
        time.sleep(delay)
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
                print(f"Cleaned up: {filepath}")
        except Exception as e:
            print(f"Error cleaning up {filepath}: {e}")
    
    thread = threading.Thread(target=delete_file)
    thread.daemon = True
    thread.start()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_report():
    try:
        # Get form data
        filename = request.form.get('filename', '').strip()
        if not filename:
            return jsonify({'error': 'Filename is required'}), 400
        
        # Get uploaded files
        hours_file = request.files.get('hoursFile')
        tips_file = request.files.get('tipsFile')
        
        if not hours_file or hours_file.filename == '':
            return jsonify({'error': 'Hours CSV file is required'}), 400
        
        # Generate unique filenames to avoid conflicts
        unique_id = str(uuid.uuid4())[:8]
        
        # Save hours file
        hours_filename = f"hours_{unique_id}_{secure_filename(hours_file.filename)}"
        hours_path = os.path.join(UPLOAD_FOLDER, hours_filename)
        hours_file.save(hours_path)
        
        # Save tips file if provided
        tips_path = None
        if tips_file and tips_file.filename != '':
            tips_filename = f"tips_{unique_id}_{secure_filename(tips_file.filename)}"
            tips_path = os.path.join(UPLOAD_FOLDER, tips_filename)
            tips_file.save(tips_path)
        
        # Generate output filename with original name preserved
        clean_filename = secure_filename(filename)
        output_filename = f"{unique_id}_{clean_filename}.xlsx"
        output_path = os.path.join(UPLOAD_FOLDER, output_filename)
        
        # Call your Python function
        create_final_payroll_report(hours_path, output_path, tips_path)
        
        # Schedule cleanup of input files
        cleanup_file(hours_path)
        if tips_path:
            cleanup_file(tips_path)
        
        # Note: Output file will be deleted immediately after download
        
        return jsonify({
            'success': True,
            'download_url': f'/download/{output_filename}',
            'original_name': f"{clean_filename}.xlsx",
            'message': 'Excel report generated successfully!'
        })
        
    except Exception as e:
        print(f"Error generating report: {e}")
        return jsonify({'error': f'Error processing files: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': 'File not found'}), 404
        
        # Extract original filename (after the unique ID and underscore)
        if '_' in filename:
            original_name = '_'.join(filename.split('_')[1:])  # Skip the unique ID part
        else:
            original_name = filename
        
        # Read file into memory first
        with open(file_path, 'rb') as f:
            file_data = f.read()
        
        # Delete the file immediately after reading it
        try:
            os.remove(file_path)
            print(f"File deleted immediately: {file_path}")
            
            # Delete all remaining files in temp_uploads folder
            for remaining_file in os.listdir(UPLOAD_FOLDER):
                remaining_path = os.path.join(UPLOAD_FOLDER, remaining_file)
                if os.path.isfile(remaining_path):
                    os.remove(remaining_path)
                    print(f"Deleted remaining file: {remaining_path}")
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
        
        # Create a BytesIO object to send the file data
        from io import BytesIO
        file_obj = BytesIO(file_data)
        
        # Send the file from memory
        return send_file(
            file_obj,
            as_attachment=True,
            download_name=original_name,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Error downloading file: {e}")
        return jsonify({'error': 'Error downloading file'}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)