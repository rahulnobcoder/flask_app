from flask import Flask, request, redirect, url_for, flash, render_template, send_file
from werkzeug.utils import secure_filename
import os
from extract.extraction import *

app = Flask(__name__)
app.secret_key = 'hleo'

# Configure upload folder and allowed extensions (optional)
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')  # Use os.path.join for cross-platform compatibility
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'docx', 'doc'}

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        # Get uploaded files from request.files dictionary
        uploaded_files = request.files.getlist('files')
        if uploaded_files:
            for file in uploaded_files:
                if file.filename == '':
                    flash('No selected file')
                    return redirect(request.url)
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    # Save each file to the upload folder
                    file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                    flash('Files uploaded successfully!')
                else:
                    flash('Allowed file types: ' + ', '.join(ALLOWED_EXTENSIONS))
            return render_template('download.html')
    return render_template('index.html')

@app.route('/download')
def download_file():
    path = 'uploads'
    df = get_df(path)
    df.to_excel('details.xlsx', index=False)
    file_path = 'details.xlsx'
    clear_folder_contents("uploads")
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
