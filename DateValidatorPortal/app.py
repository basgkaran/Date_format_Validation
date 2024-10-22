from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import pandas as pd

app = Flask(__name__)

# Set upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

@app.route('/')
def index():
    return render_template('upload.html')  # Use 'upload.html' in the templates folder

@app.route('/upload', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files.get('file')  # Use .get() to avoid KeyError
        if file and file.filename.endswith('.xlsx'):
            # Save the uploaded file
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
            file.save(filepath)

            # Process the Excel file
            df = pd.read_excel(filepath)

            # Function to format dates
            def format_date(value):
                if pd.isnull(value):  # Check for NaT (Not a Time)
                    return ""
                # Convert to datetime if necessary
                if isinstance(value, str):
                    # Try to convert the string to a datetime object
                    try:
                        value = pd.to_datetime(value)
                    except ValueError:
                        return value  # Return the original value if conversion fails
                # Format the date as DD-MM-YYYY or DDM-MM-YYYY HH:MM:MM
                return value.strftime('%d-%m-%Y') if value.hour is None else value.strftime('%d-%m-%Y %H:%M:%S')

            # Apply date formatting to all columns in the dataframe
            for column in df.columns:
                df[column] = df[column].apply(format_date)

            # Save the processed file
            processed_filename = 'processed_' + file.filename
            processed_filepath = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
            df.to_excel(processed_filepath, index=False)

            # Redirect to download the processed file
            return redirect(url_for('download_file', filename=processed_filename))

    return "File upload failed", 400  # Return a bad request response

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)  
 