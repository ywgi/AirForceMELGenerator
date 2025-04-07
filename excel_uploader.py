from flask import Flask, request, render_template
import pandas as pd
import os

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('upload.html', message='No file part')

        file = request.files['file']
        if file.filename == '':
            return render_template('upload.html', message='No file selected')

        if file and (
                file.filename.endswith('.csv') or file.filename.endswith('.xlsx') or file.filename.endswith('.xls')):
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            # Read and process the file
            if file.filename.endswith('.csv'):
                df = pd.read_csv(filepath)
            else:
                df = pd.read_excel(filepath)

            # You can process the data further here
            preview = df.head().to_html()
            os.remove(filepath)
            return render_template('upload.html', message='File successfully uploaded!', preview=preview)

        return render_template('upload.html', message='Invalid file type. Please upload CSV or Excel files only.')

    return render_template('upload.html')


if __name__ == '__main__':
    app.run(debug=True)