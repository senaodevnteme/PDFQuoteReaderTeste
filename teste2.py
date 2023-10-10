from flask import Flask, render_template, request, send_file
import openpyxl
import io
import re
import pdfplumber as pdftool

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])

def buscarTexto(filepath):
    with pdftool.open(filepath) as tool:
        for p_no, pagina in enumerate(tool.pages, 1):
            print('<--- página número', p_no, '--->')
            data = pagina.extract_text()
            print(data)

buscarTexto('sample.pdf')

        
    
if __name__ == '__main__':
    app.run(debug=True)
