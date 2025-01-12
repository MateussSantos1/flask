import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import unicodedata
from flask import Flask, render_template, request, redirect, url_for, jsonify
from flask_socketio import SocketIO, emit

app = Flask(__name__)
socketio = SocketIO(app)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    if 'pdf_file' not in request.files or 'excel_file' not in request.files:
        return "Por favor, envie ambos os arquivos (PDF e Excel).", 400

    pdf_file = request.files['pdf_file']
    excel_file = request.files['excel_file']
    output_directory = request.form.get('output_directory')
    sheet_name = request.form.get('sheet_name')

    if not output_directory:
        return "Por favor, forneça o diretório de saída.", 400

    if not sheet_name:
        return "Por favor, forneça o nome da aba.", 400

    try:
        # Salva os arquivos enviados
        pdf_path = os.path.join(output_directory, pdf_file.filename)
        excel_path = os.path.join(output_directory, excel_file.filename)
        pdf_file.save(pdf_path)
        excel_file.save(excel_path)

        # Processa os arquivos e captura os nomes não encontrados
        nomes_nao_encontrados, count_salvos, count_nao_salvos, output_nao_encontrados = process_files(pdf_path, excel_path, sheet_name, output_directory)

        # Redireciona para a página de resultados com os dados do processamento
        return render_template('results.html', count_salvos=count_salvos, count_nao_salvos=count_nao_salvos,
                               output_directory=output_directory, nomes_nao_encontrados=nomes_nao_encontrados,
                               output_nao_encontrados=output_nao_encontrados)

    except Exception as e:
        return jsonify(error=str(e)), 500

def process_files(pdf_path, excel_path, sheet_name, output_directory):
    aba_escolhida = pd.read_excel(excel_path, sheet_name=sheet_name)
    nomes = aba_escolhida["NOME"].tolist()
    pdf = PdfReader(pdf_path)

    nomes_nao_encontrados = []  # Lista para armazenar nomes não encontrados
    count_salvos = 0
    count_nao_salvos = 0

    for nome in nomes:
        nome = unicodedata.normalize('NFKD', nome).encode('ASCII', 'ignore').decode('ASCII')
        partes_nome = nome.split()
        encontrado = False

        for page in pdf.pages:
            texto = page.extract_text()
            if all(p in texto for p in partes_nome):
                pdf_writer = PdfWriter()
                pdf_writer.add_page(page)
                output_pdf_path = os.path.join(output_directory, f"{nome}.pdf")
                with open(output_pdf_path, "wb") as f:
                    pdf_writer.write(f)

                # Enviar feedback para o cliente
                socketio.emit('feedback', {'message': f'Comprovante de {nome} salvo.'})
                count_salvos += 1
                encontrado = True
                break  # Se encontrou, não precisa continuar buscando nas outras páginas

        if not encontrado:
            nomes_nao_encontrados.append(nome)
            count_nao_salvos += 1
            # Enviar feedback para o cliente
            socketio.emit('feedback', {'message': f'Comprovante de {nome} não encontrado.'})

    # Se houver nomes não encontrados, cria um arquivo Excel com a lista
    output_nao_encontrados = None
    if nomes_nao_encontrados:
        output_nao_encontrados = os.path.join(output_directory, "nomes_nao_encontrados.xlsx")
        df_nomes_nao_encontrados = pd.DataFrame({'NOMES NÃO ENCONTRADOS': nomes_nao_encontrados})
        df_nomes_nao_encontrados.to_excel(output_nao_encontrados, index=False)

    return nomes_nao_encontrados, count_salvos, count_nao_salvos, output_nao_encontrados

if __name__ == '__main__':
    socketio.run(app, debug=True)
