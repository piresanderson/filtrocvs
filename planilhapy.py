from flask import Flask, render_template, request, redirect, flash, send_file
import os
import pandas as pd
import zipfile
from waitress import serve  # Importando o Waitress

app = Flask(__name__)
app.secret_key = "supersecretkey"

# Configurações de upload
UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

# Verifica se o arquivo é permitido
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Função para processar as planilhas
def processar_planilha(file_path):
    try:
        # Carregar a planilha Excel
        df = pd.read_excel(file_path, sheet_name='Sheet 1', header=None)

        # Identificar automaticamente a linha do cabeçalho
        header_row_index = df[df.iloc[:, 0] == "Nome"].index[0]
        df.columns = df.iloc[header_row_index]  # Define os nomes das colunas com base na linha do cabeçalho
        df = df[header_row_index + 1:].reset_index(drop=True)  # Remove as linhas acima do cabeçalho

        # Encontrar dinamicamente as colunas pelos nomes reais
        coluna_nome = "Nome"
        coluna_sobrenome = "Sobrenome"
        coluna_idade = "Idade"
        coluna_telefone = "Telefone"
        coluna_celular = "Celular"
        coluna_salario_minimo = "Salário mínimo"
        coluna_salario_maximo = "Salário máximo"
        coluna_experiencia_profissional = "Experiência profissional"

        # Selecionar as colunas relevantes (incluindo salário mínimo e máximo)
        df_filtered = df[[coluna_nome, coluna_sobrenome, coluna_idade, coluna_telefone,
                          coluna_celular, coluna_salario_minimo, coluna_salario_maximo,
                          coluna_experiencia_profissional]]

        # Tratar valores ausentes e inconsistências na coluna de idade
        df_filtered[coluna_idade] = df_filtered[coluna_idade].fillna('0 Anos')
        df_filtered[coluna_idade] = df_filtered[coluna_idade].str.extract(r'(\d+)').astype(float)

        # Função para verificar baixa rotatividade (menos de 3 empregos listados)
        def check_baixa_rotatividade(experiencias):
            if isinstance(experiencias, str):
                jobs = experiencias.split("---------")  # Separar por empregos listados
                return len(jobs) <= 3  # Considere baixa rotatividade se houver até 3 empregos
            return False

        # Filtrar candidatos com base nos critérios:
        filtered_candidates = df_filtered[
            (
                (df_filtered[coluna_idade] >= 17) &  # Idade mínima de 17 anos
                (df_filtered[coluna_idade] <= 45) &  # Idade máxima de 45 anos
                (
                    (df_filtered[coluna_idade] < 21) |  # Candidatos com menos de 21 anos
                    (
                        (df_filtered[coluna_experiencia_profissional].str.contains('Vendedor', na=False)) &  # Experiência como vendedor
                        (df_filtered[coluna_experiencia_profissional].apply(check_baixa_rotatividade))  # Baixa rotatividade
                    )
                )
            )
        ]

        # Salvar os resultados em um arquivo Excel na pasta 'processed'
        output_file_name = os.path.join(app.config['PROCESSED_FOLDER'], f"filtrado_{os.path.basename(file_path)}")
        filtered_candidates.to_excel(output_file_name, index=False, sheet_name='Filtrados')

        return output_file_name

    except Exception as e:
        print(f"Erro ao processar a planilha: {e}")
        return None

# Rota principal para upload e processamento das planilhas
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "files[]" not in request.files:
            flash("Nenhum arquivo foi enviado.")
            return redirect(request.url)

        files = request.files.getlist("files[]")
        
        if len(files) > 5:
            flash("Você só pode enviar até 5 arquivos por vez.")
            return redirect(request.url)

        processed_files = []
        
        for file in files:
            if file and allowed_file(file.filename):
                filename = file.filename
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)

                # Processar a planilha usando sua lógica personalizada
                processed_file = processar_planilha(file_path)
                if processed_file:
                    processed_files.append(processed_file)

        if processed_files:
            # Compactar os arquivos processados em um único arquivo ZIP para download direto
            zip_filename = os.path.join(app.config['PROCESSED_FOLDER'], "planilhas_processadas.zip")
            with zipfile.ZipFile(zip_filename, 'w') as zipf:
                for file in processed_files:
                    zipf.write(file, os.path.basename(file))

            return send_file(zip_filename, as_attachment=True)

    return render_template("index.html", processed_files=None)

# Adicionando a estrutura básica para iniciar o servidor Flask com Waitress
if __name__ == "__main__":
    # Garantir que as pastas necessárias existam
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)
    
    print("Iniciando servidor com Waitress...")
    serve(app, host="0.0.0.0", port=5000)
