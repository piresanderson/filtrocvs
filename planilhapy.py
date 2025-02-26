from flask import Flask, render_template, request, redirect, flash, send_file 
import os
import pandas as pd
from waitress import serve  # Importando o Waitress
from datetime import datetime

app = Flask(__name__)
app.secret_key = "supersecretkey"

UPLOAD_FOLDER = 'uploads'
PROCESSED_FOLDER = 'processed'
ALLOWED_EXTENSIONS = {'xlsx'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['PROCESSED_FOLDER'] = PROCESSED_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def processar_planilha(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_name = xls.sheet_names[0]
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

        header_row_index = df[df.iloc[:, 0] == "Nome"].index[0]
        df.columns = df.iloc[header_row_index]
        df = df[header_row_index + 1:].reset_index(drop=True)

        coluna_nome = "Nome"
        coluna_sobrenome = "Sobrenome"
        coluna_idade = "Idade"
        coluna_telefone = "Telefone"
        coluna_celular = "Celular"
        coluna_salario_minimo = "Salário mínimo"
        coluna_salario_maximo = "Salário máximo"
        coluna_experiencia_profissional = "Experiência profissional"
        coluna_treinamento = "Treinamento"
        coluna_cursos = "Cursos"

        if coluna_treinamento in df.columns:
            df.rename(columns={coluna_treinamento: coluna_cursos}, inplace=True)
        else:
            df[coluna_cursos] = "Não informado"

        df_filtered = df[[coluna_nome, coluna_sobrenome, coluna_idade, coluna_telefone,
                          coluna_celular, coluna_salario_minimo, coluna_salario_maximo,
                          coluna_experiencia_profissional, coluna_cursos]]

        df_filtered[coluna_idade] = df_filtered[coluna_idade].fillna('0 Anos')
        df_filtered[coluna_idade] = df_filtered[coluna_idade].astype(str).str.extract(r'(\d+)').astype(float)

        cursos_excluir = ["Enfermagem", "Direito", "Medicina", "Advocacia", "Biomedicina"]

        def calcular_duracao_experiencia(experiencia):
            if "Atual" in experiencia:
                experiencia = experiencia.replace("Atual", datetime.today().strftime("%m/%Y"))
            return experiencia

        def verificar_rotatividade(idade, experiencia):
            if idade < 20:
                return True  # Permitir candidatos menores de 20 anos sem experiência
            
            if isinstance(experiencia, str):
                experiencia = calcular_duracao_experiencia(experiencia)
                empregos = experiencia.split("---------")
                total_empregos = len(empregos)
                
                if idade < 21:
                    return total_empregos <= 4
                elif 21 <= idade <= 23:
                    if total_empregos <= 4:
                        ultimas_exp = empregos[-2:] if len(empregos) >= 2 else empregos[-1:]
                        return any("1 ano" in exp or "mais de 1 ano" in exp for exp in ultimas_exp)
                    return False
                elif 25 <= idade <= 27:
                    return all("2 anos" in exp or "mais de 2 anos" in exp for exp in empregos[-2:])
                elif 27 <= idade <= 30:
                    return "3 anos" in empregos[-1] and "2 anos" in empregos[-2]
                elif 30 <= idade <= 35:
                    return any("4 anos" in exp for exp in empregos[-2:])
                elif 35 <= idade <= 45:
                    return any("5 anos" in exp for exp in empregos[-2:])
                
                exp_terceira = empregos[-3] if len(empregos) >= 3 else ""
                exp_quarta = empregos[-4] if len(empregos) >= 4 else ""
                
                if any("5 anos" in exp for exp in [exp_terceira, exp_quarta]) and all("6 meses" in exp or "menos de 6 meses" in exp for exp in empregos[-2:]):
                    return False
                
                return True
            return False

        def verificar_areas_experiencia(idade, experiencia):
            areas_desejadas = ["Vendedor", "Atendente", "Telemarketing", "Subgerente", "Gerente", "Consultor", "Vendas"]
            if isinstance(experiencia, str) and 23 <= idade <= 45:
                empregos = experiencia.split("---------")
                ultimas_exp = empregos[-3:] if len(empregos) >= 3 else empregos
                return any(any(area in exp for area in areas_desejadas) for exp in ultimas_exp)
            return True

        filtered_candidates = df_filtered[
            (
                (df_filtered[coluna_idade] >= 17) &  
                (df_filtered[coluna_idade] <= 45) &  
                (~df_filtered[coluna_cursos].str.contains('|'.join(cursos_excluir), case=False, na=False)) &  
                (df_filtered.apply(lambda x: verificar_rotatividade(x[coluna_idade], x[coluna_experiencia_profissional]), axis=1)) &
                (df_filtered.apply(lambda x: verificar_areas_experiencia(x[coluna_idade], x[coluna_experiencia_profissional]), axis=1))
            )
        ]

        output_file_name = os.path.join(app.config['PROCESSED_FOLDER'], f"filtrado_{os.path.basename(file_path)}")
        filtered_candidates.to_excel(output_file_name, index=False, sheet_name='Filtrados')

        return output_file_name

    except Exception as e:
        print(f"Erro ao processar a planilha: {e}")
        return None

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

        for file in files:
            if file and allowed_file(file.filename):
                filename = file.filename
                file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(file_path)

                processed_file = processar_planilha(file_path)
                if processed_file:
                    return send_file(processed_file, as_attachment=True)

    return render_template("index.html", processed_files=None)

if __name__ == "__main__":
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    os.makedirs(app.config['PROCESSED_FOLDER'], exist_ok=True)
    print("Iniciando servidor com Waitress...")
    serve(app, host="0.0.0.0", port=5000)
