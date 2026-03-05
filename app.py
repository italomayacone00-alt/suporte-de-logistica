from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
from solver import resolver_problema_logistica
import os
from werkzeug.utils import secure_filename
import tempfile
import io

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'

# Configurações para upload de arquivos
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Criar pasta de uploads se não existir
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('dashboard_professional.html')

@app.route('/dashboard')
def dashboard():
    return render_template('dashboard_professional.html')

@app.route('/tools')
def tools():
    return render_template('index.html')

@app.route('/otimizador_cds')
def otimizador_cds():
    return render_template('otimizador_cds.html')

# --- ROTA PARA GERAR O TEMPLATE NO FORMATO ORIGINAL ---
@app.route('/gerar_template', methods=['POST'])
def gerar_template():
    try:
        # Pega os dados do formulário
        num_cds = int(request.form.get('num_cds', 0))
        num_clientes = int(request.form.get('num_clientes', 0))
        
        if num_cds <= 0 or num_clientes <= 0:
            flash('Número de CDs e Clientes deve ser maior que zero.', 'error')
            return redirect(url_for('otimizador_cds'))

        # Criar as listas de nomes (Linhas e Colunas)
        nomes_cds = [f'CD {i}' for i in range(1, num_cds + 1)]
        nomes_clientes = [f'Cliente {i}' for i in range(1, num_clientes + 1)]
        
        # Cabeçalho completo
        colunas = ['Origem / Destino'] + nomes_clientes + ['Custo Fixo', 'Capacidade']
        
        # Montar os dados das linhas
        dados = []
        for cd in nomes_cds:
            # Para cada CD: [Nome do CD, (espaços vazios pros clientes), espaço fixo, espaço cap]
            linha = [cd] + ['' for _ in nomes_clientes] + ['', '']
            dados.append(linha)
            
        # Adicionar a linha final de Demanda
        linha_demanda = ['Demanda Total'] + ['' for _ in nomes_clientes] + ['-', '-']
        dados.append(linha_demanda)
        
        # Criar o DataFrame (Tabela)
        df = pd.DataFrame(dados, columns=colunas)
        
        # Salvar na memória como Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Localizacao_CDs', index=False)
            
            # Ajustar largura das colunas para ficar bonito
            worksheet = writer.sheets['Localizacao_CDs']
            for col in worksheet.columns:
                max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
                worksheet.column_dimensions[col[0].column_letter].width = max_length + 2

        output.seek(0)
        
        return send_file(
            output,
            download_name="Template_Localizacao_CDs.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        flash(f'Erro ao gerar template: {str(e)}', 'error')
        return redirect(url_for('otimizador_cds'))

# --- ATUALIZAÇÃO NA ROTA DE UPLOAD ---
@app.route('/resolver', methods=['POST'])
def resolver():
    try:
        if 'arquivo_planilha' not in request.files:
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('otimizador_cds'))
        
        file = request.files['arquivo_planilha']
        
        if file.filename == '':
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('otimizador_cds'))
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Ler dados do Excel (agora lê apenas a primeira aba com a matriz)
            df_principal = pd.read_excel(filepath)
            
            # Resolver o problema passando a tabela matriz
            resultado = resolver_problema_logistica(df_principal)
            
            # Limpar arquivo temporário
            os.remove(filepath)
            
            if resultado.get('status') == 'Erro':
                flash(f"Erro na otimização: {resultado.get('mensagem')}", 'error')
                return redirect(url_for('otimizador_cds'))
            
            return render_template('resultado.html', resultado=resultado)
                
        else:
            flash('Formato de arquivo inválido. Por favor, envie um arquivo .xlsx.', 'error')
            return redirect(url_for('otimizador_cds'))
            
    except Exception as e:
        flash(f'Erro ao processar arquivo: {str(e)}', 'error')
        return redirect(url_for('otimizador_cds'))

if __name__ == '__main__':
    app.run(debug=True)
