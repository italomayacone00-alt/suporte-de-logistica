from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
from solver import resolver_problema_logistica
import os
from werkzeug.utils import secure_filename
import tempfile
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

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
        
        # Salvar na memória como Excel com Design Profissional
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Localizacao_CDs', index=False)
            worksheet = writer.sheets['Localizacao_CDs']
            
            # Criar aba de instruções
            from openpyxl import Workbook
            instructions_worksheet = writer.book.create_sheet('Como Usar')
            
            # --- DEFININDO OS ESTILOS PARA ABA DE INSTRUÇÕES ---
            # Cor do Cabeçalho (Azul Escuro) e Fonte Branca
            instructions_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            instructions_header_font = Font(color="FFFFFF", bold=True, size=12)
            
            # Cor de destaque (Azul Claro)
            highlight_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            instructions_bold_font = Font(bold=True, size=11)
            
            # Fonte normal
            normal_font = Font(size=10)
            
            # Alinhamento e Bordas
            center_align = Alignment(horizontal="center", vertical="center")
            left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            # --- ADICIONAR INSTRUÇÕES NA ABA "Como Usar" ---
            instructions_data = [
                ['📋 GUIA DE USO - PLANILHA DE OTIMIZAÇÃO DE CDs', ''],
                ['', ''],
                ['🎯 OBJETIVO', 'Encontrar a configuração ótima de Centros de Distribuição que minimiza custos totais'],
                ['', ''],
                ['📝 COMO PREENCHER A PLANILHA "Localizacao_CDs":', ''],
                ['1. CUSTOS DE TRANSPORTE (Células Cinzas)', 'Nas células cinza, informe o custo unitário (R$/unidade) para transportar mercadorias de cada CD para cada cliente'],
                ['2. CUSTO FIXO DOS CDs (Coluna "Custo Fixo")', 'Informe o custo mensal fixo (R$) para operar cada CD candidato'],
                ['3. CAPACIDADE DOS CDs (Coluna "Capacidade")', 'Informe a capacidade máxima (unidades/mês) que cada CD pode processar'],
                ['4. DEMANDA DOS CLIENTES (Linha "Demanda Total")', 'Informe a demanda mensal (unidades) que cada cliente precisa receber'],
                ['', ''],
                ['🎨 GUIA VISUAL DAS CORES:', ''],
                ['AZUL ESCURO (Cabeçalho)', 'Títulos das colunas - não editar'],
                ['AZUL CLARO (Primeira coluna)', 'Nomes dos CDs e linha de demanda - não editar'],
                ['CINZA CLARO (Miolo da tabela)', 'ÁREA PARA PREENCHER SEUS DADOS'],
                ['', ''],
                ['💡 EXEMPLO PRÁTICO:', ''],
                ['Para 2 CDs (SP, RJ) e 3 Clientes (A, B, C):', ''],
                ['• Custo SP → Cliente A: R$ 5,50/unidade', ''],
                ['• Custo Fixo SP: R$ 15.000/mês', ''],
                ['• Capacidade SP: 1.000 unidades/mês', ''],
                ['• Demanda Cliente A: 300 unidades/mês', ''],
                ['', ''],
                ['⚠️ DICAS IMPORTANTES:', ''],
                ['• Use números sem formatação (ex: 15000, não R$ 15.000)', ''],
                ['• Verifique se a capacidade total ≥ demanda total', ''],
                ['• Custos de transporte devem ser por unidade de produto', ''],
                ['• O sistema decidirá quais CDs abrir e fechar automaticamente', ''],
                ['', ''],
                ['🚀 PRÓXIMOS PASSOS:', ''],
                ['1. Preencha todos os dados na aba "Localizacao_CDs"', ''],
                ['2. Salve o arquivo Excel', ''],
                ['3. Faça upload no sistema', ''],
                ['4. Aguarde o resultado da otimização', ''],
            ]
            
            # Adicionar instruções na worksheet
            for row_idx, (col1, col2) in enumerate(instructions_data, 1):
                # Coluna 1
                cell1 = instructions_worksheet.cell(row=row_idx, column=1, value=col1)
                if row_idx == 1:
                    cell1.font = instructions_header_font
                    cell1.fill = instructions_header_fill
                    cell1.alignment = center_align
                elif col1.startswith('🎯') or col1.startswith('📝') or col1.startswith('🎨') or col1.startswith('💡') or col1.startswith('⚠️') or col1.startswith('🚀'):
                    cell1.font = instructions_bold_font
                    cell1.fill = highlight_fill
                    cell1.alignment = left_align
                else:
                    cell1.font = normal_font
                    cell1.alignment = left_align
                cell1.border = thin_border
                
                # Coluna 2
                cell2 = instructions_worksheet.cell(row=row_idx, column=2, value=col2)
                cell2.font = normal_font
                cell2.alignment = left_align
                cell2.border = thin_border
            
            # Ajustar largura das colunas da aba de instruções
            instructions_worksheet.column_dimensions['A'].width = 50
            instructions_worksheet.column_dimensions['B'].width = 80
            
            # Congelar painel para facilitar navegação
            instructions_worksheet.freeze_panes = 'A2'
            
            # --- DEFININDO OS ESTILOS PARA ABA DE DADOS ---
            # Cor do Cabeçalho (Azul Escuro) e Fonte Branca
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            # Cor da Primeira Coluna (Azul Claro) e Negrito
            col1_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            bold_font = Font(bold=True)
            
            # Cor da Área de Digitação (Cinza bem claro para guiar o usuário)
            input_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            
            # Alinhamento Centralizado e Bordas Finas
            center_align = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            # --- APLICANDO OS ESTILOS ---
            max_row = worksheet.max_row
            max_col = worksheet.max_column

            # 1. Estilizar o Cabeçalho (Linha 1)
            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = center_align
                cell.border = thin_border

            # 2. Estilizar o restante da tabela
            for row in range(2, max_row + 1):
                for col in range(1, max_col + 1):
                    cell = worksheet.cell(row=row, column=col)
                    cell.border = thin_border
                    cell.alignment = center_align
                    
                    # Se for a primeira coluna (Nomes dos CDs e Demanda)
                    if col == 1:
                        cell.font = bold_font
                        cell.fill = col1_fill
                    # Se for a área onde o usuário vai digitar os números
                    else:
                        cell.fill = input_fill

            # 3. Ajustar largura das colunas automaticamente
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter # Pega a letra da coluna (A, B, C...)
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                # Dá um respiro de tamanho (+4) para não ficar apertado
                worksheet.column_dimensions[column].width = max_length + 4

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
