from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import pandas as pd
import os
from werkzeug.utils import secure_filename
import tempfile
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
import zipfile

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'

# Configurações para upload de arquivos
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'csv'}
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

@app.route('/tradicional')
def tradicional():
    return render_template('otimizador_cds.html')

@app.route('/otimizador_cds')
def otimizador_cds():
    return render_template('otimizador_cds.html')

@app.route('/p_medianas')
def p_medianas():
    return render_template('p_medianas.html')

@app.route('/pmediana_simples')
def pmediana_simples():
    return render_template('pmediana_simples.html')

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
            
            # Resolver o problema tradicional
            from solver_tradicional import resolver_problema_logistica
            resultado = resolver_problema_logistica(df_principal)
            
            # Limpar arquivo temporário
            os.remove(filepath)
            
            if resultado.get('status') == 'Erro':
                flash(f"Erro na otimização: {resultado.get('mensagem')}", 'error')
                return redirect(url_for('otimizador_cds'))
            
            return render_template('resultado_tradicional.html', resultado=resultado)
                
        else:
            flash('Formato de arquivo inválido. Por favor, envie um arquivo .xlsx.', 'error')
            return redirect(url_for('otimizador_cds'))
            
    except Exception as e:
        flash(f'Erro ao processar arquivo: {str(e)}', 'error')
        return redirect(url_for('p_centros'))

def criar_excel_pcentros(df, num_cds, num_clientes, p_cds):
    """Função para criar Excel estilizado para p-Centros"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Dados', index=False)
        
        # Aplicar estilização igual às outras ferramentas
        worksheet = writer.sheets['Dados']
        max_row = worksheet.max_row
        max_col = worksheet.max_column
        
        # Definir estilos
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="9c27b0", end_color="9c27b0", fill_type="solid")  # Roxo para p-Centros
        center_align = Alignment(horizontal="center", vertical="center")
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                           top=Side(style='thin'), bottom=Side(style='thin'))
        col1_fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
        input_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
        bold_font = Font(bold=True)
        
        # 1. Estilizar cabeçalho
        for col in range(1, max_col + 1):
            cell = worksheet.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
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
        
        # 4. Adicionar aba de instruções
        instrucoes = pd.DataFrame([
            ['Como Usar', ''],
            ['', ''],
            ['1. Preencha os dados:', ''],
            ['- Linha 1: Cabeçalho com IDs (C1, C2, ..., CD1, CD2, ...)', ''],
            ['- Linhas seguintes: CDs (CD1, CD2, ...)', ''],
            ['- Valores: Distâncias entre CDs e clientes', ''],
            ['- Diagonal CD×CD: Mantenha zero (0)', ''],
            ['', ''],
            ['2. Salve o arquivo:', ''],
            ['- Clique em Arquivo > Salvar Como', ''],
            ['- Escolha formato .xlsx', ''],
            ['', ''],
            ['3. Faça upload:', ''],
            ['- Volte para a página do sistema', ''],
            ['- Clique em "Escolher arquivo" e selecione a planilha', ''],
            ['- Clique em "Otimizar" para calcular os p-Centros', ''],
            ['', ''],
            ['Observações:', ''],
            ['- p-Centros minimiza a MAIOR distância', ''],
            ['- Garante equidade no atendimento', ''],
            ['- Ideal para serviços emergenciais', ''],
            ['- Sem Demanda e sem Custo Fixo', ''],
            ['', ''],
            ['Parâmetros configurados:', ''],
            [f'- CDs candidatos: {num_cds}', ''],
            [f'- Clientes: {num_clientes}', ''],
            [f'- CDs a abrir: {p_cds}']
        ])
        instrucoes.to_excel(writer, sheet_name='Como Usar', index=False)
    
    return output

# --- ROTA PARA MÁXIMA COBERTURA ---
@app.route('/max_cobertura')
def max_cobertura():
    return render_template('max_cobertura.html')

@app.route('/gerar_template_maxcobertura', methods=['POST'])
def gerar_template_maxcobertura():
    try:
        num_cds = int(request.form.get('num_cds', 0))
        num_clientes = int(request.form.get('num_clientes', 0))
        
        if num_cds <= 0 or num_clientes <= 0:
            flash('Por favor, insira valores válidos para gerar o template.', 'error')
            return redirect(url_for('max_cobertura'))
            
        colunas = ['Origem / Destino'] + [f'Cliente {i}' for i in range(1, num_clientes + 1)]
        df_template = pd.DataFrame(columns=colunas)
        
        for i in range(1, num_cds + 1):
            df_template.loc[i-1] = [f'CD {i}'] + ['' for _ in range(num_clientes)]
            
        # Voltando ao padrão logístico: Demanda Total
        df_template.loc[num_cds] = ['Demanda Total'] + ['' for _ in range(num_clientes)]
            
        guia_uso = [
            ["📋 GUIA DE USO - MÁXIMA COBERTURA", ""],
            ["1. DISTÂNCIAS", "Informe a distância (ou tempo) entre o CD e o Cliente."],
            ["2. MATRIZ BINÁRIA", "Se não tiver distâncias exatas, coloque 1 (se o CD alcança) ou 0 (não alcança). Nesse caso, o Raio no site deve ser 0."],
            ["3. DEMANDA", "Na última linha, informe a demanda em quantidade de produtos/pedidos de cada cliente."]
        ]
        df_como_usar = pd.DataFrame(guia_uso)
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False, sheet_name='Maxima_Cobertura')
            df_como_usar.to_excel(writer, index=False, header=False, sheet_name='Como Usar')
            
            workbook = writer.book
            
            # --- ESTILIZANDO A ABA 1 (Máxima Cobertura) ---
            worksheet = writer.sheets['Maxima_Cobertura']
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') # Azul escuro
            cd_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')     # Azul claro
            input_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')  # Cinza
            demanda_fill = PatternFill(start_color='E8F5E8', end_color='E8F5E8', fill_type='solid') # Verde claro
            
            white_font = Font(color='FFFFFF', bold=True)
            bold_font = Font(bold=True)
            center_alignment = Alignment(horizontal='center', vertical='center')
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            # Cabeçalho
            for col_num, col_name in enumerate(colunas, 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = center_alignment
                cell.border = thin_border
                worksheet.column_dimensions[cell.column_letter].width = 15
                
            # CDs (linhas 2 até num_cds+1)
            for row_num in range(2, num_cds + 2):
                cd_cell = worksheet.cell(row=row_num, column=1)
                cd_cell.fill = cd_fill
                cd_cell.font = bold_font
                cd_cell.alignment = center_alignment
                cd_cell.border = thin_border
                
                for col_num in range(2, num_clientes + 2):
                    input_cell = worksheet.cell(row=row_num, column=col_num)
                    input_cell.fill = input_fill
                    input_cell.alignment = center_alignment
                    input_cell.border = thin_border
            
            # Linha de demanda (última linha)
            row_demanda = num_cds + 2
            demanda_cell = worksheet.cell(row=row_demanda, column=1)
            demanda_cell.fill = cd_fill
            demanda_cell.font = bold_font
            demanda_cell.alignment = center_alignment
            demanda_cell.border = thin_border
            
            for col_num in range(2, num_clientes + 2):
                demanda_input_cell = worksheet.cell(row=row_demanda, column=col_num)
                demanda_input_cell.fill = demanda_fill
                demanda_input_cell.alignment = center_alignment
                demanda_input_cell.border = thin_border

            # --- ESTILIZANDO A ABA 2 (Como Usar) ---
            ws_como_usar = writer.sheets['Como Usar']
            ws_como_usar.column_dimensions['A'].width = 45
            ws_como_usar.column_dimensions['B'].width = 90
            
            for row in range(1, len(guia_uso) + 1):
                cell_a = ws_como_usar.cell(row=row, column=1)
                cell_b = ws_como_usar.cell(row=row, column=2)
                
                # Colocar negrito na primeira coluna
                if cell_a.value:
                    cell_a.font = bold_font
                
                # Destacar os cabeçalhos principais
                if row in [1]:
                    cell_a.font = Font(bold=True, size=12, color='1F4E78')
                    
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Template_MaximaCobertura.xlsx'
        )
    except Exception as e:
        flash(f'Erro ao gerar template: {str(e)}', 'error')
        return redirect(url_for('max_cobertura'))

@app.route('/resolver_maxcobertura', methods=['POST'])
def resolver_maxcobertura_route():
    try:
        print(f"=== DEBUG - ROTA MÁXIMA COBERTURA ===")
        
        # Puxando as variáveis corretas do HTML (mesmo padrão do p-Centros)
        p = int(request.form.get('p_cds', 0))
        raio = float(request.form.get('raio_cobertura', 0))
        
        print(f"Parâmetros recebidos: p={p}, raio={raio}")
        
        file = request.files.get('file')
        print(f"Arquivo recebido: {file}")
        print(f"Nome do arquivo: {file.filename if file else 'None'}")
        
        if p <= 0:
            flash('Número de CDs a abrir deve ser maior que zero.', 'error')
            return redirect(url_for('max_cobertura'))
            
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            print(f"Arquivo salvo em: {filepath}")
            
            df = pd.read_excel(filepath)
            print(f"DataFrame lido com sucesso: {df.shape}")
            
            from solver_maxcobertura import resolver_maxcobertura
            resultado = resolver_maxcobertura(df, p, raio)
            
            print(f"Resultado do solver: {resultado}")
            
            os.remove(filepath)
            
            if resultado.get('status') == 'Erro':
                flash(f"Erro na otimização: {resultado.get('mensagem')}", 'error')
                return redirect(url_for('max_cobertura'))
            
            # Adicionar mensagem de sucesso
            if resultado.get('status') == 'Sucesso':
                cds_selecionados = resultado.get('cds_selecionados', [])
                demanda_coberta = resultado.get('demanda_coberta', 0)
                percentual = resultado.get('percentual_cobertura', 0)
                mensagem = f"✅ Otimização concluída com sucesso! {len(cds_selecionados)} CDs selecionados cobrindo {percentual:.1f}% da demanda total ({demanda_coberta:.0f} unidades)."
                flash(mensagem, 'success')
                print(f"✅ Mensagem de sucesso configurada: {mensagem}")
                print(f"✅ Redirecionando para template resultado_maxcobertura.html")
                
            return render_template('resultado_maxcobertura.html', resultado=resultado)
        else:
            print("❌ Arquivo inválido ou não recebido")
            flash('Arquivo inválido ou não selecionado.', 'error')
            return redirect(url_for('max_cobertura'))
    except Exception as e:
        print(f"❌ Erro na rota: {str(e)}")
        flash(f'Erro ao processar: {str(e)}', 'error')
        return redirect(url_for('max_cobertura'))

# --- ROTA PARA p-CENTROS ---
@app.route('/p_centros')
def p_centros():
    return render_template('p_centros.html')

@app.route('/gerar_template_pcentros', methods=['POST'])
def gerar_template_pcentros():
    try:
        num_cds = int(request.form.get('num_cds', 0))
        num_clientes = int(request.form.get('num_clientes', 0))
        
        if num_cds <= 0 or num_clientes <= 0:
            flash('Por favor, insira valores válidos para gerar o template.', 'error')
            return redirect(url_for('p_centros'))
            
        # 1. Montando os dados da Aba 1 (P-Centros)
        colunas = ['Origem / Destino'] + [f'Cliente {i}' for i in range(1, num_clientes + 1)]
        df_template = pd.DataFrame(columns=colunas)
        
        for i in range(1, num_cds + 1):
            df_template.loc[i-1] = [f'CD {i}'] + ['' for _ in range(num_clientes)]
            
        # 2. Montando os dados da Aba 2 (Como Usar)
        guia_uso = [
            ["📋 GUIA DE USO - P-CENTROS", ""],
            ["", ""],
            ["🎯 OBJETIVO", "Abrir a quantidade desejada de CDs para minimizar a MAIOR distância de atendimento."],
            ["", ""],
            ["📝 COMO PREENCHER A PLANILHA \"P-Centros\":", ""],
            ["1. DISTÂNCIAS/TEMPOS (Células Cinzas)", "Informe a distância (km) ou tempo de cada CD para cada cliente."],
            ["2. MATRIZ SIMPLIFICADA", "No p-Centros não usamos Demanda, Custo Fixo ou Capacidade. Preencha apenas rotas."],
            ["", ""],
            ["🎨 GUIA VISUAL DAS CORES:", ""],
            ["AZUL ESCURO (Cabeçalho)", "Títulos das colunas - não editar."],
            ["AZUL CLARO (Primeira coluna)", "Nomes dos CDs candidatos - não editar."],
            ["CINZA CLARO (Miolo da tabela)", "ÁREA PARA PREENCHER SEUS DADOS DE DISTÂNCIA."],
            ["", ""],
            ["⚠️ DICAS IMPORTANTES:", ""],
            ["• Formato Numérico", "Use números sem formatação de texto (ex: 5.50, não escreva 'R$ 5,50' ou 'km')."],
            ["• Foco no Pior Cenário", "O sistema escolherá os CDs garantindo que nenhum cliente fique longe demais."]
        ]
        df_como_usar = pd.DataFrame(guia_uso)
            
        # 3. Criando o arquivo Excel com as duas abas
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_template.to_excel(writer, index=False, sheet_name='P-Centros')
            df_como_usar.to_excel(writer, index=False, header=False, sheet_name='Como Usar')
            
            workbook = writer.book
            
            # --- ESTILIZANDO A ABA 1 (P-Centros) ---
            worksheet = writer.sheets['P-Centros']
            header_fill = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid') # Azul escuro
            cd_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')     # Azul claro
            input_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')  # Cinza
            
            white_font = Font(color='FFFFFF', bold=True)
            bold_font = Font(bold=True)
            center_alignment = Alignment(horizontal='center', vertical='center')
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            for col_num, col_name in enumerate(colunas, 1):
                cell = worksheet.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = white_font
                cell.alignment = center_alignment
                cell.border = thin_border
                worksheet.column_dimensions[cell.column_letter].width = 15
                
            for row_num in range(2, num_cds + 2):
                cd_cell = worksheet.cell(row=row_num, column=1)
                cd_cell.fill = cd_fill
                cd_cell.font = bold_font
                cd_cell.alignment = center_alignment
                cd_cell.border = thin_border
                
                for col_num in range(2, num_clientes + 2):
                    input_cell = worksheet.cell(row=row_num, column=col_num)
                    input_cell.fill = input_fill
                    input_cell.alignment = center_alignment
                    input_cell.border = thin_border

            # --- ESTILIZANDO A ABA 2 (Como Usar) ---
            ws_como_usar = writer.sheets['Como Usar']
            ws_como_usar.column_dimensions['A'].width = 45
            ws_como_usar.column_dimensions['B'].width = 90
            
            for row in range(1, len(guia_uso) + 1):
                cell_a = ws_como_usar.cell(row=row, column=1)
                cell_b = ws_como_usar.cell(row=row, column=2)
                
                # Colocar negrito na primeira coluna
                if cell_a.value:
                    cell_a.font = bold_font
                
                # Destacar os cabeçalhos principais
                if row in [1, 3, 5, 9, 14]: 
                    cell_a.font = Font(bold=True, size=12, color='1F4E78')
                    
        output.seek(0)
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='Template_P-Centros.xlsx'
        )
        
    except Exception as e:
        flash(f'Erro ao gerar template: {str(e)}', 'error')
        return redirect(url_for('p_centros'))

@app.route('/resolver_pcentros', methods=['POST'])
def resolver_pcentros_route():
    try:
        if 'file' not in request.files:
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('p_centros'))
        
        file = request.files['file']
        
        if file.filename == '':
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('p_centros'))
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Ler dados do Excel
            df_principal = pd.read_excel(filepath)
            
            # Obter valor de p do formulário
            p_cds = int(request.form.get('p_cds', 1))
            
            # Resolver o problema p-Centros
            from solver_pcentros import resolver_pcentros
            resultado = resolver_pcentros(df_principal, p_cds)
            
            # Limpar arquivo temporário
            os.remove(filepath)
            
            if resultado.get('status') == 'Erro':
                flash(f"Erro na otimização: {resultado.get('mensagem')}", 'error')
                return redirect(url_for('p_centros'))
            
            return render_template('resultado_pcentros.html', resultado=resultado)
                
        else:
            flash('Formato de arquivo inválido. Por favor, envie um arquivo .xlsx ou .csv.', 'error')
            return redirect(url_for('p_centros'))
            
    except Exception as e:
        flash(f'Erro ao processar arquivo: {str(e)}', 'error')
        return redirect(url_for('p_centros'))

# --- ROTA PARA p-MEDIANAS ---
@app.route('/gerar_template_pmedianas', methods=['POST'])
def gerar_template_pmedianas():
    try:
        # Pega os dados do formulário
        num_cds = int(request.form.get('num_cds', 0))
        num_clientes = int(request.form.get('num_clientes', 0))
        p_cds = int(request.form.get('p_cds', 0))
        
        if num_cds <= 0 or num_clientes <= 0 or p_cds <= 0:
            flash('Todos os valores devem ser maiores que zero.', 'error')
            return redirect(url_for('p_medianas'))
        
        if p_cds > num_cds:
            flash('O número de CDs a abrir (p) não pode ser maior que o número de CDs candidatos.', 'error')
            return redirect(url_for('p_medianas'))

        # Criar as listas de nomes
        nomes_cds = [f'CD {i}' for i in range(1, num_cds + 1)]
        nomes_clientes = [f'Cliente {i}' for i in range(1, num_clientes + 1)]
        
        # Cabeçalho completo (com custo fixo e capacidade para manter consistência)
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
        
        # Criar o DataFrame
        df = pd.DataFrame(dados, columns=colunas)
        
        # Salvar na memória como Excel com Design Profissional
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='P-Medianas', index=False)
            worksheet = writer.sheets['P-Medianas']
            
            # Criar aba de instruções
            instructions_worksheet = writer.book.create_sheet('Como Usar')
            
            # --- DEFININDO OS ESTILOS PARA ABA DE INSTRUÇÕES ---
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True, size=12)
            highlight_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            bold_font = Font(bold=True, size=11)
            normal_font = Font(size=10)
            center_align = Alignment(horizontal="center", vertical="center")
            left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            # --- ADICIONAR INSTRUÇÕES NA ABA "Como Usar" ---
            instructions_data = [
                ['📋 GUIA DE USO - P-MEDIANAS', ''],
                ['', ''],
                ['🎯 OBJETIVO', f'Abrir exatamente {p_cds} CDs para minimizar a distância total ponderada pela demanda'],
                ['', ''],
                ['📝 COMO PREENCHER A PLANILHA "P-Medianas":', ''],
                ['1. CUSTOS DE TRANSPORTE (Células Cinzas)', 'Nas células cinza, informe o custo unitário (R$/unidade) ou distância (km) de cada CD para cada cliente'],
                ['2. DEMANDA DOS CLIENTES (Linha "Demanda Total")', 'Informe a demanda mensal (unidades) de cada cliente'],
                ['3. CUSTO FIXO E CAPACIDADE (Colunas finais)', 'Essas colunas existem para manter formato, mas NÃO são usadas no modelo p-Medianas'],
                ['', ''],
                ['🎨 GUIA VISUAL DAS CORES:', ''],
                ['AZUL ESCURO (Cabeçalho)', 'Títulos das colunas - não editar'],
                ['AZUL CLARO (Primeira coluna)', 'Nomes dos CDs e linha de demanda - não editar'],
                ['CINZA CLARO (Miolo da tabela)', 'ÁREA PARA PREENCHER SEUS DADOS'],
                ['', ''],
                ['💡 DIFERENÇAS PARA O MODELO ANTERIOR:', ''],
                ['• Sem Custo Fixo', 'Coluna existe mas não é usada no cálculo'],
                ['• Sem Capacidade', 'Coluna existe mas os CDs têm capacidade infinita'],
                ['• Número Fixo de CDs', f'Exatamente {p_cds} CDs serão abertos'],
                ['• Atendimento Exclusivo', 'Cada cliente é atendido por apenas 1 CD'],
                ['', ''],
                ['⚠️ DICAS IMPORTANTES:', ''],
                ['• Use números sem formatação (ex: 5.50, não R$ 5,50)', ''],
                ['• O sistema atribuirá cada cliente ao CD mais próximo', ''],
                ['• Demanda pondera a importância de cada cliente', ''],
                ['', ''],
                ['🚀 PRÓXIMOS PASSOS:', ''],
                ['1. Preencha todos os dados na aba "P-Medianas"', ''],
                ['2. Salve o arquivo Excel', ''],
                ['3. Faça upload no sistema', ''],
                ['4. Aguarde o resultado da otimização', ''],
            ]
            
            # Adicionar instruções na worksheet
            for row_idx, (col1, col2) in enumerate(instructions_data, 1):
                cell1 = instructions_worksheet.cell(row=row_idx, column=1, value=col1)
                if row_idx == 1:
                    cell1.font = header_font
                    cell1.fill = header_fill
                    cell1.alignment = center_align
                elif col1.startswith('🎯') or col1.startswith('📝') or col1.startswith('🎨') or col1.startswith('💡') or col1.startswith('⚠️') or col1.startswith('🚀'):
                    cell1.font = bold_font
                    cell1.fill = highlight_fill
                    cell1.alignment = left_align
                else:
                    cell1.font = normal_font
                    cell1.alignment = left_align
                cell1.border = thin_border
                
                cell2 = instructions_worksheet.cell(row=row_idx, column=2, value=col2)
                cell2.font = normal_font
                cell2.alignment = left_align
                cell2.border = thin_border
            
            # Ajustar largura das colunas da aba de instruções
            instructions_worksheet.column_dimensions['A'].width = 50
            instructions_worksheet.column_dimensions['B'].width = 80
            instructions_worksheet.freeze_panes = 'A2'
            
            # --- DEFININDO OS ESTILOS PARA ABA DE DADOS ---
            header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            col1_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
            bold_font = Font(bold=True)
            input_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
            center_align = Alignment(horizontal="center", vertical="center")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                 top=Side(style='thin'), bottom=Side(style='thin'))
            
            # --- APLICANDO OS ESTILOS NA ABA DE DADOS ---
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
                    
                    if col == 1:
                        cell.font = bold_font
                        cell.fill = col1_fill
                    else:
                        cell.fill = input_fill

            # 3. Ajustar largura das colunas automaticamente
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[column].width = max_length + 4

        output.seek(0)
        
        return send_file(
            output,
            download_name="Template_P-Medianas.xlsx",
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        flash(f'Erro ao gerar template: {str(e)}', 'error')
        return redirect(url_for('p_medianas'))

@app.route('/resolver_pmedianas', methods=['POST'])
def resolver_pmedianas():
    try:
        if 'arquivo_planilha' not in request.files:
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('p_medianas'))
        
        file = request.files['arquivo_planilha']
        
        if file.filename == '':
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('p_medianas'))
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Ler dados do Excel
            df_principal = pd.read_excel(filepath)
            
            # Extrair p do formulário (número de CDs a abrir)
            p_cds = int(request.form.get('p_cds', 0))
            
            # Resolver o problema p-medianas
            from solver_pmedianas import resolver_pmedianas
            resultado = resolver_pmedianas(df_principal, p_cds)
            
            # Limpar arquivo temporário
            os.remove(filepath)
            
            if resultado.get('status') == 'Erro':
                flash(f"Erro na otimização: {resultado.get('mensagem')}", 'error')
                return redirect(url_for('p_medianas'))
            
            return render_template('resultado_pmedianas.html', resultado=resultado)
                
        else:
            flash('Formato de arquivo inválido. Por favor, envie um arquivo .xlsx ou .csv.', 'error')
            return redirect(url_for('p_medianas'))
            
    except Exception as e:
        flash(f'Erro ao processar arquivo: {str(e)}', 'error')
        return redirect(url_for('p_medianas'))

@app.route('/gerar_template_pmediana_simples', methods=['POST'])
def gerar_template_pmediana_simples():
    try:
        # Pega os dados do formulário
        num_cds = int(request.form.get('num_cds', 0))
        num_clientes = int(request.form.get('num_clientes', 0))
        p_cds = int(request.form.get('p_cds', 1))
        
        if num_cds <= 0 or num_clientes <= 0:
            flash('Número de CDs e clientes deve ser maior que zero.', 'error')
            return redirect(url_for('pmediana_simples'))
        
        # Criar DataFrame para a planilha (apenas coordenadas)
        dados = []
        for i in range(num_cds):
            dados.append({
                'ID': f'P{i+1}',
                'Coordenada X': 0.0,
                'Coordenada Y': 0.0
            })
        
        df = pd.DataFrame(dados)
        
        # Criar arquivo Excel em memória
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Dados', index=False)
            
            # Adicionar aba de instruções
            instrucoes = pd.DataFrame([
                ['Como Usar', ''],
                ['', ''],
                ['1. Preencha os dados:', ''],
                ['- ID: Identificador único para cada ponto', ''],
                ['- Coordenada X: Posição X no plano cartesiano', ''],
                ['- Coordenada Y: Posição Y no plano cartesiano', ''],
                ['', ''],
                ['2. Salve o arquivo:', ''],
                ['- Clique em Arquivo > Salvar Como', ''],
                ['- Escolha formato .xlsx', ''],
                ['', ''],
                ['3. Faça upload:', ''],
                ['- Volte para a página do sistema', ''],
                ['- Clique em "Escolher arquivo" e selecione a planilha', ''],
                ['- Clique em "Calcular p-Mediana"', ''],
                ['', ''],
                ['Observações:', ''],
                ['- Mínimo de 1 ponto necessário', ''],
                ['- As coordenadas podem ser decimais', ''],
                ['- O sistema encontrará o ponto ótimo automaticamente', ''],
                ['', ''],
                ['Parâmetros configurados:', ''],
                [f'- Pontos de análise: {num_cds}', ''],
                [f'- Clientes (referência): {num_clientes}', ''],
                [f'- p-Medianas: {p_cds} (será usado para cálculo)', '']
            ])
            instrucoes.to_excel(writer, sheet_name='Como Usar', index=False)
        
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=f'template_pmediana_simples_{num_cds}_pontos.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        flash(f'Erro ao gerar template: {str(e)}', 'error')
        return redirect(url_for('pmediana_simples'))

# --- ROTA PARA p-MEDIANA SIMPLES ---
@app.route('/resolver_pmediana_simples', methods=['POST'])
def resolver_pmediana_simples_route():
    try:
        if 'arquivo_planilha' not in request.files:
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('pmediana_simples'))
        
        file = request.files['arquivo_planilha']
        
        if file.filename == '':
            flash('Nenhum arquivo foi selecionado.', 'error')
            return redirect(url_for('pmediana_simples'))
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            # Ler dados do Excel
            df_principal = pd.read_excel(filepath)
            
            # Resolver o problema p-Mediana simples
            from solver_pmediana_simples import resolver_pmediana_simples
            resultado = resolver_pmediana_simples(df_principal)
            
            # Limpar arquivo temporário
            os.remove(filepath)
            
            if resultado.get('status') == 'Erro':
                flash(f"Erro na otimização: {resultado.get('mensagem')}", 'error')
                return redirect(url_for('pmediana_simples'))
            
            return render_template('resultado_pmediana_simples.html', resultado=resultado)
                
        else:
            flash('Formato de arquivo inválido. Por favor, envie um arquivo .xlsx ou .csv.', 'error')
            return redirect(url_for('pmediana_simples'))
            
    except Exception as e:
        flash(f'Erro ao processar arquivo: {str(e)}', 'error')
        return redirect(url_for('pmediana_simples'))

if __name__ == '__main__':
    app.run(debug=True)
