"""
Solver para o Problema de Localização Capacitada (Modelo Original)
Otimização de CDs com custos fixos e capacidades
"""

import pandas as pd
from pulp import LpProblem, LpMinimize, LpVariable, lpSum, LpStatus, value

def resolver_problema_logistica(df):
    """
    Resolve o problema de localização-alocação capacitada.
    
    Modelo:
    - Minimiza custos fixos + custos de transporte
    - Respeita capacidades dos CDs
    - Determina número ótimo de CDs a abrir
    
    Parâmetros:
    - df: DataFrame com dados da planilha
    
    Retorna:
    - Dicionário com solução ótima
    """
    try:
        print("=== DEBUG: Iniciando solver tradicional ===")
        print(f"Shape do DataFrame: {df.shape}")
        print(f"Colunas: {list(df.columns)}")
        print(f"Primeiras linhas:\n{df.head()}")
        
        # 1. Separar a linha de Demanda dos CDs candidatos
        nome_primeira_coluna = df.columns[0] 
        print(f"Nome da primeira coluna: {nome_primeira_coluna}")
        
        linha_demanda = df[df[nome_primeira_coluna] == 'Demanda Total']
        print(f"Linha de demanda encontrada:\n{linha_demanda}")
        
        df_cds = df[df[nome_primeira_coluna] != 'Demanda Total'].dropna(subset=[nome_primeira_coluna])
        print(f"DataFrame de CDs:\n{df_cds}")
        
        # 2. Identificar listas de CDs e Clientes
        cds = df_cds[nome_primeira_coluna].astype(str).tolist()
        print(f"CDs encontrados: {cds}")
        
        clientes = df.columns[1:-2].tolist()  # Excluindo Custo Fixo e Capacidade
        print(f"Clientes encontrados: {clientes}")
        
        # 3. Extrair os Dicionários de Parâmetros
        demanda = {cliente: pd.to_numeric(linha_demanda[cliente].values[0]) for cliente in clientes}
        print(f"Demanda: {demanda}")
        
        # Extrair Custos Fixos
        custos_fixos = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            custos_fixos[cd_atual] = pd.to_numeric(row['Custo Fixo'])
        print(f"Custos fixos: {custos_fixos}")
        
        # Extrair Capacidades
        capacidades = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            capacidades[cd_atual] = pd.to_numeric(row['Capacidade'])
        print(f"Capacidades: {capacidades}")
        
        # Extrair Custos de Transporte
        custos_transporte = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            for cliente in clientes:
                custos_transporte[(cd_atual, cliente)] = pd.to_numeric(row[cliente])
        print(f"Custos de transporte (primeiros 5): {dict(list(custos_transporte.items())[:5])}")
        
        # ==========================================
        # MODELAGEM MATEMÁTICA - LOCALIZAÇÃO CAPACITADA
        # ==========================================
        prob = LpProblem("Localizacao_Capacitada", LpMinimize)
        
        # Variáveis de Decisão
        # Y_i: 1 se o CD for aberto, 0 caso contrário
        y = LpVariable.dicts("Abrir_CD", cds, cat='Binary')
        
        # X_ij: quantidade enviada do CD i para o cliente j
        x = LpVariable.dicts("Envio", (cds, clientes), lowBound=0)
        
        # Função Objetivo: Minimizar custo total (fixo + transporte)
        custo_fixo_total = lpSum([custos_fixos[i] * y[i] for i in cds])
        custo_transporte_total = lpSum([custos_transporte[(i, j)] * x[i][j] for i in cds for j in clientes])
        
        prob += custo_fixo_total + custo_transporte_total, "Custo_Total"
        
        # Restrição 1: Cada cliente deve ter sua demanda atendida
        for j in clientes:
            prob += lpSum([x[i][j] for i in cds]) == demanda[j], f"Atender_Demanda_{j}"
        
        # Restrição 2: Não enviar mais que a capacidade de cada CD
        for i in cds:
            prob += lpSum([x[i][j] for j in clientes]) <= capacidades[i] * y[i], f"Capacidade_{i}"
        
        # 4. Resolver o problema
        print("=== DEBUG: Iniciando otimização tradicional ===")
        prob.solve()
        print(f"Status da solução: {LpStatus[prob.status]}")
        
        # 5. Formatar o Resultado
        status = LpStatus[prob.status]
        
        if status != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução viável."}
            
        cds_abertos = []
        transportes = []
        capacidade_utilizada = {}
        
        # Identificar CDs abertos e calcular capacidade utilizada
        for i in cds:
            if value(y[i]) == 1:
                cds_abertos.append(i)
                capacidade_total = capacidades[i]
                capacidade_usada = sum([value(x[i][j]) for j in clientes])
                percentual_uso = (capacidade_usada / capacidade_total) * 100 if capacidade_total > 0 else 0
                
                capacidade_utilizada[i] = {
                    'disponivel': capacidade_total,
                    'utilizada': capacidade_usada,
                    'percentual_uso': percentual_uso
                }
        
        # Identificar transportes
        for i in cds:
            for j in clientes:
                quantidade = value(x[i][j])
                if quantidade > 0:
                    transportes.append({
                        'origem': i,
                        'destino': j,
                        'quantidade': quantidade,
                        'custo_unitario': custos_transporte[(i, j)],
                        'custo_total_rota': quantidade * custos_transporte[(i, j)]
                    })
        
        resultado_final = {
            "status": "Sucesso",
            "modelo": "Localização Capacitada",
            "custo_total": value(prob.objective),
            "custo_fixo_total": value(custo_fixo_total),
            "custo_transporte_total": value(custo_transporte_total),
            "cds_abertos": cds_abertos,
            "num_cds_abertos": len(cds_abertos),
            "transportes": transportes,
            "num_transportes": len(transportes),
            "capacidade_utilizada": capacidade_utilizada,
            "demanda_total": sum(demanda.values()),
            "demanda_atendida": sum([t['quantidade'] for t in transportes])
        }
        
        print("=== DEBUG: Resultado final tradicional ===")
        print(f"Status: {resultado_final['status']}")
        print(f"Custo total: {resultado_final['custo_total']}")
        print(f"CDs abertos: {resultado_final['cds_abertos']}")
        print(f"Transportes: {len(resultado_final['transportes'])}")
        
        return resultado_final

    except Exception as e:
        print(f"=== ERRO: {str(e)} ===")
        import traceback
        traceback.print_exc()
        return {"status": "Erro", "mensagem": f"Erro ao processar os dados da planilha: {str(e)}"}
