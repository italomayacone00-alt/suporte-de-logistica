from pulp import *
import pandas as pd

def resolver_problema_logistica(df):
    """
    Resolve o problema de localização de CDs baseado em uma matriz única (estilo Solver Excel).
    A última linha deve ser a "Demanda Total".
    """
    try:
        print("=== DEBUG: Iniciando processamento ===")
        print(f"Shape do DataFrame: {df.shape}")
        print(f"Colunas: {list(df.columns)}")
        print(f"Primeiras linhas:\n{df.head()}")
        
        # 1. Separar a linha de Demanda dos CDs candidatos
        # O sistema procura a linha onde a primeira coluna diz "Demanda Total"
        nome_primeira_coluna = df.columns[0] 
        print(f"Nome da primeira coluna: {nome_primeira_coluna}")
        
        linha_demanda = df[df[nome_primeira_coluna] == 'Demanda Total']
        print(f"Linha de demanda encontrada:\n{linha_demanda}")
        
        df_cds = df[df[nome_primeira_coluna] != 'Demanda Total'].dropna(subset=[nome_primeira_coluna])
        print(f"DataFrame de CDs:\n{df_cds}")
        
        # 2. Identificar listas de CDs e Clientes
        cds = df_cds[nome_primeira_coluna].astype(str).tolist()
        print(f"CDs encontrados: {cds}")
        
        # Clientes são todas as colunas entre a primeira e as duas últimas (Custo Fixo e Capacidade)
        clientes = df.columns[1:-2].tolist()
        print(f"Clientes encontrados: {clientes}")
        
        # 3. Extrair os Dicionários de Parâmetros
        capacidade = dict(zip(cds, pd.to_numeric(df_cds['Capacidade'])))
        print(f"Capacidade: {capacidade}")
        
        custo_fixo = dict(zip(cds, pd.to_numeric(df_cds['Custo Fixo'])))
        print(f"Custo Fixo: {custo_fixo}")
        
        # Extrair a demanda da última linha
        demanda = {cliente: pd.to_numeric(linha_demanda[cliente].values[0]) for cliente in clientes}
        print(f"Demanda: {demanda}")
        
        # Extrair Custos de Transporte (Matriz Cij)
        custos_transporte = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            for cliente in clientes:
                custos_transporte[(cd_atual, cliente)] = pd.to_numeric(row[cliente])
        print(f"Custos de transporte (primeiros 5): {dict(list(custos_transporte.items())[:5])}")

        # ==========================================
        # 4. MODELAGEM MATEMÁTICA (PuLP)
        # ==========================================
        prob = LpProblem("Otimizacao_Malha_Logistica", LpMinimize)
        
        # Variáveis de Decisão
        # Y_i: 1 se o CD for aberto, 0 caso contrário
        y = LpVariable.dicts("Abrir_CD", cds, cat='Binary')
        
        # X_ij: Quantidade enviada do CD i para o Cliente j
        x = LpVariable.dicts("Envio", (cds, clientes), lowBound=0, cat='Continuous')
        
        # Função Objetivo: Minimizar (Custo Fixo + Custo de Transporte)
        prob += lpSum([custo_fixo[i] * y[i] for i in cds]) + \
                lpSum([custos_transporte[(i, j)] * x[i][j] for i in cds for j in clientes]), "Custo_Total"
                
        # Restrição 1: Toda a demanda dos clientes DEVE ser atendida (=)
        for j in clientes:
            prob += lpSum([x[i][j] for i in cds]) == demanda[j], f"Demanda_{j}"
            
        # Restrição 2: O envio de um CD não pode ultrapassar sua capacidade (<=)
        # E se o CD estiver fechado (y[i] = 0), a capacidade é 0.
        for i in cds:
            prob += lpSum([x[i][j] for j in clientes]) <= capacidade[i] * y[i], f"Capacidade_{i}"
            
        # 5. Resolver o problema
        print("=== DEBUG: Iniciando otimização ===")
        prob.solve()
        print(f"Status da solução: {LpStatus[prob.status]}")
        
        # 6. Formatar o Resultado para enviar de volta ao Front-end
        status = LpStatus[prob.status]
        
        if status != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução viável com os dados fornecidos. Verifique se a capacidade total é maior que a demanda."}
            
        cds_abertos = []
        fluxo_envios = []
        
        for i in cds:
            if value(y[i]) == 1:
                cds_abertos.append(i)
                print(f"CD {i} foi aberto")
                for j in clientes:
                    qtd = value(x[i][j])
                    if qtd > 0:
                        fluxo_envios.append({
                            "origem": i,
                            "destino": j,
                            "quantidade": qtd,
                            "custo_unitario": custos_transporte[(i, j)],
                            "custo_total_rota": qtd * custos_transporte[(i, j)]
                        })
                        print(f"  Rota: {i} -> {j}: {qtd} unidades")
        
        print(f"Total de CDs abertos: {len(cds_abertos)}")
        print(f"Total de rotas: {len(fluxo_envios)}")
                        
        resultado_final = {
            "status": "Sucesso",
            "custo_total": value(prob.objective),
            "cds_abertos": cds_abertos,
            "num_cds_abertos": len(cds_abertos),
            "fluxo": fluxo_envios,
            "transportes": fluxo_envios,
            "capacidade_utilizada": {}
        }
        
        # Adicionar informações de capacidade utilizada
        for i in cds:
            if i in cds_abertos:
                capacidade_total = capacidade[i]
                capacidade_usada = sum([envio["quantidade"] for envio in fluxo_envios if envio["origem"] == i])
                resultado_final["capacidade_utilizada"][i] = {
                    "disponivel": capacidade_total,
                    "utilizada": capacidade_usada,
                    "percentual_uso": round((capacidade_usada / capacidade_total) * 100, 2) if capacidade_total > 0 else 0
                }
                print(f"Capacidade {i}: {capacidade_usada}/{capacidade_total} ({resultado_final['capacidade_utilizada'][i]['percentual_uso']}%)")
        
        print("=== DEBUG: Resultado final ===")
        print(f"Status: {resultado_final['status']}")
        print(f"Custo total: {resultado_final['custo_total']}")
        print(f"CDs abertos: {resultado_final['cds_abertos']}")
        print(f"Transportes: {len(resultado_final['transportes'])}")
        print(f"Capacidade utilizada: {resultado_final['capacidade_utilizada']}")
        
        return resultado_final

    except Exception as e:
        print(f"=== ERRO: {str(e)} ===")
        import traceback
        traceback.print_exc()
        return {"status": "Erro", "mensagem": f"Erro ao processar os dados da planilha: {str(e)}"}
