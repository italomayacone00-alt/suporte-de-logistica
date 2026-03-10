"""
Solver para o Problema dos p-Centros
Objetivo: Minimizar a MAIOR distância entre qualquer cliente e o CD que o atende.
"""

import pandas as pd
from pulp import LpProblem, LpMinimize, LpVariable, lpSum, LpStatus, value

def resolver_pcentros(df, p, tipo_dado='distancia'):
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO p-CENTROS (p={p}, tipo={tipo_dado}) ===")
        print(f"Shape do DataFrame: {df.shape}")
        print(f"Colunas: {list(df.columns)}")
        print(f"Primeiras linhas:\n{df.head()}")
        
        # 1. TRATAMENTO DOS DADOS (Planilha limpa)
        cds = df.iloc[:, 0].tolist() # Pega todas as linhas da coluna 0
        clientes = df.columns[1:].tolist() # Pega todas as colunas a partir da 1
        
        print(f"CDs encontrados: {cds}")
        print(f"Clientes encontrados: {clientes}")
        
        # Vamos pegar apenas os valores (distâncias ou custos)
        valores = {}
        for idx, cd in enumerate(cds):
            valores[cd] = {}
            for cliente in clientes:
                valores[cd][cliente] = float(df.iloc[idx][cliente])
        
        print(f"Matriz de {tipo_dado}s criada com {len(valores)} CDs e {len(clientes)} clientes")

        # 2. MODELO
        prob = LpProblem("Problema_p_Centros", LpMinimize)
        
        # 3. VARIÁVEIS DE DECISÃO
        Y = LpVariable.dicts("Y", cds, cat='Binary') # 1 se CD aberto, 0 se fechado
        X = LpVariable.dicts("X", [(i, j) for i in cds for j in clientes], cat='Binary') # Atribuição
        
        # O pulo do gato do p-Centros: Z é a MAIOR distância de todas
        Z = LpVariable("Z", lowBound=0, cat='Continuous')
        
        # 4. FUNÇÃO OBJETIVO: Queremos apenas minimizar o Z (o pior valor)
        prob += Z, f"Minimizar_Maior_{tipo_dado}"
        
        # 5. RESTRIÇÕES
        # R1: Abrir exatamente 'p' CDs
        prob += lpSum(Y[i] for i in cds) == p, "Qtd_CDs"
        
        # R2: Cada cliente é atendido por 1 único CD
        for j in clientes:
            prob += lpSum(X[i, j] for i in cds) == 1, f"Atendimento_{j}"
            
        # R3: Um CD só pode atender se estiver aberto
        for i in cds:
            for j in clientes:
                prob += X[i, j] <= Y[i], f"Logica_Aberto_{i}_{j}"
                
        # R4: A MÁGICA - Z tem que ser maior ou igual a qualquer valor de cliente atendido
        for i in cds:
            for j in clientes:
                prob += Z >= valores[i][j] * X[i, j], f"Max_{tipo_dado}_{i}_{j}"
                
        # 6. RESOLUÇÃO
        prob.solve()
        
        if LpStatus[prob.status] != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução."}
            
        # 7. EXTRAÇÃO DE RESULTADOS
        cds_selecionados = [i for i in cds if value(Y[i]) > 0.5]
        valor_maximo = value(Z) # O nosso pior cenário otimizado
        
        atribuicoes = []
        total_clientes_por_cd = {cd: 0 for cd in cds_selecionados}
        
        for i in cds_selecionados:
            for j in clientes:
                if value(X[i, j]) > 0.5:
                    atribuicoes.append({
                        "cd": i,
                        "cliente": j,
                        "valor": valores[i][j],
                        "tipo_valor": tipo_dado
                    })
                    total_clientes_por_cd[i] += 1
                    
        return {
            "status": "Sucesso",
            "modelo": f"p-Centros (Min-Max) - {tipo_dado}s",
            "p": p,
            "valor_maximo": valor_maximo,
            "tipo_valor": tipo_dado,
            "cds_selecionados": cds_selecionados,
            "num_cds_selecionados": len(cds_selecionados),
            "atribuicoes": atribuicoes,
            "clientes_por_cd": total_clientes_por_cd
        }

    except Exception as e:
        return {"status": "Erro", "mensagem": f"Erro interno: {str(e)}"}
