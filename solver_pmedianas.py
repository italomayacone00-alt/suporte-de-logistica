"""
Solver: Customizado para espelhar a planilha Excel (Valor Alvo: 9581)
Regras: Soma Custo Fixo + Distâncias Simples (Não Ponderadas), Respeitando a Capacidade e p=2.
"""

import pandas as pd
from pulp import LpProblem, LpMinimize, LpVariable, lpSum, LpStatus, value

def resolver_pmedianas(df, p, tipo_dado='distancia'):
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO p-MEDIANAS (p={p}, tipo={tipo_dado}) ===")
        
        # 1. TRATAMENTO DOS DADOS
        cds = df.iloc[:-1, 0].tolist()
        clientes = df.columns[1:-2].tolist()
        
        demanda = {cliente: float(df.iloc[-1][cliente]) for cliente in clientes}
        
        # Pegando as colunas extras que o Excel usa no cálculo
        custos_fixos = {}
        capacidades = {}
        for idx, cd in enumerate(cds):
            custos_fixos[cd] = float(df.iloc[idx]['Custo Fixo'])
            capacidades[cd] = float(df.iloc[idx]['Capacidade'])
            
        valores = {}
        for idx, cd in enumerate(cds):
            valores[cd] = {}
            for cliente in clientes:
                valores[cd][cliente] = float(df.iloc[idx][cliente])

        # 2. MODELO
        prob = LpProblem("Modelo_Excel_Original", LpMinimize)
        
        # 3. VARIÁVEIS DE DECISÃO
        Y = LpVariable.dicts("Y", cds, cat='Binary') 
        X = LpVariable.dicts("X", [(i, j) for i in cds for j in clientes], cat='Binary')
        
        # 4. FUNÇÃO OBJETIVO: Valores + Custo Fixo 
        prob += lpSum(valores[i][j] * X[i, j] for i in cds for j in clientes) + \
                lpSum(custos_fixos[i] * Y[i] for i in cds)
        
        # 5. RESTRIÇÕES
        # R1: Abrir exatamente 'p' CDs
        prob += lpSum(Y[i] for i in cds) == p, "Qtd_CDs"
        
        # R2: Cada cliente é atendido por 1 único CD
        for j in clientes:
            prob += lpSum(X[i, j] for i in cds) == 1, f"Atendimento_{j}"
            
        # R3: Um CD só pode atender clientes se estiver aberto
        for i in cds:
            for j in clientes:
                prob += X[i, j] <= Y[i], f"Logica_Aberto_{i}_{j}"
                
        # R4: RESTRIÇÃO DE CAPACIDADE (A que fez o valor ir de 9094 para 9581)
        for i in cds:
            prob += lpSum(demanda[j] * X[i, j] for j in clientes) <= capacidades[i] * Y[i], f"Capacidade_{i}"
                
        # 6. RESOLUÇÃO
        prob.solve()
        
        if LpStatus[prob.status] != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução."}
            
        # 7. EXTRAÇÃO DE RESULTADOS
        cds_selecionados = [i for i in cds if value(Y[i]) > 0.5]
        custo_total = value(prob.objective)
        
        atribuicoes = []
        total_clientes_por_cd = {cd: 0 for cd in cds_selecionados}
        
        for i in cds_selecionados:
            for j in clientes:
                if value(X[i, j]) > 0.5:
                    atribuicoes.append({
                        "cd": i,
                        "cliente": j,
                        "valor": valores[i][j],
                        "demanda_atendida": demanda[j],
                        "tipo_valor": tipo_dado,
                        "custo_ponderado": valores[i][j] # Mantendo compatibilidade
                    })
                    total_clientes_por_cd[i] += 1
                    
        return {
            "status": "Sucesso",
            "modelo": f"Capacitado com Custo Fixo - {tipo_dado}s",
            "p": p,
            "custo_total_ponderado": custo_total, # Mantendo compatibilidade
            "tipo_valor": tipo_dado,
            "cds_selecionados": cds_selecionados,
            "num_cds_selecionados": len(cds_selecionados),
            "atribuicoes": atribuicoes,
            "clientes_por_cd": total_clientes_por_cd,
            "demanda_total": sum(demanda.values())
        }

    except Exception as e:
        return {"status": "Erro", "mensagem": f"Erro interno: {str(e)}"}