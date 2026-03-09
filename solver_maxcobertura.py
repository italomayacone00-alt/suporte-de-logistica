"""
Solver para Máxima Cobertura (MCLP) - Padrão Logístico (CDs e Clientes)
Baseado na abordagem simples do p-Centros
"""
import pandas as pd
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpStatus, value

def resolver_maxcobertura(df, p, raio_cobertura=0.0):
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO MÁXIMA COBERTURA (p={p}, raio={raio_cobertura}) ===")
        print(f"Shape do DataFrame: {df.shape}")
        print(f"Colunas: {list(df.columns)}")
        print(f"Primeiras linhas:\n{df.head()}")
        
        # 1. TRATAMENTO DOS DADOS (mesma abordagem do p-Centros)
        ultima_linha_nome = str(df.iloc[-1, 0]).strip().lower()
        
        if ultima_linha_nome in ['demanda', 'demanda total', 'peso']:
            print("✅ Linha de demanda encontrada")
            demanda = {cliente: float(df.iloc[-1][cliente]) for cliente in df.columns[1:]}
            cds = df.iloc[:-1, 0].tolist()
            matriz_valores = df.iloc[:-1]
        else:
            print("❌ Linha de demanda NÃO encontrada, usando peso = 1.0")
            demanda = {cliente: 1.0 for cliente in df.columns[1:]}
            cds = df.iloc[:, 0].tolist()
            matriz_valores = df
        
        clientes = df.columns[1:].tolist()
        
        print(f"CDs encontrados: {cds}")
        print(f"Clientes encontrados: {clientes}")
        print(f"Demanda extraída: {demanda}")
        
        # 2. MATRIZ DE DISTÂNCIAS (mesma abordagem do p-Centros)
        distancias = {}
        for idx, cd in enumerate(cds):
            distancias[cd] = {}
            for cliente in clientes:
                distancias[cd][cliente] = float(matriz_valores.iloc[idx][cliente])
        
        print(f"Matriz de distâncias criada com {len(distancias)} CDs e {len(clientes)} clientes")
        
        # 3. CONJUNTO DE COBERTURA (quais CDs podem atender quais clientes dentro do raio)
        N_alcance = {cliente: [] for cliente in clientes}
        
        for cd in cds:
            for cliente in clientes:
                val = distancias[cd][cliente]
                
                # Raio normal ou Matriz Binária
                if raio_cobertura > 0:
                    if val <= raio_cobertura: 
                        N_alcance[cliente].append(cd)
                        print(f"✅ {cliente} coberto por {cd} (distância: {val} <= raio: {raio_cobertura})")
                else:
                    if val > 0: 
                        N_alcance[cliente].append(cd)
                        print(f"✅ {cliente} coberto por {cd} (binário: {val} > 0)")
        
        print(f"Conjuntos de alcance: {N_alcance}")

        # 4. MODELO MATEMÁTICO
        prob = LpProblem("Maxima_Cobertura", LpMaximize)
        
        # 5. VARIÁVEIS DE DECISÃO
        Y = LpVariable.dicts("Y", cds, cat='Binary')  # 1 se CD aberto, 0 se fechado
        Z = LpVariable.dicts("Z", clientes, cat='Binary')  # 1 se cliente coberto, 0 se não
        
        # 6. FUNÇÃO OBJETIVO: Maximizar demanda coberta
        prob += lpSum(demanda[cliente] * Z[cliente] for cliente in clientes), "Max_Demanda_Coberta"
        
        # 7. RESTRIÇÕES
        # R1: Abrir exatamente 'p' CDs
        prob += lpSum(Y[cd] for cd in cds) == p, "Qtd_CDs"
        
        # R2: Lógica de cobertura - cliente só é coberto se algum CD dentro do raio estiver aberto
        for cliente in clientes:
            if len(N_alcance[cliente]) == 0:
                # Se nenhum CD pode cobrir o cliente, força Z[cliente] = 0
                prob += Z[cliente] == 0, f"Sem_Cobertura_{cliente}"
            else:
                prob += Z[cliente] <= lpSum(Y[cd] for cd in N_alcance[cliente]), f"Cobertura_{cliente}"
                
        # 8. RESOLUÇÃO
        prob.solve()
        
        if LpStatus[prob.status] != 'Optimal':
            return {"status": "Erro", "mensagem": "Solução ótima não encontrada."}
            
        # 9. EXTRAÇÃO DE RESULTADOS
        cds_selecionados = [cd for cd in cds if value(Y[cd]) > 0.5]
        clientes_cobertos = [cliente for cliente in clientes if value(Z[cliente]) > 0.5]
        
        demanda_coberta = sum(demanda[cliente] for cliente in clientes_cobertos)
        demanda_total = sum(demanda.values())
        percentual_cobertura = (demanda_coberta / demanda_total * 100) if demanda_total > 0 else 0
        
        # Detalhar atribuições
        atribuicoes = []
        
        for cliente in clientes:
            coberto = value(Z[cliente]) > 0.5
            if coberto:
                # Encontrar qual CD atende o cliente (primeiro que encontrar dentro do raio)
                cd_atendente = [cd for cd in cds_selecionados if cd in N_alcance[cliente]][0]
                dist = distancias[cd_atendente][cliente]
            else:
                cd_atendente = "-"
                dist = "-"
                
            atribuicoes.append({
                "cliente": cliente,
                "cd": cd_atendente,
                "demanda": demanda[cliente],
                "distancia": dist,
                "coberto": coberto
            })
                    
        return {
            "status": "Sucesso",
            "p": p,
            "raio_cobertura": raio_cobertura if raio_cobertura > 0 else "Binário",
            "demanda_coberta": demanda_coberta,
            "percentual_cobertura": percentual_cobertura,
            "cds_selecionados": cds_selecionados,
            "atribuicoes": atribuicoes
        }
        
    except Exception as e:
        print(f"ERRO na otimização: {str(e)}")
        return {"status": "Erro", "mensagem": f"Erro interno: {str(e)}"}
