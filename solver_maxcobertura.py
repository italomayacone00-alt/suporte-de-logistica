"""
Solver Genérico para Máxima Cobertura (MCLP)
Baseado na formulação de Church e ReVelle (1974)
"""
import pandas as pd
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpStatus, value

def resolver_maxcobertura(df, p, raio_cobertura=None):
    """
    Resolve o Problema de Máxima Cobertura (MCLP) de forma genérica
    
    Parâmetros:
    - df: DataFrame com matriz de distâncias e pesos
    - p: número de instalações que podem ser abertas
    - raio_cobertura: distância máxima para considerar um ponto coberto (None para matriz binária)
    
    Retorna:
    - Dicionário com resultados da otimização
    """
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO MÁXIMA COBERTURA (p={p}, raio={raio_cobertura}) ===")
        print(f"Shape do DataFrame: {df.shape}")
        print(f"Colunas: {list(df.columns)}")
        print(f"Primeiras linhas:\n{df.head()}")
        
        # 1. TRATAMENTO DOS DADOS GENÉRICOS
        pontos = df.columns[1:].tolist()
        ultima_linha_nome = str(df.iloc[-1, 0]).strip().lower()
        
        # Aceita 'peso', 'demanda', 'peso / demanda', 'população', etc.
        if ultima_linha_nome in ['demanda', 'peso', 'peso / demanda', 'população']:
            pesos = {ponto: float(df.iloc[-1][ponto]) for ponto in pontos}
            instalacoes = df.iloc[:-1, 0].tolist()
            matriz_valores = df.iloc[:-1]
        else:
            # Se não tiver linha de peso, assume peso = 1 para todos
            pesos = {ponto: 1.0 for ponto in pontos}
            instalacoes = df.iloc[:, 0].tolist()
            matriz_valores = df
        
        print(f"Instalações encontradas: {instalacoes}")
        print(f"Pontos encontrados: {pontos}")
        print(f"Pesos extraídos: {pesos}")
        
        # Criar matriz de valores (distâncias ou binário)
        distancias = {}
        for idx, inst in enumerate(instalacoes):
            distancias[inst] = {}
            for ponto in pontos:
                try:
                    distancias[inst][ponto] = float(matriz_valores.iloc[idx][ponto])
                except (ValueError, TypeError):
                    distancias[inst][ponto] = float('inf')  # Distância infinita se não for número
        
        print(f"Matriz de valores criada")
        
        # 2. CONJUNTO DE COBERTURA
        # Para cada ponto, lista quais instalações podem cobri-lo
        N_alcance = {ponto: [] for ponto in pontos}
        for inst in instalacoes:
            for ponto in pontos:
                val = distancias[inst][ponto]
                
                if raio_cobertura is not None and raio_cobertura > 0:
                    # Modo com raio de cobertura
                    if val <= raio_cobertura:
                        N_alcance[ponto].append(inst)
                else:
                    # Modo matriz binária (1 = cobre, 0 = não cobre)
                    if val > 0:
                        N_alcance[ponto].append(inst)
        
        print(f"Conjuntos de alcance criados")
        
        # 3. MODELO MATEMÁTICO
        prob = LpProblem("Maxima_Cobertura_Geral", LpMaximize)
        
        # 4. VARIÁVEIS DE DECISÃO
        Y = LpVariable.dicts("Y", instalacoes, cat='Binary')  # 1 se instalação aberta, 0 se fechada
        Z = LpVariable.dicts("Z", pontos, cat='Binary')      # 1 se ponto coberto, 0 se não
        
        # 5. FUNÇÃO OBJETIVO: Maximizar peso total coberto
        prob += lpSum(pesos[ponto] * Z[ponto] for ponto in pontos), "Maximizar_Peso_Coberto"
        
        # 6. RESTRIÇÕES
        # R1: Abrir exatamente 'p' instalações
        prob += lpSum(Y[inst] for inst in instalacoes) == p, "Qtd_Instalacoes"
        
        # R2: Lógica de cobertura - ponto só é coberto se alguma instalação dentro do raio estiver aberta
        for ponto in pontos:
            if len(N_alcance[ponto]) == 0:
                # Se nenhuma instalação pode cobrir o ponto, força Z[ponto] = 0
                prob += Z[ponto] == 0, f"Sem_Cobertura_{ponto}"
            else:
                prob += Z[ponto] <= lpSum(Y[inst] for inst in N_alcance[ponto]), f"Cobertura_{ponto}"
        
        # 7. RESOLUÇÃO
        prob.solve()
        
        if LpStatus[prob.status] != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução ótima."}
            
        # 8. EXTRAÇÃO DE RESULTADOS
        instalacoes_selecionadas = [inst for inst in instalacoes if value(Y[inst]) > 0.5]
        pontos_cobertos = [ponto for ponto in pontos if value(Z[ponto]) > 0.5]
        pontos_nao_cobertos = [ponto for ponto in pontos if value(Z[ponto]) <= 0.5]
        
        # Calcular métricas
        peso_total = sum(pesos.values())
        peso_coberto = sum(pesos[ponto] for ponto in pontos_cobertos)
        percentual_cobertura = (peso_coberto / peso_total * 100) if peso_total > 0 else 0
        
        # Detalhar atribuições
        atribuicoes = []
        for ponto in pontos_cobertos:
            # Encontrar qual instalação atende o ponto (primeiro que encontrar dentro do raio)
            for inst in instalacoes_selecionadas:
                if inst in N_alcance[ponto]:
                    atribuicoes.append({
                        "instalacao": inst,
                        "ponto": ponto,
                        "distancia": distancias[inst][ponto],
                        "peso": pesos[ponto],
                        "coberto": True
                    })
                    break
        
        # Pontos não cobertos
        for ponto in pontos_nao_cobertos:
            # Encontrar instalação mais próxima (mesmo que fora do raio)
            inst_mais_proxima = min(instalacoes, key=lambda i: distancias[i][ponto])
            atribuicoes.append({
                "instalacao": inst_mais_proxima,
                "ponto": ponto,
                "distancia": distancias[inst_mais_proxima][ponto],
                "peso": pesos[ponto],
                "coberto": False
            })
        
        return {
            "status": "Sucesso",
            "modelo": "Máxima Cobertura (MCLP)",
            "p": p,
            "raio_cobertura": raio_cobertura,
            "instalacoes_selecionadas": instalacoes_selecionadas,
            "num_instalacoes_selecionadas": len(instalacoes_selecionadas),
            "pontos_cobertos": pontos_cobertos,
            "pontos_nao_cobertos": pontos_nao_cobertos,
            "num_pontos_cobertos": len(pontos_cobertos),
            "num_pontos_nao_cobertos": len(pontos_nao_cobertos),
            "peso_total": peso_total,
            "peso_coberto": peso_coberto,
            "percentual_cobertura": percentual_cobertura,
            "atribuicoes": atribuicoes,
            "total_pontos": len(pontos)
        }
        
    except Exception as e:
        print(f"ERRO na otimização: {str(e)}")
        return {"status": "Erro", "mensagem": f"Erro durante a otimização: {str(e)}"}
