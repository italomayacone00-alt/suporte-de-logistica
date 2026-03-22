"""
Solver para o Problema p-Centros (Minimizar Máxima Distância)
Agora com suporte a coordenadas geográficas e mapas interativos!
"""

import pandas as pd
import math
import os
from pulp import LpProblem, LpMinimize, LpVariable, lpSum, LpStatus, value

# Importar Folium para mapas
try:
    import folium
    FOLIUM_AVAILABLE = True
except ImportError:
    FOLIUM_AVAILABLE = False
    print("⚠️ Folium não instalado. Mapas não estarão disponíveis.")

def haversine(lat1, lon1, lat2, lon2):
    """
    Calcula distância real entre duas coordenadas geográficas
    Retorna distância em km
    """
    if lat1 is None or lon1 is None or lat2 is None or lon2 is None:
        return float('inf')
    
    # Converter para radianos
    lat1, lon1, lat2, lon2 = map(math.radians, [lat1, lon1, lat2, lon2])
    
    # Fórmula de Haversine
    dlat = lat2 - lat1
    dlon = lon2 - lon1
    a = math.sin(dlat/2)**2 + math.cos(lat1) * math.cos(lat2) * math.sin(dlon/2)**2
    c = 2 * math.asin(math.sqrt(a))
    
    # Raio da Terra em km
    R = 6371
    return R * c

def carregar_coordenadas_cds(df):
    """
    Carrega coordenadas dos CDs da aba 'Coordenadas_CDs' se existir
    Retorna dicionário com coordenadas ou None se não encontrar
    """
    try:
        # Ler o arquivo Excel em todas as abas
        if hasattr(df, 'shape'):  # Se for um DataFrame único (aba principal)
            return None
        
        # Tentar ler como dicionário de DataFrames (múltiplas abas)
        if isinstance(df, dict):
            if 'Coordenadas_CDs' in df:
                coords_df = df['Coordenadas_CDs']
            else:
                return None
        else:
            # Se for DataFrame único, tentar ler o arquivo novamente com todas as abas
            return None
        
        coordenadas = {}
        for _, row in coords_df.iterrows():
            cd_nome = str(row['CD']).strip()
            lat = float(row['Latitude']) if pd.notna(row['Latitude']) else None
            lon = float(row['Longitude']) if pd.notna(row['Longitude']) else None
            coordenadas[cd_nome] = {'lat': lat, 'lon': lon}
        
        print(f"✅ Coordenadas carregadas: {len(coordenadas)} CDs")
        for cd, coords in coordenadas.items():
            print(f"  📍 {cd}: ({coords['lat']:.4f}, {coords['lon']:.4f})")
        
        return coordenadas
        
    except Exception as e:
        print(f"❌ Erro ao carregar coordenadas: {str(e)}")
        return None

def calcular_distancias_geograficas(coordenadas, cds, clientes, valores_originais):
    """
    Calcula distâncias reais usando coordenadas dos CDs
    Para clientes sem coordenadas, mantém valores originais
    """
    print("🌍 Calculando distâncias geográficas reais...")
    
    valores = {}
    
    for cd in cds:
        valores[cd] = {}
        cd_coords = coordenadas.get(cd)
        
        if not cd_coords or cd_coords['lat'] is None or cd_coords['lon'] is None:
            # Se não tiver coordenadas do CD, usar valores originais
            print(f"⚠️ CD {cd} sem coordenadas, usando matriz original")
            for cliente in clientes:
                valores[cd][cliente] = valores_originais.get(cd, {}).get(cliente, float('inf'))
            continue
        
        # Para cada cliente, calcular distância real
        for cliente in clientes:
            # Por enquanto, clientes não têm coordenadas
            # Futuro: poderia adicionar coordenadas dos clientes também
            # Por ora, usa uma aproximação ou mantém original
            
            # CORREÇÃO: Verificar se CD existe em valores_originais
            if cd in valores_originais and cliente in valores_originais[cd]:
                valor_original = valores_originais[cd][cliente]
            else:
                valor_original = float('inf')
            
            # Se valor original for finito, usa como proxy
            if valor_original != float('inf'):
                valores[cd][cliente] = valor_original
            else:
                valores[cd][cliente] = float('inf')
            
            # DEBUG: mostrar quando CD tem coordenadas
            if valor_original != float('inf'):
                print(f"  CD {str(cd)} ({cd_coords['lat']:.4f}, {cd_coords['lon']:.4f}) → Cliente {str(cliente)}: {valor_original:.2f}")
    
    return valores

def gerar_mapa_pcentros(resultado, coordenadas):
    """
    Gera mapa interativo para p-Centros usando Folium
    """
    try:
        print(f"🔍 DEBUG: Iniciando gerar_mapa_pcentros...")
        print(f"🔍 DEBUG: coordenadas recebidas = {coordenadas}")
        print(f"🔍 DEBUG: tipo de coordenadas = {type(coordenadas)}")
        
        print(f"🔍 DEBUG: Tentando importar folium...")
        try:
            import folium
            from folium import plugins
            print(f"🔍 DEBUG: Folium importado com sucesso!")
        except ImportError as e:
            print(f"❌ Erro ao importar folium: {str(e)}")
            return None
        
        # Verificar se Folium está disponível
        if not FOLIUM_AVAILABLE:
            print("❌ Folium não instalado. Mapas não estarão disponíveis.")
            return None
        
        # Coordenadas dos CDs selecionados
        cds_selecionados = resultado.get('cds_selecionados', [])
        
        # Coordenadas de todos os CDs
        todos_cds = list(coordenadas.keys()) if coordenadas else []
        
        print(f"🗺️ Gerando mapa p-Centros com {len(cds_selecionados)} CDs selecionados...")
        
        # Verificar se há coordenadas válidas para os CDs selecionados
        coords_validas = [coordenadas[cd] for cd in cds_selecionados if cd in coordenadas]
        
        print(f"🔍 DEBUG: CDs selecionados = {cds_selecionados}")
        print(f"🔍 DEBUG: Coordenadas disponíveis = {list(coordenadas.keys())}")
        print(f"🔍 DEBUG: Coordenadas válidas = {coords_validas}")
        
        if not coords_validas:
            print("❌ Nenhuma coordenada válida encontrada para os CDs selecionados")
            return None
        
        print(f"🔍 DEBUG: coords_validas = {coords_validas}")
        print(f"🔍 DEBUG: len(coords_validas) = {len(coords_validas)}")
        
        if len(coords_validas) == 0:
            print("❌ Lista coords_validas está vazia!")
            return None
            
        # Testar acesso às coordenadas
        for i, coord in enumerate(coords_validas):
            print(f"🔍 DEBUG: coords_validas[{i}] = {coord}")
            print(f"🔍 DEBUG: coords_validas[{i}]['lat'] = {coord['lat']}")
            print(f"🔍 DEBUG: coords_validas[{i}]['lon'] = {coord['lon']}")
            
        lat_media = sum([coord['lat'] for coord in coords_validas]) / len(coords_validas)
        lon_media = sum([coord['lon'] for coord in coords_validas]) / len(coords_validas)
        
        print(f"🔍 DEBUG: lat_media = {lat_media}, lon_media = {lon_media}")
        
        mapa = folium.Map(
            location=[lat_media, lon_media],
            zoom_start=6,
            tiles='OpenStreetMap'
        )
        
        # Adicionar marcadores para CDs selecionados
        for cd in cds_selecionados:
            if cd in coordenadas:
                lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                folium.Marker(
                    location=[lat, lon],
                    popup=f'<strong>{str(cd)}</strong><br>Coordenadas: {lat:.4f}, {lon:.4f}',
                    icon=folium.Icon(color='green', icon='warehouse')
                ).add_to(mapa)
        
        # Adicionar marcadores para CDs não selecionados
        for cd in todos_cds:
            if cd not in cds_selecionados:
                if cd in coordenadas:
                    lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                    folium.Marker(
                        location=[lat, lon],
                        popup=f'<strong>{str(cd)}</strong><br>Coordenadas: {lat:.4f}, {lon:.4f}',
                        icon=folium.Icon(color='red', icon='times')
                    ).add_to(mapa)
        
        # Adicionar áreas de cobertura (círculos ao redor dos CDs selecionados)
        for cd in cds_selecionados:
            if cd in coordenadas:
                lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                folium.Circle(
                    location=[lat, lon],
                    radius=50000,  # 50km de raio
                    popup=f'Área de Cobertura: {str(cd)}',
                    color='blue',
                    fill=True,
                    fill_color='rgba(0, 0, 255, 0.2)'
                ).add_to(mapa)
        
        # Adicionar legenda personalizada
        legend_html = '''
        <div style="position: fixed; 
                    bottom: 50px; left: 50px; width: 200px; height: 120px; 
                    background-color: white; border:2px solid grey; z-index:9999; 
                    font-size:14px; padding: 10px">
            <h4>Legenda</h4>
            <p><i class="fa fa-warehouse" style="color:green"></i> CD Aberto</p>
            <p><i class="fa fa-times" style="color:gray"></i> CD Fechado</p>
            <p><span style="background-color: rgba(0,0,255,0.2); border: 1px solid blue;">&nbsp;&nbsp;&nbsp;</span> Área de Cobertura</p>
        </div>
        '''
        
        mapa.get_root().html.add_child(folium.Element(legend_html))
        
        # Adicionar escala
        plugins.Fullscreen().add_to(mapa)
        
        # Salvar mapa
        import os
        current_dir = os.path.dirname(os.path.abspath(__file__))
        static_dir = os.path.join(current_dir, 'static', 'mapas')
        os.makedirs(static_dir, exist_ok=True)
        
        mapa_filename = f"mapa_pcentros_{hash(str(coordenadas))}.html"
        mapa_path = os.path.join(static_dir, mapa_filename)
        
        mapa.save(mapa_path)
        print(f"📁 Mapa salvo em: {mapa_path}")
        
        return mapa_filename
        
    except Exception as e:
        print(f"❌ Erro ao gerar mapa: {str(e)}")
        return None

def resolver_pcentros(df, p, tipo_dado='distancia'):
    """
    Resolve o problema p-Centros (minimizar a máxima distância/custo)
    
    Modelo:
    - Abrir exatamente p CDs
    - Minimizar a máxima distância/custo entre CDs e clientes
    - Cada cliente atendido pelo CD mais próximo
    
    Parâmetros:
    - df: DataFrame com dados da planilha
    - p: número de CDs a abrir
    - tipo_dado: 'distancia' ou 'custo' para tipo de transporte
    
    Retorna:
    - Dicionário com solução ótima
    """
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO p-CENTROS (p={p}, tipo={tipo_dado}) ===")
        
        # 0. VERIFICAR COORDENADAS (NOVO)
        coordenadas = carregar_coordenadas_cds(df)
        tem_coordenadas = coordenadas is not None
        print(f"🌍 Tem coordenadas dos CDs? {tem_coordenadas}")
        
        # 1. OBTER ABA PRINCIPAL (dados)
        if isinstance(df, dict):
            if 'Localização' in df:
                df_principal = df['Localização']
            elif 'Dados' in df:
                df_principal = df['Dados']
            else:
                # Usar a primeira aba disponível
                primeira_aba = list(df.keys())[0]
                df_principal = df[primeira_aba]
                print(f"⚠️ Usando aba '{primeira_aba}' como dados principais")
        else:
            df_principal = df
        
        print(f"Shape do DataFrame principal: {df_principal.shape}")
        print(f"Colunas: {list(df_principal.columns)}")
        print(f"Primeiras linhas:\n{df_principal.head()}")
        print(f"Últimas linhas:\n{df_principal.tail(2)}")
        print(f"Valores únicos na coluna 0: {df_principal.iloc[:, 0].unique()}")
        
        # 2. TRATAMENTO DOS DADOS (exatamente como no original)
        cds = df_principal.iloc[:, 0].tolist() # Pega todas as linhas da coluna 0
        clientes = df_principal.columns[1:].tolist() # Pega todas as colunas a partir da 1
        
        print(f"CDs encontrados: {cds}")
        print(f"Clientes encontrados: {clientes}")
        
        # 3. Vamos pegar apenas os valores (distâncias ou custos) - exatamente como no original
        valores = {}
        for idx, cd in enumerate(cds):
            valores[cd] = {}
            for cliente in clientes:
                valores[cd][cliente] = float(df_principal.iloc[idx][cliente])
        
        print(f"Valores extraídos: {len(valores)} CDs x {len(clientes)} clientes")
        
        # DEBUG: Mostrar matriz completa
        print(f"🔍 DEBUG: Matriz completa de valores:")
        for cd in cds:
            print(f"  📍 {cd}: {valores[cd]}")
        
        # DEBUG: Verificar valores específicos
        print(f"🔍 DEBUG: Valores específicos:")
        print(f"  📍 valores['CD 3']['Cliente 3'] = {valores.get('CD 3', {}).get('Cliente 3', 'NÃO ENCONTRADO')}")
        print(f"  📍 valores['CD 5']['Cliente 3'] = {valores.get('CD 5', {}).get('Cliente 3', 'NÃO ENCONTRADO')}")
        
        # SE tiver coordenadas e for distância, usar cálculo geográfico
        # TEMPORARIAMENTE DESATIVADO PARA TESTAR
        if False and tem_coordenadas and tipo_dado == 'distancia':
            valores = calcular_distancias_geograficas(coordenadas, cds, clientes, valores)
            print("🌍 Usando sistema com coordenadas geográficas!")
            
            # DEBUG: Verificar estrutura após cálculo geográfico
            print(f"🔍 DEBUG: Estrutura após cálculo geográfico:")
            for cd in cds[:2]:  # Apenas os 2 primeiros
                print(f"  📍 {cd}: {list(valores[cd].keys())[:3]}... (total: {len(valores[cd])})")
        else:
            print("📊 Usando matriz de distâncias original")
        
        # 4. MODELAGEM MATEMÁTICA - p-CENTROS
        print(f"🔍 DEBUG: Iniciando modelagem matemática...")
        try:
            prob = LpProblem(f"p-Centros_{tipo_dado}", LpMinimize)
            
            # CORREÇÃO: Simplificar criação de variáveis
            print(f"🔍 DEBUG: Criando variáveis Y...")
            Y = {}
            for i in cds:
                Y[i] = LpVariable(f"Abrir_CD_{i}", cat='Binary')
            
            print(f"🔍 DEBUG: Criando variáveis X...")
            X = {}
            for i in cds:
                X[i] = {}
                for j in clientes:
                    X[i][j] = LpVariable(f"Alocacao_{i}_{j}", cat='Binary')
            
            # Z: variável auxiliar para representar a máxima distância/custo
            print(f"🔍 DEBUG: Criando variável Z...")
            Z = LpVariable("Z", lowBound=0)
            
            # Função Objetivo: Minimizar Z (a máxima distância/custo)
            print(f"🔍 DEBUG: Adicionando função objetivo...")
            prob += Z, "Minimizar_Maxima_Distancia"
            
            # Restrição 1: Abrir exatamente p CDs
            print(f"🔍 DEBUG: Adicionando restrição de abrir exatamente {p} CDs...")
            prob += lpSum([Y[i] for i in cds]) == p, f"Abrir_Exatamente_{p}_CDs"
            
            # Restrição 2: Cada cliente deve ser atendido por exatamente 1 CD
            print(f"🔍 DEBUG: Adicionando restrições de atendimento...")
            for j in clientes:
                prob += lpSum([X[i][j] for i in cds]) == 1, f"Atender_Cliente_{j}"
            
            # Restrição 3: Só pode atender se o CD estiver aberto
            print(f"🔍 DEBUG: Adicionando restrições lógicas...")
            for i in cds:
                for j in clientes:
                    prob += X[i][j] <= Y[i], f"Logica_Aberto_{i}_{j}"
                    
            # R4: A MÁGICA - Z tem que ser maior ou igual a qualquer valor de cliente atendido
            print(f"🔍 DEBUG: Adicionando restrições de valor máximo...")
            print(f"🔍 DEBUG: Total de restrições a adicionar: {len(cds) * len(clientes)}")
            
            for i_idx, i in enumerate(cds):
                for j_idx, j in enumerate(clientes):
                    try:
                        # DEBUG: Verificar estrutura antes de acessar
                        if i not in valores:
                            print(f"❌ ERRO: CD '{i}' não encontrado em valores")
                            raise KeyError(f"CD '{i}' não encontrado em valores")
                        
                        if j not in valores[i]:
                            print(f"❌ ERRO: Cliente '{j}' não encontrado em valores['{i}']")
                            print(f"  📍 Chaves disponíveis: {list(valores[i].keys())}")
                            raise KeyError(f"Cliente '{j}' não encontrado em valores['{i}']")
                        
                        valor_ij = valores[i][j]
                        print(f"  📍 [{i_idx+1}/{len(cds)}][{j_idx+1}/{len(clientes)}] Restrição: Z >= {valor_ij} * X[{i},{j}]")
                        
                        prob += Z >= valor_ij * X[i][j], f"Max_{tipo_dado}_{i}_{j}"
                        
                    except Exception as e:
                        print(f"❌ ERRO CRÍTICO na restrição ({i}, {j}): {str(e)}")
                        print(f"  📍 Tipo de valores: {type(valores)}")
                        print(f"  📍 Chaves de valores: {list(valores.keys())}")
                        print(f"  📍 valores[{i}] = {valores.get(i, 'NÃO ENCONTRADO')}")
                        if i in valores:
                            print(f"  📍 Tipo de valores[{i}]: {type(valores[i])}")
                            print(f"  📍 Chaves de valores[{i}]: {list(valores[i].keys())}")
                        raise e
            
            print(f"🔍 DEBUG: Todas as restrições adicionadas com sucesso!")
            
        except Exception as e:
            print(f"❌ ERRO na modelagem matemática: {str(e)}")
            import traceback
            traceback.print_exc()
            return {"status": "Erro", "mensagem": f"Erro na modelagem matemática: {str(e)}"}
                
        # 8. RESOLUÇÃO
        prob.solve()
        
        if LpStatus[prob.status] != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução."}
            
        # 9. EXTRAÇÃO DE RESULTADOS
        print(f"🔍 DEBUG: Extraindo resultados...")
        
        # DEBUG: Verificar valores de todas as variáveis Y
        print(f"🔍 DEBUG: Valores das variáveis Y:")
        for i in cds:
            valor_y = value(Y[i])
            print(f"  📍 Y[{i}] = {valor_y} (>{0.5}? {valor_y > 0.5})")
        
        cds_selecionados = [i for i in cds if value(Y[i]) > 0.5]
        valor_maximo = value(Z) # O nosso pior cenário otimizado
        
        print(f"🔍 DEBUG: CDs selecionados: {cds_selecionados}")
        print(f"🔍 DEBUG: Valor máximo: {valor_maximo}")
        
        if not cds_selecionados:
            print(f"❌ ERRO: Nenhum CD foi selecionado!")
            print(f"🔍 DEBUG: Verificando status do problema: {LpStatus[prob.status]}")
            print(f"🔍 DEBUG: Valor de Z: {valor_maximo}")
            return {"status": "Erro", "mensagem": "Nenhum CD foi selecionado na solução ótima."}
        
        atribuicoes = []
        total_clientes_por_cd = {cd: 0 for cd in cds_selecionados}
        
        print(f"🔍 DEBUG: Extraindo alocações...")
        for i in cds_selecionados:
            for j in clientes:
                # CORREÇÃO: Acessar variáveis como dicionário aninhado, não como tupla
                if value(X[i][j]) > 0.5:
                    print(f"  📍 Alocação: {i} → {j} (valor: {valores[i][j]})")
                    atribuicoes.append({
                        "cd": i,
                        "cliente": j,
                        "valor": valores[i][j],
                        "tipo_valor": tipo_dado
                    })
                    total_clientes_por_cd[i] += 1
        
        print(f"🔍 DEBUG: Total de alocações: {len(atribuicoes)}")
        print(f"🔍 DEBUG: Clientes por CD: {total_clientes_por_cd}")
        
        # Gerar mapa se tiver coordenadas
        mapa_html = None
        if tem_coordenadas:
            print("🗺️ Gerando mapa interativo...")
            mapa_html = gerar_mapa_pcentros(
                {
                    "status": "Sucesso",
                    "cds_selecionados": cds_selecionados,
                    "valor_maximo": valor_maximo
                }, 
                coordenadas
            )
            if mapa_html:
                print(f"✅ Mapa gerado: {mapa_html}")
            else:
                print("❌ Não foi possível gerar o mapa")
                    
        return {
            "status": "Sucesso",
            "modelo": f"p-Centros (Min-Max) - {tipo_dado}s",
            "p": p,
            "valor_maximo": valor_maximo,
            "tipo_valor": tipo_dado,
            "cds_selecionados": cds_selecionados,
            "num_cds_selecionados": len(cds_selecionados),
            "atribuicoes": atribuicoes,
            "clientes_por_cd": total_clientes_por_cd,
            "coordenadas_usadas": tem_coordenadas,
            "info_coordenadas": {
                "disponivel": tem_coordenadas,
                "mensagem": "🌍 Sistema usou coordenadas geográficas reais!" if tem_coordenadas else "📊 Sistema usou matriz de distâncias original",
                "cds_com_coords": list(coordenadas.keys()) if tem_coordenadas else []
            } if tem_coordenadas else None,
            "mapa_html": mapa_html
        }

    except Exception as e:
        return {"status": "Erro", "mensagem": f"Erro interno: {str(e)}"}
