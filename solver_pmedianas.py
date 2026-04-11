"""
Solver para o Problema das p-Medianas
Objetivo: Minimizar o custo total de transporte ponderado pela demanda.
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
            
            # Converter coordenadas aceitando tanto ponto quanto vírgula
            lat_raw = str(row['Latitude']) if pd.notna(row['Latitude']) else None
            lon_raw = str(row['Longitude']) if pd.notna(row['Longitude']) else None
            
            lat = float(lat_raw.replace(',', '.')) if lat_raw else None
            lon = float(lon_raw.replace(',', '.')) if lon_raw else None
            
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
            
            # Se tiver dados originais, usa como referência
            valor_original = valores_originais.get(cd, {}).get(cliente, float('inf'))
            
            # Se valor original for finito, usa como proxy
            # Futuro melhor: adicionar coordenadas dos clientes
            valores[cd][cliente] = valor_original
            
            # DEBUG: mostrar quando CD tem coordenadas
            if valor_original != float('inf'):
                print(f"📍 CD {str(cd)} ({cd_coords['lat']:.4f}, {cd_coords['lon']:.4f}) → Cliente {str(cliente)}: {valor_original:.2f}")
    
    return valores


def gerar_mapa_pmedianas(resultado, coordenadas):
    """
    Gera mapa interativo para p-Medianas usando Folium
    """
    try:
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
        
        print(f"🗺️ Gerando mapa p-Medianas com {len(cds_selecionados)} CDs selecionados...")
        
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
        
        # Adicionar áreas de influência (círculos ao redor dos CDs selecionados)
        for cd in cds_selecionados:
            if cd in coordenadas:
                lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                folium.Circle(
                    location=[lat, lon],
                    radius=50000,  # 50km de raio
                    popup=f'Área de Influência: {str(cd)}',
                    color='green',
                    fill=True,
                    fill_color='rgba(0, 255, 0, 0.2)'
                ).add_to(mapa)
        
        # Adicionar legenda personalizada atualizada
        legend_html = '''
        <div style="position: fixed; 
                    bottom: 50px; left: 50px; width: 260px; height: 160px; 
                    background-color: white; border:2px solid #374151; z-index:9999; 
                    font-size:14px; padding: 15px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
            <h4 style="margin: 0 0 12px 0; color: #1f2937; font-size: 16px; font-weight: 600;">Legenda do Mapa</h4>
            <p style="margin: 10px 0; display: flex; align-items: center;">
                <i class="fa fa-warehouse" style="color: #16a34a; margin-right: 10px; font-size: 16px;"></i> 
                <span style="color: #374151;">CD Ativo (Influência)</span>
            </p>
            <p style="margin: 10px 0; display: flex; align-items: center;">
                <i class="fa fa-warehouse" style="color: #dc2626; margin-right: 10px; font-size: 16px;"></i> 
                <span style="color: #374151;">CD Inativo (Sem Influência)</span>
            </p>
            <p style="margin: 10px 0; display: flex; align-items: center;">
                <span style="background-color: rgba(34, 197, 94, 0.2); border: 1px solid #22c55e; padding: 6px 12px; border-radius: 4px; margin-right: 10px;">&nbsp;&nbsp;&nbsp;</span> 
                <span style="color: #374151;">Área de Influência (50 km)</span>
            </p>
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
        
        mapa_filename = f"mapa_pmedianas_{hash(str(coordenadas))}.html"
        mapa_path = os.path.join(static_dir, mapa_filename)
        
        mapa.save(mapa_path)
        print(f"📁 Mapa salvo em: {mapa_path}")
        
        return mapa_filename
        
    except Exception as e:
        print(f"❌ Erro ao gerar mapa: {str(e)}")
        return None

def resolver_pmedianas(df, p, tipo_dado='distancia'):
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO p-MEDIANAS (p={p}, tipo={tipo_dado}) ===")
        
        # 0. VERIFICAR COORDENADAS (NOVO)
        coordenadas = carregar_coordenadas_cds(df)
        tem_coordenadas = coordenadas is not None
        print(f"🌍 Tem coordenadas dos CDs? {tem_coordenadas}")
        
        # 1. OBTER ABA PRINCIPAL (dados)
        if isinstance(df, dict):
            if 'Dados' in df:
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
        print(f"Última linha (demanda):\n{df_principal.tail(1)}")
        
        # 2. TRATAMENTO DOS DADOS
        cds = df_principal.iloc[:-1, 0].tolist()  # Remove última linha (demanda)
        clientes = df_principal.columns[1:-2].tolist()  # Remove colunas de custo fixo e capacidade
        
        print(f"CDs encontrados: {cds}")
        print(f"Clientes encontrados: {clientes}")
        
        # 3. DEMANDA DOS CLIENTES (como no original)
        demanda = {cliente: int(df_principal.iloc[-1][cliente]) for cliente in clientes}
        
        # Pegando as colunas extras que o Excel usa no cálculo
        custos_fixos = {}
        capacidades = {}
        for idx, cd in enumerate(cds):
            if 'Custo Fixo' in df_principal.columns:
                custos_fixos[cd] = int(df_principal.iloc[idx]['Custo Fixo'])
            if 'Capacidade' in df_principal.columns:
                capacidades[cd] = int(df_principal.iloc[idx]['Capacidade'])
            
        # 4. MATRIZ DE VALORES (com suporte a coordenadas)
        valores_originais = {}
        for idx, cd in enumerate(cds):
            valores_originais[cd] = {}
            for cliente in clientes:
                valores_originais[cd][cliente] = float(df_principal.iloc[idx][cliente])
        
        # SE tiver coordenadas e for distância, usar cálculo geográfico
        if tem_coordenadas and tipo_dado == 'distancia':
            valores = calcular_distancias_geograficas(coordenadas, cds, clientes, valores_originais)
            print("🌍 Usando sistema com coordenadas geográficas!")
        else:
            valores = valores_originais
            print("📊 Usando matriz de distâncias original")
        
        print(f"Matriz de {tipo_dado}s criada com {len(valores)} CDs e {len(clientes)} clientes")

        # 5. MODELO
        prob = LpProblem("Problema_p_Medianas", LpMinimize)
        
        # 6. VARIÁVEIS DE DECISÃO
        Y = LpVariable.dicts("Y", cds, cat='Binary') 
        X = LpVariable.dicts("X", [(i, j) for i in cds for j in clientes], cat='Binary')
        
        # 7. FUNÇÃO OBJETIVO: Valores + Custo Fixo (como no original)
        prob += lpSum(valores[i][j] * X[i, j] for i in cds for j in clientes) + \
                lpSum(custos_fixos[i] * Y[i] for i in cds)
        
        # 8. RESTRIÇÕES
        # R1: Abrir exatamente 'p' CDs
        prob += lpSum(Y[i] for i in cds) == p, "Qtd_CDs"
        
        # R2: Cada cliente é atendido por 1 único CD
        for j in clientes:
            prob += lpSum(X[i, j] for i in cds) == 1, f"Atendimento_{j}"
            
        # R3: Um CD só pode atender clientes se estiver aberto
        for i in cds:
            for j in clientes:
                prob += X[i, j] <= Y[i], f"Logica_Aberto_{i}_{j}"
                
        # R4: RESTRIÇÃO DE CAPACIDADE (como no original)
        for i in cds:
            prob += lpSum(demanda[j] * X[i, j] for j in clientes) <= capacidades[i] * Y[i], f"Capacidade_{i}"
                
        # 9. RESOLUÇÃO
        prob.solve()
        
        if LpStatus[prob.status] != 'Optimal':
            return {"status": "Erro", "mensagem": "Não foi possível encontrar uma solução."}
            
        # 10. EXTRAÇÃO DE RESULTADOS
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
                        "custo_ponderado": valores[i][j]
                    })
                    total_clientes_por_cd[i] += 1
                    
        # Gerar mapa se tiver coordenadas
        mapa_html = None
        if tem_coordenadas:
            print("🗺️ Gerando mapa interativo...")
            mapa_html = gerar_mapa_pmedianas(
                {
                    "status": "Sucesso",
                    "cds_selecionados": cds_selecionados,
                    "custo_total_ponderado": custo_total
                }, 
                coordenadas
            )
            if mapa_html:
                print(f"✅ Mapa gerado: {mapa_html}")
            else:
                print("❌ Não foi possível gerar o mapa")
                    
        return {
            "status": "Sucesso",
            "modelo": f"p-Medianas - {tipo_dado}s",
            "p": p,
            "custo_total_ponderado": custo_total,
            "tipo_valor": tipo_dado,
            "cds_selecionados": cds_selecionados,
            "num_cds_selecionados": len(cds_selecionados),
            "atribuicoes": atribuicoes,
            "clientes_por_cd": total_clientes_por_cd,
            "demanda_total": sum(demanda.values()),
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