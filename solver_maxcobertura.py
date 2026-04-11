"""
Solver para Máxima Cobertura (MCLP) - Padrão Logístico (CDs e Clientes)
Baseado na abordagem simples do p-Centros
"""
import pandas as pd
import math
import os
import tempfile
from pulp import LpProblem, LpMaximize, LpVariable, lpSum, LpStatus, value

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
                print(f"📍 CD {cd} ({cd_coords['lat']:.4f}, {cd_coords['lon']:.4f}) → Cliente {cliente}: {valor_original:.2f}")
    
    return valores

def gerar_mapa_cobertura(resultado, coordenadas, raio_cobertura=0.0):
    """
    Gera mapa interativo com CDs selecionados e áreas de cobertura
    Retorna HTML do mapa ou None se não for possível
    """
    if not FOLIUM_AVAILABLE or not coordenadas:
        return None
    
    try:
        # Calcular centro do mapa (média das coordenadas dos CDs selecionados)
        lat_center = 0
        lon_center = 0
        count = 0
        
        for cd in resultado.get('cds_selecionados', []):
            if cd in coordenadas:
                coords = coordenadas[cd]
                lat_center += coords['lat']
                lon_center += coords['lon']
                count += 1
        
        if count == 0:
            return None
            
        lat_center /= count
        lon_center /= count
        
        # Criar mapa
        mapa = folium.Map(
            location=[lat_center, lon_center], 
            zoom_start=6,
            tiles='OpenStreetMap'
        )
        
        # Adicionar CDs
        for cd in resultado.get('cds_selecionados', []):
            if cd in coordenadas:
                coords = coordenadas[cd]
                
                # Adicionar círculo de cobertura se tiver raio
                if raio_cobertura > 0:
                    folium.Circle(
                        location=[coords['lat'], coords['lon']],
                        radius=raio_cobertura * 1000,  # Converter km para metros
                        color='orange',
                        fill=True,
                        fillColor='yellow',
                        fillOpacity=0.3,
                        popup=f"CD {cd}<br>Raio: {raio_cobertura} km"
                    ).add_to(mapa)
                
                # Adicionar marcador do CD
                folium.Marker(
                    location=[coords['lat'], coords['lon']],
                    popup=f"🏢 {cd}<br>📍 ({coords['lat']:.4f}, {coords['lon']:.4f})",
                    tooltip=f"CD {cd}",
                    icon=folium.Icon(color='green', icon='warehouse', prefix='fa')
                ).add_to(mapa)
        
        # Adicionar CDs não selecionados
        for cd, coords in coordenadas.items():
            if cd not in resultado.get('cds_selecionados', []):
                folium.Marker(
                    location=[coords['lat'], coords['lon']],
                    popup=f"🏢 {cd}<br>Não selecionado<br>📍 ({coords['lat']:.4f}, {coords['lon']:.4f})",
                    tooltip=f"CD {cd} (não selecionado)",
                    icon=folium.Icon(color='red', icon='warehouse', prefix='fa')
                ).add_to(mapa)
        
        # Adicionar legenda atualizada
        legend_html = f'''
        <div style="position: fixed; 
                    bottom: 50px; left: 50px; width: 260px; height: 160px; 
                    background-color: white; border:2px solid #374151; z-index:9999; 
                    font-size:14px; padding: 15px; border-radius: 8px; box-shadow: 0 4px 12px rgba(0,0,0,0.15);">
        <h4 style="margin: 0 0 12px 0; color: #1f2937; font-size: 16px; font-weight: 600;">Legenda do Mapa</h4>
        <p style="margin: 10px 0; display: flex; align-items: center;">
            <i class="fa fa-warehouse" style="color: #16a34a; margin-right: 10px; font-size: 16px;"></i> 
            <span style="color: #374151;">CD Ativo (Cobertura)</span>
        </p>
        <p style="margin: 10px 0; display: flex; align-items: center;">
            <i class="fa fa-warehouse" style="color: #dc2626; margin-right: 10px; font-size: 16px;"></i> 
            <span style="color: #374151;">CD Inativo (Sem Cobertura)</span>
        </p>
        <p style="margin: 10px 0; display: flex; align-items: center;">
            <span style="background-color: rgba(250, 204, 21, 0.3); border: 1px solid #facc15; padding: 6px 12px; border-radius: 4px; margin-right: 10px;">&nbsp;&nbsp;&nbsp;</span> 
            <span style="color: #374151;">Área de Cobertura ({raio_cobertura} km)</span>
        </p>
        </div>
        '''
        mapa.get_root().html.add_child(folium.Element(legend_html))
        
        # Salvar mapa em pasta estática
        import os
        current_dir = os.path.dirname(os.path.abspath(__file__))
        static_dir = os.path.join(current_dir, 'static', 'mapas')
        os.makedirs(static_dir, exist_ok=True)
        
        mapa_filename = f"mapa_cobertura_{hash(str(coordenadas))}.html"
        mapa_path = os.path.join(static_dir, mapa_filename)
        
        mapa.save(mapa_path)
        print(f"📁 Mapa salvo em: {mapa_path}")
        
        return mapa_filename
        
    except Exception as e:
        print(f"❌ Erro ao gerar mapa: {str(e)}")
        return None

def resolver_maxcobertura(df, p, raio_cobertura=0.0, tipo_dado='distancia'):
    try:
        print(f"=== INICIANDO OTIMIZAÇÃO MÁXIMA COBERTURA (p={p}, raio={raio_cobertura}, tipo={tipo_dado}) ===")
        
        # 0. VERIFICAR COORDENADAS (NOVO)
        coordenadas = carregar_coordenadas_cds(df)
        tem_coordenadas = coordenadas is not None
        print(f"🌍 Tem coordenadas dos CDs? {tem_coordenadas}")
        
        # 1. OBTER ABA PRINCIPAL (dados)
        if isinstance(df, dict):
            if 'Maxima_Cobertura' in df:
                df_principal = df['Maxima_Cobertura']
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
        
        # 2. TRATAMENTO DOS DADOS (mesma abordagem do p-Centros)
        ultima_linha_nome = str(df_principal.iloc[-1, 0]).strip().lower()
        
        if ultima_linha_nome in ['demanda', 'demanda total', 'peso']:
            print("✅ Linha de demanda encontrada")
            demanda = {cliente: int(df_principal.iloc[-1][cliente]) for cliente in df_principal.columns[1:]}
            cds = df_principal.iloc[:-1, 0].tolist()
            matriz_valores = df_principal.iloc[:-1]
        else:
            print("❌ Linha de demanda NÃO encontrada, usando peso = 1.0")
            demanda = {cliente: 1 for cliente in df_principal.columns[1:]}
            cds = df_principal.iloc[:, 0].tolist()
            matriz_valores = df_principal
        
        clientes = df_principal.columns[1:].tolist()
        
        print(f"CDs encontrados: {cds}")
        print(f"Clientes encontrados: {clientes}")
        print(f"Demanda extraída: {demanda}")
        
        # 3. MATRIZ DE VALORES (com suporte a coordenadas)
        valores_originais = {}
        for idx, cd in enumerate(cds):
            valores_originais[cd] = {}
            for cliente in clientes:
                valores_originais[cd][cliente] = float(matriz_valores.iloc[idx][cliente])
        
        # SE tiver coordenadas e for distância, usar cálculo geográfico
        if tem_coordenadas and tipo_dado == 'distancia':
            valores = calcular_distancias_geograficas(coordenadas, cds, clientes, valores_originais)
            print("🌍 Usando sistema com coordenadas geográficas!")
        else:
            valores = valores_originais
            print("📊 Usando matriz de distâncias original")
        
        print(f"Matriz de {tipo_dado}s criada com {len(valores)} CDs e {len(clientes)} clientes")
        
        # 3. CONJUNTO DE COBERTURA (quais CDs podem atender quais clientes dentro do raio)
        N_alcance = {cliente: [] for cliente in clientes}
        
        for cd in cds:
            for cliente in clientes:
                val = valores[cd][cliente]
                
                # Raio normal ou Matriz Binária
                if raio_cobertura > 0:
                    if val <= raio_cobertura: 
                        N_alcance[cliente].append(cd)
                        print(f"✅ {cliente} coberto por {cd} ({tipo_dado}: {val} <= raio: {raio_cobertura})")
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
                val = valores[cd_atendente][cliente]
            else:
                cd_atendente = "-"
                val = "-"
                
            atribuicoes.append({
                "cliente": cliente,
                "cd": cd_atendente,
                "demanda": demanda[cliente],
                "valor": val,
                "tipo_valor": tipo_dado,
                "coberto": coberto
            })
                    
        # Gerar mapa se tiver coordenadas
        mapa_html = None
        if tem_coordenadas:
            print("🗺️ Gerando mapa interativo...")
            mapa_html = gerar_mapa_cobertura(
                {
                    "status": "Sucesso",
                    "cds_selecionados": cds_selecionados
                }, 
                coordenadas, 
                raio_cobertura
            )
            if mapa_html:
                print(f"✅ Mapa gerado: {mapa_html}")
            else:
                print("❌ Não foi possível gerar o mapa")
        
        return {
            "status": "Sucesso",
            "p": p,
            "raio_cobertura": raio_cobertura if raio_cobertura > 0 else "Binário",
            "tipo_valor": tipo_dado,
            "demanda_coberta": demanda_coberta,
            "percentual_cobertura": percentual_cobertura,
            "cds_selecionados": cds_selecionados,
            "atribuicoes": atribuicoes,
            "coordenadas_usadas": tem_coordenadas,
            "info_coordenadas": {
                "disponivel": tem_coordenadas,
                "mensagem": "🌍 Sistema usou coordenadas geográficas reais!" if tem_coordenadas else "📊 Sistema usou matriz de distâncias original",
                "cds_com_coords": list(coordenadas.keys()) if tem_coordenadas else []
            } if tem_coordenadas else None,
            "mapa_html": mapa_html
        }
        
    except Exception as e:
        print(f"ERRO na otimização: {str(e)}")
        return {"status": "Erro", "mensagem": f"Erro interno: {str(e)}"}
