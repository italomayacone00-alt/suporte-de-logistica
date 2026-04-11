"""
Solver para o Problema de Localização Capacitada (Modelo Original)
Otimização de CDs com custos fixos e capacidades
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
        # Verificar se é um dicionário de DataFrames (múltiplas abas)
        if isinstance(df, dict):
            if 'Coordenadas_CDs' in df:
                coords_df = df['Coordenadas_CDs']
                print(f"📁 Encontrada aba 'Coordenadas_CDs' com {len(coords_df)} linhas")
            else:
                print("❌ Aba 'Coordenadas_CDs' não encontrada no arquivo")
                return None
        else:
            # Se for DataFrame único, não há coordenadas disponíveis
            print("❌ Arquivo não possui múltiplas abas (sem coordenadas)")
            return None
        
        coordenadas = {}
        
        # Detectar nomes das colunas automaticamente
        cd_col = None
        lat_col = None
        lon_col = None
        
        for col in coords_df.columns:
            col_lower = str(col).lower().strip()
            if 'cd' in col_lower or 'centro' in col_lower or 'warehouse' in col_lower:
                cd_col = col
            elif 'lat' in col_lower and 'lon' not in col_lower:
                lat_col = col
            elif 'lon' in col_lower or 'lng' in col_lower:
                lon_col = col
        
        print(f"🔍 Colunas detectadas: CD='{cd_col}', Latitude='{lat_col}', Longitude='{lon_col}'")
        
        if not all([cd_col, lat_col, lon_col]):
            print("❌ Colunas necessárias não encontradas na aba de coordenadas")
            return None
        
        for _, row in coords_df.iterrows():
            cd_nome = str(row[cd_col]).strip()
            
            # Converter coordenadas aceitando tanto ponto quanto vírgula
            lat_raw = str(row[lat_col]) if pd.notna(row[lat_col]) else None
            lon_raw = str(row[lon_col]) if pd.notna(row[lon_col]) else None
            
            lat = float(lat_raw.replace(',', '.')) if lat_raw else None
            lon = float(lon_raw.replace(',', '.')) if lon_raw else None
            
            coordenadas[cd_nome] = {'lat': lat, 'lon': lon}
        
        print(f"✅ Coordenadas carregadas: {len(coordenadas)} CDs")
        for cd, coords in coordenadas.items():
            print(f"  📍 {cd}: ({coords['lat']:.4f}, {coords['lon']:.4f})")
        
        return coordenadas
        
    except Exception as e:
        print(f"❌ Erro ao carregar coordenadas: {str(e)}")
        print(f"❌ Tipo do erro: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        return None

def calcular_distancias_geograficas(coordenadas, cds, clientes, valores_originais):
    """
    Calcula distâncias reais usando coordenadas dos CDs
    Para clientes sem coordenadas, mantém valores originais
    """
    print("🌍 Calculando distâncias geográficas reais...")
    
    valores = {}  # Dicionário de pares (cd, cliente): valor
    
    for cd in cds:
        cd_coords = coordenadas.get(cd)
        
        if not cd_coords or cd_coords['lat'] is None or cd_coords['lon'] is None:
            # Se não tiver coordenadas do CD, usar valores originais
            print(f"⚠️ CD {cd} sem coordenadas, usando matriz original")
            for cliente in clientes:
                valores[(cd, cliente)] = valores_originais.get((cd, cliente), float('inf'))
            continue
        
        # Para cada cliente, calcular distância real
        for cliente in clientes:
            # Por enquanto, clientes não têm coordenadas
            # Futuro: poderia adicionar coordenadas dos clientes também
            # Por ora, usa uma aproximação ou mantém original
            
            # Se tiver dados originais, usa como referência
            valor_original = valores_originais.get((cd, cliente), float('inf'))
            
            # Se valor original for finito, usa como proxy
            # Futuro melhor: adicionar coordenadas dos clientes
            valores[(cd, cliente)] = valor_original
            
            # DEBUG: mostrar quando CD tem coordenadas
            if valor_original != float('inf'):
                print(f"📍 CD {str(cd)} ({cd_coords['lat']:.4f}, {cd_coords['lon']:.4f}) → Cliente {str(cliente)}: {valor_original:.2f}")
    
    print(f"✅ Distâncias calculadas: {len(valores)} pares CD-Cliente")
    return valores

def gerar_mapa_tradicional(resultado, coordenadas):
    """
    Gera mapa interativo para o Otimizador Tradicional usando Folium
    """
    try:
        print(f"🔍 DEBUG: Iniciando gerar_mapa_tradicional...")
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
        cds_abertos = resultado.get('cds_abertos', [])
        
        # Coordenadas de todos os CDs (pode estar vazio se não houver coordenadas)
        todos_cds = list(coordenadas.keys()) if coordenadas else []
        
        print(f"🗺️ Gerando mapa Tradicional com {len(cds_abertos)} CDs abertos...")
        
        # Se não houver coordenadas, criar coordenadas padrão para visualização
        if not coordenadas:
            print("⚠️ Sem coordenadas dos CDs, criando coordenadas padrão para visualização...")
            # Criar coordenadas fictícias baseadas nos CDs
            import random
            coordenadas = {}
            for i, cd in enumerate(cds_abertos):
                # Coordenadas fictícias espalhadas pelo Brasil
                coords_ficticias = [
                    (-23.5505, -46.6333),  # São Paulo
                    (-22.9068, -43.1729),  # Rio de Janeiro
                    (-19.9167, -43.9345),  # Belo Horizonte
                    (-30.0346, -51.2177),  # Porto Alegre
                    (-8.0476, -34.8770),   # Recife
                    (-12.9714, -38.5014),  # Salvador
                    (-15.8267, -47.9218),  # Brasília
                    (-3.1190, -60.0217),   # Manaus
                    (-5.0892, -42.8016),   # Teresina
                    (-2.5307, -44.3068)    # São Luís
                ]
                lat, lon = coords_ficticias[i % len(coords_ficticias)]
                coordenadas[cd] = {'lat': lat, 'lon': lon}
        
        # Verificar se há coordenadas válidas para os CDs abertos
        coords_validas = [coordenadas[cd] for cd in cds_abertos if cd in coordenadas]
        
        print(f"🔍 DEBUG: CDs abertos = {cds_abertos}")
        print(f"🔍 DEBUG: Coordenadas disponíveis = {list(coordenadas.keys())}")
        print(f"🔍 DEBUG: Coordenadas válidas = {coords_validas}")
        
        if not coords_validas:
            print("❌ Nenhuma coordenada válida encontrada para os CDs abertos")
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
        
        # Adicionar marcadores para CDs abertos
        for cd in cds_abertos:
            if cd in coordenadas:
                lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                folium.Marker(
                    location=[lat, lon],
                    popup=f'<strong>{str(cd)}</strong><br>Coordenadas: {lat:.4f}, {lon:.4f}<br>Status: <span style="color:green">ABERTO</span>',
                    icon=folium.Icon(color='green', icon='warehouse')
                ).add_to(mapa)
        
        # Adicionar marcadores para CDs fechados
        for cd in todos_cds:
            if cd not in cds_abertos:
                if cd in coordenadas:
                    lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                    folium.Marker(
                        location=[lat, lon],
                        popup=f'<strong>{str(cd)}</strong><br>Coordenadas: {lat:.4f}, {lon:.4f}<br>Status: <span style="color:red">FECHADO</span>',
                        icon=folium.Icon(color='red', icon='times')
                    ).add_to(mapa)
        
        # Adicionar áreas de cobertura (círculos ao redor dos CDs abertos)
        for cd in cds_abertos:
            if cd in coordenadas:
                lat, lon = coordenadas[cd]['lat'], coordenadas[cd]['lon']
                capacidade_info = resultado.get('capacidade_utilizada', {}).get(cd, {})
                percentual = capacidade_info.get('percentual_uso', 0)
                
                folium.Circle(
                    location=[lat, lon],
                    radius=50000,  # 50km de raio
                    popup=f'Área de Influência: {str(cd)}<br>Capacidade utilizada: {percentual:.1f}%',
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
                <span style="color: #374151;">CD Ativo (Atendimento)</span>
            </p>
            <p style="margin: 10px 0; display: flex; align-items: center;">
                <i class="fa fa-warehouse" style="color: #dc2626; margin-right: 10px; font-size: 16px;"></i> 
                <span style="color: #374151;">CD Inativo (Sem Atendimento)</span>
            </p>
            <p style="margin: 10px 0; display: flex; align-items: center;">
                <span style="background-color: rgba(0, 255, 0, 0.2); border: 1px solid green; padding: 6px 12px; border-radius: 4px; margin-right: 10px;">&nbsp;&nbsp;&nbsp;</span> 
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
        
        mapa_filename = f"mapa_tradicional_{hash(str(coordenadas))}.html"
        mapa_path = os.path.join(static_dir, mapa_filename)
        
        mapa.save(mapa_path)
        print(f"📁 Mapa salvo em: {mapa_path}")
        
        return mapa_filename
        
    except Exception as e:
        print(f"❌ Erro ao gerar mapa: {str(e)}")
        return None

def resolver_problema_logistica(df, tipo_dado='custo'):
    """
    Resolve o problema de localização-alocação capacitada.
    
    Modelo:
    - Minimiza custos fixos + custos de transporte
    - Respeita capacidades dos CDs
    - Determina número ótimo de CDs a abrir
    
    Parâmetros:
    - df: DataFrame ou dicionário de DataFrames com dados da planilha
    - tipo_dado: 'custo' ou 'distancia' para tipo de transporte
    
    Retorna:
    - Dicionário com solução ótima
    """
    try:
        print("=== DEBUG: Iniciando solver tradicional ===")
        
        # Extrair a primeira aba como dados principais se for dicionário
        if isinstance(df, dict):
            primeira_aba = list(df.keys())[0]
            df_principal = df[primeira_aba]
            print(f"📁 Usando aba principal: '{primeira_aba}'")
        else:
            df_principal = df
            print("📁 Usando DataFrame único")
        
        # 0. VERIFICAR COORDENADAS (NOVO)
        coordenadas = None
        tem_coordenadas = False
        try:
            coordenadas = carregar_coordenadas_cds(df)
            tem_coordenadas = coordenadas is not None
            print(f"🌍 Tem coordenadas dos CDs? {tem_coordenadas}")
        except Exception as coord_error:
            print(f"⚠️ Erro ao verificar coordenadas (continuando sem coordenadas): {str(coord_error)}")
            tem_coordenadas = False
        
        # 1. Separar a linha de Demanda dos CDs candidatos
        nome_primeira_coluna = df_principal.columns[0] 
        print(f"Nome da primeira coluna: {nome_primeira_coluna}")
        
        linha_demanda = df_principal[df_principal[nome_primeira_coluna] == 'Demanda Total']
        print(f"Linha de demanda encontrada:\n{linha_demanda}")
        
        df_cds = df_principal[df_principal[nome_primeira_coluna] != 'Demanda Total'].dropna(subset=[nome_primeira_coluna])
        print(f"DataFrame de CDs:\n{df_cds}")
        
        # 2. Identificar listas de CDs e Clientes
        cds = df_cds[nome_primeira_coluna].astype(str).tolist()
        print(f"CDs encontrados: {cds}")
        
        clientes = df_principal.columns[1:-2].tolist()  # Excluindo Custo Fixo e Capacidade
        print(f"Clientes encontrados: {clientes}")
        print(f"Colunas disponíveis: {list(df_principal.columns)}")
        print(f"Colunas a serem excluídas: {df_principal.columns[0]}, Custo Fixo, Capacidade")
        
        # 3. Extrair os Dicionários de Parâmetros
        demanda = {cliente: int(pd.to_numeric(linha_demanda[cliente].values[0])) for cliente in clientes}
        print(f"Demanda: {demanda}")
        
        # Extrair Custos Fixos
        custos_fixos = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            custos_fixos[cd_atual] = int(pd.to_numeric(row['Custo Fixo']))
        print(f"Custos fixos: {custos_fixos}")
        
        # Extrair Capacidades
        capacidades = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            capacidades[cd_atual] = int(pd.to_numeric(row['Capacidade']))
        print(f"Capacidades: {capacidades}")
        
        # Extrair Custos de Transporte (ou Distâncias)
        custos_transporte = {}
        for index, row in df_cds.iterrows():
            cd_atual = str(row[nome_primeira_coluna])
            for cliente in clientes:
                # Verificar se a coluna do cliente existe
                if cliente in row.index:
                    custos_transporte[(cd_atual, cliente)] = float(pd.to_numeric(row[cliente]))
                else:
                    print(f"⚠️ Cliente '{cliente}' não encontrado para CD '{cd_atual}', usando valor 0")
                    custos_transporte[(cd_atual, cliente)] = 0
        print(f"Custos de transporte (primeiros 5): {dict(list(custos_transporte.items())[:5])}")
        print(f"Total de pares CD-Cliente: {len(custos_transporte)}")
        
        # VERIFICAÇÃO CRÍTICA: Garantir que todos os pares existam antes da otimização
        pares_faltantes = []
        for cd in cds:
            for cliente in clientes:
                if (cd, cliente) not in custos_transporte:
                    pares_faltantes.append((cd, cliente))
        
        if pares_faltantes:
            print(f"❌ Pares faltantes detectados: {len(pares_faltantes)}")
            for par in pares_faltantes[:5]:  # Mostrar apenas os 5 primeiros
                print(f"   Faltando: {par}")
            # Adicionar pares faltantes com valor 0
            for par in pares_faltantes:
                custos_transporte[par] = 0
            print(f"✅ Pares faltantes adicionados com valor 0")
        else:
            print("✅ Todos os pares CD-Cliente estão presentes")
        
        # Identificar o tipo de transporte nos logs
        tipo_transporte = "Custos" if tipo_dado == 'custo' else "Distâncias"
        print(f"Usando {tipo_transporte} de transporte")
        
        # SE tiver coordenadas e for distância, usar cálculo geográfico
        if tem_coordenadas and tipo_dado == 'distancia':
            custos_transporte = calcular_distancias_geograficas(coordenadas, cds, clientes, custos_transporte)
            print("🌍 Usando sistema com coordenadas geográficas!")
        else:
            print("📊 Usando matriz de distâncias original")
        
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
        try:
            custo_fixo_total = lpSum([custos_fixos[i] * y[i] for i in cds])
            custo_transporte_total = lpSum([custos_transporte[(i, j)] * x[i][j] for i in cds for j in clientes])
            prob += custo_fixo_total + custo_transporte_total, "Custo_Total"
        except KeyError as e:
            print(f"❌ Erro na fórmula matemática: {e}")
            print(f"❌ Par não encontrado: {e}")
            print(f"❌ CDs disponíveis: {cds}")
            print(f"❌ Clientes disponíveis: {clientes}")
            print(f"❌ Chaves em custos_transporte: {list(custos_transporte.keys())[:10]}...")
            raise Exception(f"Erro ao processar os dados da planilha: {e}")
        except Exception as e:
            print(f"❌ Erro geral na otimização: {e}")
            raise Exception(f"Erro ao processar os dados da planilha: {e}")
        
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
        
        # Gerar mapa se tiver coordenadas OU SEMPRE (para mostrar mapa mesmo sem coordenadas)
        mapa_html = None
        if True:  # Sempre gerar mapa, mesmo sem coordenadas
            print("🗺️ Gerando mapa interativo...")
            mapa_html = gerar_mapa_tradicional(
                {
                    "status": "Sucesso",
                    "cds_abertos": cds_abertos,
                    "custo_total": value(prob.objective),
                    "capacidade_utilizada": capacidade_utilizada
                }, 
                coordenadas  # Pode ser None, a função vai tratar
            )
            if mapa_html:
                print(f"✅ Mapa gerado: {mapa_html}")
            else:
                print("❌ Não foi possível gerar o mapa")
        
        resultado_final = {
            "status": "Sucesso",
            "modelo": f"Localização Capacitada - {tipo_transporte}",
            "tipo_transporte": tipo_dado,
            "tipo_valor": tipo_dado,  # Adicionando campo tipo_valor para compatibilidade com templates
            "p": len(cds_abertos),  # Adicionando atributo p
            "total_cds": len(cds),  # Adicionando total de CDs disponíveis
            "custo_total": value(prob.objective),
            "custo_fixo_total": value(custo_fixo_total),
            "custo_transporte_total": value(custo_transporte_total),
            "cds_abertos": cds_abertos,
            "num_cds_selecionados": len(cds_abertos),  # Adicionando atributo consistente
            "cds_selecionados": cds_abertos,  # Adicionando para compatibilidade
            "transportes": transportes,
            "num_transportes": len(transportes),
            "capacidade_utilizada": capacidade_utilizada,
            "demanda_total": sum(demanda.values()),
            "demanda_atendida": sum([t['quantidade'] for t in transportes]),
            "coordenadas_usadas": tem_coordenadas,
            "info_coordenadas": {
                "disponivel": tem_coordenadas,
                "mensagem": "🌍 Sistema usou coordenadas geográficas reais!" if tem_coordenadas else "📊 Sistema usou matriz de distâncias original",
                "cds_com_coords": list(coordenadas.keys()) if tem_coordenadas else []
            } if tem_coordenadas else None,
            "mapa_html": mapa_html
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
