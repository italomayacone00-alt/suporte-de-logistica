"""
Modelo de banco de dados para salvar projetos de otimização logística
"""
import sqlite3
import json
from datetime import datetime
import os

# Caminho do banco de dados
DB_PATH = os.path.join(os.path.dirname(__file__), 'projetos_logistica.db')

def init_database():
    """Inicializa o banco de dados com as tabelas necessárias"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # Tabela de projetos
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS projetos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT NOT NULL,
            tipo_analise TEXT NOT NULL,  -- 'p_centros', 'p_medianas', 'max_cobertura', 'tradicional'
            data_criacao TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            parametros TEXT,  -- JSON com parâmetros de entrada
            status TEXT DEFAULT 'ativo'  -- 'ativo', 'arquivado', 'excluido'
        )
    ''')
    
    # Tabela de resultados
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS resultados (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            projeto_id INTEGER NOT NULL,
            dados TEXT NOT NULL,  -- JSON com todos os resultados da otimização
            mapa_html TEXT,  -- nome do arquivo do mapa se existir
            criado_em TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (projeto_id) REFERENCES projetos (id)
        )
    ''')
    
    conn.commit()
    conn.close()
    print("✅ Banco de dados inicializado com sucesso!")

def salvar_projeto(nome, tipo_analise, parametros, resultados, mapa_html=None):
    """Salva um novo projeto com seus resultados"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        # Inserir projeto
        cursor.execute('''
            INSERT INTO projetos (nome, tipo_analise, parametros)
            VALUES (?, ?, ?)
        ''', (nome, tipo_analise, json.dumps(parametros)))
        
        projeto_id = cursor.lastrowid
        
        # Inserir resultados
        cursor.execute('''
            INSERT INTO resultados (projeto_id, dados, mapa_html)
            VALUES (?, ?, ?)
        ''', (projeto_id, json.dumps(resultados), mapa_html))
        
        conn.commit()
        return projeto_id
        
    except Exception as e:
        conn.rollback()
        return None
    finally:
        conn.close()

def listar_projetos(tipo_analise=None):
    """Lista todos os projetos salvos"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        if tipo_analise:
            cursor.execute('''
                SELECT id, nome, tipo_analise, data_criacao, status
                FROM projetos 
                WHERE tipo_analise = ? AND status = 'ativo'
                ORDER BY data_criacao DESC
            ''', (tipo_analise,))
        else:
            cursor.execute('''
                SELECT id, nome, tipo_analise, data_criacao, status
                FROM projetos 
                WHERE status = 'ativo'
                ORDER BY data_criacao DESC
            ''')
        
        projetos = []
        for row in cursor.fetchall():
            projetos.append({
                'id': row[0],
                'nome': row[1],
                'tipo_analise': row[2],
                'data_criacao': row[3],
                'status': row[4]
            })
        
        return projetos
        
    except Exception as e:
        print(f"❌ Erro ao listar projetos: {str(e)}")
        return []
    finally:
        conn.close()

def carregar_resultados_projeto(projeto_id):
    """Carrega os resultados de um projeto específico"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
            SELECT p.id, p.nome, p.tipo_analise, p.data_criacao, 
                   p.parametros, r.dados, r.mapa_html
            FROM projetos p
            LEFT JOIN resultados r ON p.id = r.projeto_id
            WHERE p.id = ? AND p.status = 'ativo'
        ''', (projeto_id,))
        
        row = cursor.fetchone()
        if row:
            return {
                'id': row[0],
                'nome': row[1],
                'tipo_analise': row[2],
                'data_criacao': row[3],
                'parametros': json.loads(row[4]) if row[4] else {},
                'resultados': json.loads(row[5]) if row[5] else {},
                'mapa_html': row[6]
            }
        return None
        
    except Exception as e:
        print(f"❌ Erro ao carregar resultados: {str(e)}")
        return None
    finally:
        conn.close()

def excluir_projeto(projeto_id):
    """Exclui (marca como excluído) um projeto"""
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    try:
        cursor.execute('''
            UPDATE projetos 
            SET status = 'excluido' 
            WHERE id = ?
        ''', (projeto_id,))
        
        conn.commit()
        print(f"✅ Projeto {projeto_id} excluído com sucesso!")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao excluir projeto: {str(e)}")
        return False
    finally:
        conn.close()

# Inicializar o banco de dados quando o módulo for importado
if __name__ == "__main__":
    init_database()
