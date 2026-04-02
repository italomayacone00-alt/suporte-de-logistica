# 🚛 Sistema de Otimização Logística

Sistema web completo para otimização de logística utilizando modelos matemáticos avançados de p-Medianas, p-Centros, Máxima Cobertura e modelo tradicional.

## 📋 Funcionalidades

### 🎯 Modelos de Otimização
- **p-Medianas**: Minimização de custos/distâncias totais ponderadas pela demanda
- **p-Centros**: Minimização da máxima distância/custo de atendimento
- **Máxima Cobertura**: Maximização do número de clientes atendidos dentro de um raio
- **Modelo Tradicional**: Otimização clássica de distribuição logística

### 🗺️ Visualização Geográfica
- Mapas interativos com Leaflet.js
- Áreas de influência e cobertura visual
- Identificação de CDs ativos e inativos
- Estatísticas geográficas em tempo real

### 📊 Análises e Relatórios
- Gráficos interativos com Chart.js
- Análises estatísticas avançadas
- Recomendações estratégicas automáticas
- Exportação para PDF e Excel

### 💾 Gerenciamento de Projetos
- Salvar e carregar análises
- Dashboard com histórico de projetos
- Comparação entre diferentes cenários

## 🛠️ Tecnologias Utilizadas

### Backend
- **Python 3.9+**
- **Flask** - Framework web
- **SQLite** - Banco de dados
- **PuLP** - Otimização matemática
- **Pandas** - Análise de dados

### Frontend
- **HTML5** com Jinja2 templates
- **CSS3** com design responsivo
- **JavaScript** vanilla
- **Chart.js** - Visualização de dados
- **Leaflet.js** - Mapas interativos

### UI/UX
- **Material Symbols** - Ícones modernos
- **Design responsivo** para mobile/tablet/desktop
- **Interface intuitiva** com feedback visual

## 📦 Instalação

### Pré-requisitos
- Python 3.9 ou superior
- pip (gerenciador de pacotes Python)

### Passos para instalação

1. **Clone o repositório:**
```bash
git clone https://github.com/italomayacone00/suporte-de-logistica.git
cd suporte-de-logistica
```

2. **Crie ambiente virtual:**
```bash
python -m venv venv
# Windows
venv\Scripts\activate
# Linux/Mac
source venv/bin/activate
```

3. **Instale dependências:**
```bash
pip install -r requirements.txt
```

4. **Inicie o aplicativo:**
```bash
python app.py
```

5. **Acesse o sistema:**
```
http://127.0.0.1:5000
```

## 📁 Estrutura do Projeto

```
suporte-de-logistica/
├── app.py                 # Aplicação Flask principal
├── database.py           # Configuração do banco de dados
├── requirements.txt      # Dependências Python
├── 
├── solver_*.py          # Modelos de otimização
│   ├── solver_maxcobertura.py
│   ├── solver_pcentros.py
│   ├── solver_pmedianas.py
│   └── solver_tradicional.py
├── 
├── static/              # Arquivos estáticos
│   ├── css/
│   │   └── style.css    # Estilos principais
│   ├── js/
│   │   └── script.js    # Scripts frontend
│   └── mapas/           # Mapas gerados
├── 
├── templates/           # Templates HTML
│   ├── base.html
│   ├── index.html
│   ├── *_form.html      # Formulários de entrada
│   └── resultado_*.html # Páginas de resultado
├── 
├── uploads/             # Upload de arquivos
└── projetos_logistica.db # Banco de dados SQLite
```

## 🚀 Como Usar

### 1. **Importar Dados**
- Acesse a página inicial
- Escolha o modelo de otimização desejado
- Faça upload da planilha com dados (formato .xlsx)

### 2. **Configurar Parâmetros**
- Defina o número de CDs a serem selecionados (parâmetro p)
- Ajuste configurações específicas de cada modelo
- Configure tipo de análise (distância vs custo)

### 3. **Executar Otimização**
- Clique em "Executar Análise"
- Aguarde processamento do modelo matemático
- Visualize resultados em tempo real

### 4. **Analisar Resultados**
- Mapa interativo com localização ótima
- Estatísticas detalhadas de desempenho
- Gráficos e análises comparativas
- Recomendações estratégicas

### 5. **Exportar e Salvar**
- Exporte relatórios em PDF
- Baixe planilhas com resultados detalhados
- Salve projetos para acesso futuro

## 📊 Formato de Dados

### Planilha de Entrada (.xlsx)
A planilha deve conter:

| Coluna | Descrição | Exemplo |
|--------|-----------|---------|
| CD | Identificação do Centro de Distribuição | CD1, CD2, CD3 |
| Cliente | Nome/Código do Cliente | Cliente1, Cliente2 |
| Distância | Distância (km) ou Custo (R$) | 150, 75.5 |
| Demanda | Quantidade demandada | 100, 250 |

### Coordenadas Geográficas (Opcional)
Para visualização em mapas, inclua:
- `CD_Lat`, `CD_Lng` - Coordenadas dos CDs
- `Cliente_Lat`, `Cliente_Lng` - Coordenadas dos clientes

## 🎯 Modelos Matemáticos

### p-Medianas
**Objetivo:** Minimizar custo total ponderado
```
Minimizar Σ (distância_ij × quantidade_ij)
Sujeito a: Σ Y_i = p (exatamente p CDs)
```

### p-Centros
**Objetivo:** Minimizar máxima distância de atendimento
```
Minimizar Max(distância_ij × X_ij)
Sujeito a: Σ Y_i = p (exatamente p CDs)
```

### Máxima Cobertura
**Objetivo:** Maximizar clientes dentro do raio
```
Maximizar Σ (peso_i × Z_i)
Sujeito a: distância_ij ≤ raio_cobertura
```

## 🔧 Configuração

### Variáveis de Ambiente
```bash
FLASK_ENV=development          # Modo desenvolvimento
FLASK_DEBUG=True              # Debug ativado
DATABASE_URL=sqlite:///projetos_logistica.db
```

### Configurações do Aplicativo
- `UPLOAD_FOLDER`: Diretório para uploads
- `MAX_CONTENT_LENGTH`: Tamanho máximo de arquivos
- `SECRET_KEY`: Chave de segurança Flask

## 🐛 Troubleshooting

### Problemas Comuns

**Erro: "Banco de dados não encontrado"**
```bash
python database.py  # Recriar banco de dados
```

**Erro: "Planilha inválida"**
- Verifique formato .xlsx
- Confirme colunas obrigatórias
- Remova células mescladas

**Erro: "Mapa não carrega"**
- Verifique coordenadas na planilha
- Confirme conexão com internet
- Limpe cache do navegador

### Logs e Debug
```bash
# Ver logs do aplicativo
python app.py --debug

# Verificar dependências
pip freeze
```

## 🤝 Contribuição

### Como Contribuir
1. Fork o repositório
2. Crie branch (`git checkout -b feature/nova-funcionalidade`)
3. Commit (`git commit -m 'feat: adicionar nova funcionalidade'`)
4. Push (`git push origin feature/nova-funcionalidade`)
5. Abra Pull Request

### Padrões de Código
- Python: PEP 8
- JavaScript: ES6+
- CSS: BEM methodology
- Commits: Conventional Commits

## 📄 Licença

Este projeto está licenciado sob a MIT License - veja o arquivo [LICENSE](LICENSE) para detalhes.

## 👥 Autores

- **Italo Mayacone** - *Desenvolvimento inicial* - [italomayacone00](https://github.com/italomayacone00)

## 🙏 Agradecimentos

- **PuLP** - Biblioteca de otimização matemática
- **Flask** - Framework web Python
- **Chart.js** - Visualização de dados
- **Leaflet.js** - Biblioteca de mapas
- **Material Symbols** - Conjunto de ícones

## 📞 Suporte

- **Issues**: [GitHub Issues](https://github.com/italomayacone00/suporte-de-logistica/issues)
- **Email**: italo.mayacone@example.com
- **Documentação**: [Wiki do Projeto](https://github.com/italomayacone00/suporte-de-logistica/wiki)

---

⭐ **Se este projeto ajudou você, por favor deixe uma estrela!**
