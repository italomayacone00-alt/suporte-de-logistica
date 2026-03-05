# SAD Logística - Sistema de Apoio à Decisão

Um sistema web completo para otimização de malha logística utilizando Programação Linear Inteira com o solver PuLP.

## 🚀 Funcionalidades

- **Geração Automática de Templates**: Cria planilhas Excel pré-formatadas com base no número de CDs e clientes
- **Otimização Matemática**: Utiliza Programação Linear Inteira para encontrar a solução ótima
- **Interface Moderna**: Design responsivo e intuitivo com animações suaves
- **Análise Completa**: Relatórios detalhados com custos, rotas e indicadores de performance
- **Exportação de Resultados**: Possibilidade de exportar análises em diversos formatos

## 📁 Estrutura do Projeto

```
meu_sistema_logistico/
│
├── app.py                 # Rotas Flask e lógica web
├── solver.py              # Algoritmo de otimização PuLP
├── requirements.txt       # Dependências Python
├── README.md             # Documentação do projeto
│
├── static/                # Arquivos estáticos
│   ├── css/
│   │   └── style.css      # Estilos modernos e responsivos
│   └── js/
│       └── script.js      # Interatividade e validações
│
└── templates/             # Templates HTML
    ├── base.html          # Layout base com header/footer
    ├── index.html         # Página principal com formulários
    └── resultado.html     # Página de resultados da otimização
```

## 🛠️ Instalação e Configuração

### Pré-requisitos
- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

### Passos para instalação

1. **Clone o repositório** (se estiver usando Git):
   ```bash
   git clone <url-do-repositorio>
   cd meu_sistema_logistico
   ```

2. **Crie um ambiente virtual** (recomendado):
   ```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Linux/Mac
   source venv/bin/activate
   ```

3. **Instale as dependências**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Execute a aplicação**:
   ```bash
   python app.py
   ```

5. **Acesse o sistema**:
   Abra seu navegador e acesse `http://localhost:5000`

## 📊 Como Usar

### Passo 1: Gerar Template
1. Informe o número de CDs candidatos (1-50)
2. Informe o número de clientes/destinos (1-200)
3. Clique em "Gerar Template Excel"
4. Baixe o arquivo Excel gerado

### Passo 2: Preencher Dados
O arquivo Excel contém 3 abas:

**Aba "CDs"**:
- CD: Identificação do centro de distribuição
- Capacidade: Capacidade máxima de atendimento
- Custo_Fixo: Custo mensal fixo para operação

**Aba "Clientes"**:
- Cliente: Identificação do ponto de destino
- Demanda: Quantidade demandada mensalmente

**Aba "Custos_Transporte"**:
- CD: Centro de distribuição de origem
- Cliente: Cliente de destino
- Custo_Transporte: Custo unitário por unidade transportada

### Passo 3: Otimizar
1. Preencha todos os dados na planilha
2. Volte ao sistema web
3. Faça upload do arquivo preenchido
4. Clique em "Otimizar e Resolver"
5. Analise os resultados detalhados

## 🧮 Modelo Matemático

O sistema resolve um problema de localização de facilidades clássico:

### Variáveis de Decisão
- `y[i]`: Binária - 1 se o CD i for aberto, 0 caso contrário
- `x[i,j]`: Contínua - quantidade transportada do CD i para o cliente j

### Função Objetivo
Minimizar: `Σ(Custo_Fixo[i] × y[i]) + Σ(Custo_Transporte[i,j] × x[i,j])`

### Restrições
1. **Atendimento da Demanda**: `Σ(x[i,j]) = Demanda[j]` para cada cliente j
2. **Capacidade dos CDs**: `Σ(x[i,j]) ≤ Capacidade[i] × y[i]` para cada CD i
3. **Integridade**: `y[i] ∈ {0,1}` e `x[i,j] ≥ 0`

## 🎯 Recursos Técnicos

### Backend
- **Flask**: Framework web Python
- **PuLP**: Biblioteca de otimização matemática
- **Pandas**: Manipulação de dados
- **OpenPyXL**: Leitura/escrita de arquivos Excel

### Frontend
- **HTML5 Semântico**: Estrutura acessível
- **CSS3 Moderno**: Design responsivo com Grid e Flexbox
- **JavaScript Vanilla**: Interatividade sem dependências
- **Animações CSS**: Transições suaves e feedback visual

### Features Implementadas
- ✅ Validação de entrada em tempo real
- ✅ Drag & drop para upload de arquivos
- ✅ Animações de carregamento e transições
- ✅ Design responsivo para mobile
- ✅ Acessibilidade (ARIA labels, navegação por teclado)
- ✅ Notificações não intrusivas
- ✅ Exportação de resultados
- ✅ Impressão otimizada

## 🔧 Configurações Avançadas

### Personalização do Solver
No arquivo `solver.py`, você pode modificar:

- **Limites de CDs**: Descomente as linhas 85-86 para limitar número mínimo/máximo
- **Solver alternativo**: Configure solvers como CBC, Gurobi ou CPLEX
- **Parâmetros de tolerância**: Ajuste precisão e limites de tempo

### Customização Visual
No arquivo `static/css/style.css`:
- **Cores**: Modifique as variáveis de cor no topo do arquivo
- **Tipografia**: Altere fontes e tamanhos
- **Layout**: Ajuste grids e responsividade

## 📈 Indicadores de Performance

O sistema calcula automaticamente:
- **Custo Total**: Soma de custos fixos e variáveis
- **Taxa de Utilização**: Percentual de capacidade utilizada
- **Custo por Unidade**: Eficiência operacional
- **Número Ótimo de CDs**: Balanceamento entre fixo e variável

## 🐛 Solução de Problemas

### Erros Comuns
1. **"Demanda maior que capacidade"**: Aumente capacidade ou reduza demanda
2. **"Arquivo inválido"**: Verifique se é .xlsx e não está corrompido
3. **"Solver não encontrou solução"**: Verifique consistência dos dados

### Logs e Debug
Ative modo debug no Flask:
```python
app.run(debug=True)
```

## 🤝 Contribuições

Contribuições são bem-vindas! Áreas para melhoria:
- Novos algoritmos de otimização
- Visualizações gráficas
- Integração com APIs de logística
- Multi-usuário e banco de dados

## 📄 Licença

Este projeto está sob licença MIT. Sinta-se livre para usar e modificar.

## 📞 Suporte

Para dúvidas ou suporte:
1. Verifique a documentação acima
2. Consulte os logs de erro
3. Teste com dados de exemplo pequenos

---

**Desenvolvido com ❤️ para otimização logística**
