# 🧮 Athena Office - Sistema de Transporte de Dados (Dashboard Integrador)

## 📖 Visão Geral

Este projeto foi desenvolvido em **Python** com o objetivo de automatizar o **transporte e consolidação de dados financeiros** (planilhas de despesas) de várias filiais para dentro de um **Dashboard final** centralizado.

> ⚠️ **Importante:** O arquivo `DASHBOARDFINAL.xlsx` **deve existir previamente** dentro da pasta selecionada, pois o sistema apenas atualiza seus dados — não o cria.

O sistema oferece uma **interface gráfica amigável** construída com `customtkinter`, e possui funcionalidades automáticas de:
- Instalação e verificação de dependências;
- Leitura resiliente de arquivos `.xls` e `.xlsx`;
- Mapeamento e normalização de categorias de despesas;
- Backup automático do dashboard antes da atualização;
- Log detalhado de processamento e validação de dados.

---

## 🧰 Requisitos

### 📦 Dependências Python

As dependências são automaticamente verificadas e instaladas pelo próprio script no momento da execução.  
Entretanto, é possível instalá-las manualmente com:

```bash
pip install -r requirements.txt
```

### Lista de Bibliotecas
- `pandas`
- `openpyxl`
- `xlrd`
- `customtkinter`
- `pillow`

---

## 🏗️ Estrutura do Projeto

```bash
.
├── main.py                      # Script principal (interface + lógica)
├── DASHBOARDFINAL.xlsx          # Dashboard final (precisa existir previamente)
├── JoãoPessoa.xlsx              # Exemplo de planilha de cidade
├── SãoPaulo.xlsx
├── ...
└── requirements.txt             # Dependências opcionais (para empacotamento)
```

---

## ⚙️ Funcionalidades Principais

### 🔍 1. Detecção e Instalação de Dependências
A função `setup_environment()` garante que todas as bibliotecas necessárias estejam instaladas.  
Caso o script esteja empacotado (`.exe`), ele apenas alerta as ausentes.

### 🧩 2. Normalização de Dados
Funções utilitárias como:
- `strip_accents(s)` — remove acentos e caracteres especiais;
- `norm_text(s)` — uniformiza textos (minúsculas, sem espaços extras);
- `to_float(v)` — converte valores monetários em `float`, aceitando formatos brasileiros.

Essas funções garantem que os dados de diferentes planilhas possam ser comparados corretamente.

### 🗂️ 3. Leitura Inteligente de Arquivos Excel
A função `read_excel_any(path)` detecta o tipo de planilha:
- `.xlsx` → lida com `openpyxl`
- `.xls` → lida com `xlrd`
- fallback inteligente para casos ambíguos.

Isso permite trabalhar com planilhas antigas ou exportadas de sistemas diversos.

### 🧭 4. Mapeamento de Categorias
O script contém um dicionário `RAW_CATEGORY_MAPPING` que relaciona as **categorias brutas** (presentes nas planilhas das cidades) com as **categorias padronizadas** usadas no Dashboard.

Exemplo:
```python
'DESPESAS ADMINISTRATIVAS : Energia Elétrica' → 'Energia Elétrica'
```

Esse mapeamento é automaticamente normalizado e utilizado para vincular os dados corretos.

### 💻 5. Interface Gráfica (GUI)

A interface foi desenvolvida com `customtkinter`, oferecendo:
- Botão **📁 Selecionar Pasta**
- Botão **🚀 Processar Dados**
- Botão **🧹 Limpar Dashboard**
- Botão **❓ Ajuda**
- Barra de progresso e área de logs
- Indicadores de quantidade de **Cidades, Categorias e Atualizações**

A janela principal possui um layout moderno, com tema claro e elementos responsivos.

---

## 🧮 6. Lógica de Processamento

Quando o usuário clica em **“Processar Dados”**, o fluxo é:

1. **Backup Automático** do arquivo `DASHBOARDFINAL.xlsx` (timestamped).
2. **Carregamento** do dashboard via `openpyxl`.
3. **Leitura** de todas as planilhas de cidades (`*.xls` e `*.xlsx`).
4. **Extração de despesas** verticais e normalizadas.
5. **Atualização** das abas do dashboard correspondentes às cidades.
6. **Salvamento** do dashboard atualizado.
7. **Atualização** das estatísticas e logs de processamento.

### 🔁 Mecanismo de Matching de Categorias
A busca é feita de forma tolerante a diferenças de formatação e grafia.
Usa-se `SequenceMatcher` (do módulo `difflib`) para comparar similaridade entre textos, com limiar mínimo de **0.8**.

Exemplo:
```python
similarity_score("Tarifas Bancárias TED", "Tarifas Bancarias - TED") ≈ 0.95
```

Além disso, há tratamento especial para categorias sensíveis como **tarifas bancárias (PIX, TED, Boletos, Cartão)**.

---

## 💾 Backups e Logs

- A cada execução, é criado automaticamente um backup:
  ```
  DASHBOARDFINAL_backup_YYYYMMDD_HHMMSS.xlsx
  ```

- Todas as ações (instalações, carregamentos, atualizações e erros) são registradas na área de logs da interface, garantindo transparência durante o processo.

---

## 🧹 Função Extra: Limpar Dashboard

O botão **🧹 Limpar Dashboard** (não mostrado integralmente no código acima) é responsável por redefinir os valores da planilha para um estado inicial, útil antes de um novo processamento em lote.

---

## 🪟 Interface Gráfica - Exemplo Visual

```
+-------------------------------------------------------------+
| Athena Office - Transporte de Dados                         |
|-------------------------------------------------------------|
| [🚀 Processar Dados] [📁 Selecionar Pasta] [🧹 Limpar]       |
|-------------------------------------------------------------|
| LOG:                                                        |
|  ✅ pandas disponível                                       |
|  🚀 Iniciando processamento...                              |
|  📋 JoãoPessoa.xlsx -> 14 categorias                        |
|  💾 Dashboard salvo com sucesso!                            |
|-------------------------------------------------------------|
| Cidades: 3 | Atualizadas: 27 | Categorias: 48               |
+-------------------------------------------------------------+
```

---

## 🧱 Estrutura das Funções Principais

| Função | Responsabilidade |
|--------|------------------|
| `setup_environment()` | Verifica dependências e ambiente |
| `norm_text(s)` | Normaliza textos para matching |
| `read_excel_any(path)` | Lê arquivos Excel com tolerância a formatos |
| `DashboardApp` | Classe principal da GUI |
| `process_data()` | Faz o transporte de dados entre planilhas |
| `extract_expenses_vertical()` | Extrai categorias e valores de planilhas |
| `update_dashboard_city_sheet()` | Atualiza as células corretas no Dashboard |
| `find_city_sheet()` | Localiza a aba correspondente à cidade |

---

## 🧩 Compatibilidade

- **Sistemas operacionais:** Windows, macOS e Linux
- **Versão mínima recomendada do Python:** 3.8
- **Formatos suportados:** `.xls` e `.xlsx`

---

## 🚀 Execução

1. Certifique-se de ter o Python instalado.
2. Coloque o arquivo `DASHBOARDFINAL.xlsx` na pasta desejada.
3. Adicione as planilhas das cidades no mesmo diretório.
4. Execute o script:

```bash
python main.py
```

5. Use a interface gráfica para selecionar a pasta e clicar em **“Processar Dados”**.

---

## 🧑‍💻 Autor

**Desenvolvido por:** Gabriel Cunha Ramos  
**Organização:** Athena Office  
**Linguagem:** Python  
**Interface:** CustomTkinter

---

## 🏁 Licença

Este projeto é de uso interno e está licenciado sob a política de software interno da **Athena Office**.  
Reprodução, redistribuição ou modificação externa requerem autorização prévia.

---
