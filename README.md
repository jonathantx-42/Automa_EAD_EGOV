# Central de Automações EAD - EGOV (Escola de Governo do DF)

Aplicativo Desktop em Python com interface moderna (**CustomTkinter**) que unifica todo o ciclo de gestão de cursos EAD na Escola de Governo do DF.

---

## 📁 Estrutura Oficial de Pastas e Fluxo de Arquivos

O fluxo de trabalho organiza todos os arquivos locais em 4 pastas padronizadas e, na 5ª etapa, realiza a sincronização com as pastas oficiais do servidor de rede:

```text
┌──────────────────────────────────────────┐
│ 1. Arquivos_Originais_Sistema/           │ <- CSVs originais baixados do Portal GDF
└────────────────────┬─────────────────────┘
                     │
                     ▼ (Tratamento)
┌──────────────────────────────────────────┐
│ 2. Arquivos_CARREGAR/                    │ <- CSVs tratados prontos para subir no Moodle
└────────────────────┬─────────────────────┘
                     │
┌────────────────────┴─────────────────────┐
│ 3. Arquivos_Originais_Moodle/            │ <- Planilhas brutas de notas (.xlsx) do Moodle
└────────────────────┬─────────────────────┘
                     │
                     ▼ (Avaliação & Base Legal)
┌──────────────────────────────────────────┐
│ 4. Arquivos_Finais_SEDOC/                │ <- Planilhas finais oficiais (..._SEDOC.xlsx)
└────────────────────┬─────────────────────┘
                     │
                     ▼ (Cópia Organizada com Proteção de Conflitos)
┌──────────────────────────────────────────────────────────────────────────────────┐
│ 5. Servidor EGOV (Disco D:)                                                      │
│    D:\SEGEA\RECENTES\EGOV\CEAD\Cursos Autoinstrucionais\# 2026\                  │
│    ├── [SIGLA]\Turma XX.YYYY\                                                    │
│    └── #PROAMIS\[SUBSIGLA]\Turma XX.YYYY\                                        │
└──────────────────────────────────────────────────────────────────────────────────┘
```

---

## 📦 Como Instalar em um Novo Computador

### Passo 1: Instalar o Python
- Baixe o instalador do Python (versão 3.10 ou superior) em [python.org](https://www.python.org/downloads/).
- **MUITO IMPORTANTE:** Na primeira tela da instalação, marque a caixinha **"Add Python to PATH"** (Adicionar Python ao PATH) antes de clicar em instalar.

### Passo 2: Instalar as Dependências (1 Clique)
- Basta dar duplo clique no arquivo [`instalar_requisitos.bat`](file:///C:/Users/jonathan.lima/.gemini/antigravity/scratch/automacao_egov/instalar_requisitos.bat).
- Ele instalará automaticamente todas as bibliotecas necessárias (`pandas`, `openpyxl`, `playwright`, `customtkinter`, `tkcalendar`, `babel`) e o navegador Chromium do Playwright.

*Ou manualmente pelo terminal:*
```powershell
pip install -r requirements.txt
playwright install chromium
```

---

## 🖥️ Como Abrir o Aplicativo Visual (Front-End)

1. **Com 2 cliques no Windows:** basta dar duplo clique no arquivo [`Iniciar_Painel_EGOV.bat`](file:///C:/Users/jonathan.lima/.gemini/antigravity/scratch/automacao_egov/Iniciar_Painel_EGOV.bat).
2. **Pelo Terminal:**
   ```powershell
   cd C:\Users\jonathan.lima\.gemini\antigravity\scratch\automacao_egov
   python app.py
   ```

---

## 🧩 As 6 Etapas no Aplicativo

### 1. 📥 Aba: Inscritos GDF (`automacao.py`)
- **O que faz:** Faz login seguro no Portal de Sistemas do GDF com base na planilha `EAD_Controle_Geral.xlsx` (aba TURMAS).
- **Destino dos arquivos:** Salva em `Arquivos_Originais_Sistema/`.
- **Botão rápido:** `📂 Abrir Arquivos_Originais_Sistema`.

### 2. 🛠️ Aba: Tratar p/ Moodle (`tratar_csv_moodle.py`)
- **O que faz:** Lê de `Arquivos_Originais_Sistema/`, formata nomes em Title Case, formata CPFs (username e password) com 11 dígitos/zeros e converte e-mails para minúsculas.
- **Destino dos arquivos:** Salva em `Arquivos_CARREGAR/` no padrão `... - CARREGAR.csv`.
- **Botão rápido:** `📂 Abrir Pasta Arquivos_CARREGAR`.

### 3. 🎓 Aba: Notas Moodle (`baixar_notas_moodle.py`)
- **O que faz:** Faz login no Egovvirtual e baixa as planilhas de notas brutas das turmas listadas na planilha `EAD_Controle_Geral.xlsx` (aba TURMAS).
- **Destino dos arquivos:** Salva em `Arquivos_Originais_Moodle/`.
- **Botão rápido:** `📂 Abrir Arquivos_Originais_Moodle`.

### 4. 📊 Aba: Formatar SEDOC (`formatar_notas_moodle.py`)
- **O que faz:** Lê de `Arquivos_Originais_Moodle/`, seleciona o período de realização via calendário, aplica a regra de aprovação (RPIM 60%, demais 70%) e gera as planilhas oficiais com base legal.
- **Destino dos arquivos:** Salva em `Arquivos_Finais_SEDOC/` no padrão `..._SEDOC.xlsx`.
- **Botão rápido:** `📂 Abrir Pasta Arquivos_Finais_SEDOC`.

### 5. 📁 Aba: Mover Arquivos Finais (`copiar_para_servidor.py`)
- **O que faz:** 
  - Lê os arquivos de `Arquivos_CARREGAR/` (.csv) e `Arquivos_Finais_SEDOC/` (.xlsx).
  - Identifica automaticamente o curso e a turma:
    - Cursos PROAMIS (ex: `PROAMIS_AFG 06.2026...`) vão para `#PROAMIS\AFG\Turma 06.2026\`.
    - Cursos Gerais (ex: `AGLP 01.2026...`) vão para `AGLP\Turma 01.2026\`.
  - **Preservação local:** Mantém seus arquivos originais no seu computador e cria cópias organizadas no servidor.
  - **Proteção contra sobrescrita:** Se um arquivo já existir no destino, a aplicação abre uma janela de aviso com opções: *Sim*, *Sim para Todos*, *Não*, *Pular Todos* e *Cancelar*.
- **Destino padrão informado na aba:**
  `D:\SEGEA\RECENTES\EGOV\CEAD\Cursos Autoinstrucionais\# 2026\`
- **Botão rápido:** `📂 Abrir Pasta no Servidor (D:)`.

### 6. 🔒 Aba: Fechar Turmas (`fechar_turmas_sistema.py`)
- **O que faz:**
  - Faz login no Portal de Sistemas do GDF (`https://sistemas.df.gov.br/PortalDeServicos/Login`).
  - Navega até o sistema EGOV e aplica os filtros: **Modalidade: EAD** e **Situação Abertura Turma: Aberta**.
  - Lê a planilha `EAD_Controle_Geral.xlsx` (aba `TURMAS`) e localiza as turmas abertas prontas para serem encerradas.
  - Para cada turma, clica no ícone de **cadeado aberto** (`PopUp_FecharTurma.aspx`), confirma o fechamento no botão vermelho **Sim**, aguarda o processamento do sistema e valida a mensagem oficial (*"Esta turma foi fechada para realização. A listagem de alunos não poderá ser modificada"*).
- **Botão rápido:** `📂 Abrir Planilha EAD_Controle_Geral`.

---

## 🔐 Segurança e Boas Práticas
- Nenhuma senha fica salva no código ou no aplicativo.
- O campo de senha mascara os caracteres digitados (`••••••`).
- As tarefas rodam em segundo plano (threads) sem travar a interface gráfica.
- Totalmente compatível com a estrutura de rede da EGOV.
