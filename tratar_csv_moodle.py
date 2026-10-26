"""
Automação EGOV - Tratamento de Arquivos CSV para Importação no Moodle

O que este script faz:
1. Localiza os arquivos .csv baixados do sistema (na pasta 'downloads/' ou na pasta atual).
2. Lê os arquivos com a codificação correta do sistema (UTF-16LE com separador ';').
3. Remove linhas vazias.
4. Padroniza Nomes e Sobrenomes para Capitalizado (Primeira Letra Maiúscula).
5. Formata o CPF (username) com 11 dígitos numéricos e zeros à esquerda.
6. Padroniza e-mails para letras minúsculas.
7. Substitui "Sistema" por "CARREGAR" no nome do arquivo.
8. Salva os arquivos prontos e compatíveis na pasta 'Prontos_Para_Moodle/'.
"""

import os
import sys
import glob
import re
import pandas as pd

# Configuração segura para evitar erros de acentuação no terminal do Windows
if hasattr(sys.stdout, 'encoding') and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def tratar_arquivos_csv():
    # 1. Identifica a pasta atual do script
    pasta_atual = os.path.dirname(os.path.abspath(__file__))

    # 2. Cria a subpasta para salvar os arquivos tratados (se não existir)
    pasta_saida = os.path.join(pasta_atual, "Arquivos_CARREGAR")
    os.makedirs(pasta_saida, exist_ok=True)

    # 3. Localiza os arquivos .csv (procura primeiro na pasta 'Arquivos_Originais_Sistema')
    pasta_origem = os.path.join(pasta_atual, "Arquivos_Originais_Sistema")
    if os.path.exists(pasta_origem):
        arquivos_csv = glob.glob(os.path.join(pasta_origem, "*.csv"))
    else:
        arquivos_csv = []

    # Se não houver, procura na pasta antiga 'downloads' ou na raiz
    if not arquivos_csv:
        pasta_antiga = os.path.join(pasta_atual, "downloads")
        if os.path.exists(pasta_antiga):
            arquivos_csv = glob.glob(os.path.join(pasta_antiga, "*.csv"))

    if not arquivos_csv:
        arquivos_csv = glob.glob(os.path.join(pasta_atual, "*.csv"))

    # Remove da lista qualquer arquivo que por ventura já esteja na pasta de saída
    arquivos_csv = [arq for arq in arquivos_csv if "Arquivos_CARREGAR" not in arq and "Prontos_Para_Moodle" not in arq]

    if len(arquivos_csv) == 0:
        print("❌ Nenhum arquivo CSV foi encontrado para processar (verifique a pasta 'Arquivos_Originais_Sistema').")
        return

    print("=" * 60)
    print("🛠️ TRATAMENTO DE ARQUIVOS CSV PARA O MOODLE")
    print("=" * 60)
    print(f"Iniciando automação... Foram encontrados {len(arquivos_csv)} arquivos CSV.")
    print("-" * 60)

    total_tratados = 0

    # 4. Processa cada arquivo individualmente
    for arquivo in arquivos_csv:
        nome_arquivo = os.path.basename(arquivo)
        print(f"Processando: {nome_arquivo}...")

        try:
            # Lê o arquivo original (UTF-16LE com separador ponto e vírgula)
            try:
                df = pd.read_csv(arquivo, encoding='utf-16le', sep=';', on_bad_lines='skip')
            except Exception:
                # Fallback caso algum arquivo venha em utf-8 ou latin1
                df = pd.read_csv(arquivo, encoding='utf-8', sep=';', on_bad_lines='skip')

            # Remove linhas totalmente vazias
            df = df.dropna(how='all')

            # Padroniza Nomes e Sobrenomes (Primeira Letra Maiúscula)
            for coluna in ['firstname', 'lastname']:
                if coluna in df.columns:
                    df[coluna] = df[coluna].astype(str).str.strip().str.title()

            # Formata o CPF (username) com 11 dígitos e zeros à esquerda
            if 'username' in df.columns:
                df['username'] = df['username'].astype(str).str.split('.').str[0]
                df['username'] = df['username'].str.replace(r'\D', '', regex=True).str.zfill(11)

            # Padroniza E-mails para minúsculas
            if 'email' in df.columns:
                df['email'] = df['email'].astype(str).str.strip().str.lower()

            # 5. Salva o arquivo com o novo nome substituindo "Sistema" por "CARREGAR"
            novo_nome = re.sub(r'sistema', 'CARREGAR', nome_arquivo, flags=re.IGNORECASE)
            
            # Se por acaso o nome original não tinha a palavra "Sistema", garante o sufixo CARREGAR
            if novo_nome == nome_arquivo:
                nome_sem_ext = os.path.splitext(nome_arquivo)[0]
                novo_nome = f"{nome_sem_ext} - CARREGAR.csv"

            caminho_saida = os.path.join(pasta_saida, novo_nome)
            
            # Exporta em latin1 / sep ';' conforme padrão de importação do Moodle
            df.to_csv(caminho_saida, index=False, encoding='latin1', sep=';')

            print(f"  -> ✅ Concluído! Salvo como: {novo_nome} ({len(df)} alunos)")
            total_tratados += 1

        except Exception as e:
            print(f"  -> ❌ ERRO ao processar '{nome_arquivo}': {e}")

    print("-" * 60)
    print(f"🎉 Automação finalizada! {total_tratados} arquivos prontos na pasta:")
    print(f"📂 {pasta_saida}")
    print("=" * 60)

if __name__ == "__main__":
    tratar_arquivos_csv()
