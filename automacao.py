"""
Automação EGOV - Escola de Governo do DF
Fluxo Completo:
1. Solicitação interativa e segura de credenciais (Usuário e Senha oculta com getpass)
2. Login no Portal de Sistemas do GDF com o órgão escolhido
3. Acesso ao sistema EGOV na Galeria de Sistemas
4. Aplicação dos filtros (Modalidade: EAD | Situação: Fechada) e Pesquisa
5. Leitura da planilha 'EAD_Controle_Geral.xlsx' (aba TURMAS)
6. Download automático do arquivo .CSV dos inscritos de cada turma da lista
7. Renomeação no padrão: "SIGLA TURMA.ANO - Sistema.csv"
8. Logs exibidos no console informando os resultados

Biblioteca: Playwright + Pandas + Getpass
"""

import os
import sys
import re
import getpass
from datetime import datetime
import pandas as pd
from playwright.sync_api import sync_playwright

# Configuração segura para evitar erros de acentuação no terminal do Windows
if hasattr(sys.stdout, 'encoding') and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# --- CONFIGURAÇÕES DE ACESSO ---
URL_LOGIN = "https://sistemas.df.gov.br/PortalDeServicos/Login"

# --- ARQUIVOS E PASTAS ---
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_PLANILHA = os.path.join(DIRETORIO_ATUAL, "EAD_Controle_Geral.xlsx")
PASTA_DOWNLOADS = os.path.join(DIRETORIO_ATUAL, "Arquivos_Originais_Sistema")

def extrair_turma_id(url_sistema):
    """Extrai o número do TurmaId da URL do sistema (ex: TurmaId=6125 -> 6125)"""
    if pd.isna(url_sistema) or not url_sistema:
        return None
    match = re.search(r'TurmaId=(\d+)', str(url_sistema), re.IGNORECASE)
    return match.group(1) if match else None

def formatar_turma(valor_turma):
    """Garante que a turma tenha pelo menos 2 dígitos (ex: 6 -> '06', 12 -> '12')"""
    try:
        return f"{int(float(valor_turma)):02d}"
    except Exception:
        val_str = str(valor_turma).split('.')[0].strip()
        return val_str.zfill(2) if val_str else "01"

def extrair_ano(valor_data):
    """Extrai o ano da data de início (ex: '2026-10-05' -> '2026')"""
    try:
        dt = pd.to_datetime(valor_data)
        if not pd.isna(dt):
            return str(dt.year)
    except Exception:
        pass
    match = re.search(r'(20\d{2})', str(valor_data))
    if match:
        return match.group(1)
    return str(datetime.now().year)

def coletar_credenciais():
    """Solicita login e senha com segurança no terminal, sem salvar no código"""
    print("\n" + "=" * 60)
    print("🔐 IDENTIFICAÇÃO DO USUÁRIO (ACESSO PORTAL GDF)")
    print("=" * 60)
    print("Insira suas credenciais corporativas para iniciar a automação.\n")

    usuario = input("👤 Digite seu Usuário (ex: nome.sobrenome): ").strip()
    while not usuario:
        print("   ⚠️ O usuário não pode ser vazio!")
        usuario = input("👤 Digite seu Usuário: ").strip()

    # getpass oculta a senha enquanto o usuário digita (não aparece na tela)
    senha = getpass.getpass("🔑 Digite sua Senha (a digitação fica invisível): ")
    while not senha:
        print("   ⚠️ A senha não pode ser vazia!")
        senha = getpass.getpass("🔑 Digite sua Senha: ")

    orgao = input("🏢 Órgão [pressione ENTER para 'SEEC']: ").strip()
    if not orgao:
        orgao = "SEEC"

    return usuario, senha, orgao

def executar_automacao():
    usuario, senha, orgao = coletar_credenciais()
    executar_automacao_gui(usuario, senha, orgao)

def executar_automacao_gui(usuario, senha, orgao="SEEC", log_fn=print):
    print("=" * 60)
    print(">> AUTOMAÇÃO EGOV - DOWNLOAD DE INSCRITOS DOS CURSOS EAD")
    print("=" * 60)

    # 1. Garante que a pasta de downloads existe
    os.makedirs(PASTA_DOWNLOADS, exist_ok=True)

    # 2. Carrega a planilha de cursos a baixar
    if not os.path.exists(ARQUIVO_PLANILHA):
        print(f"[ERRO] Planilha '{ARQUIVO_PLANILHA}' não encontrada!")
        return

    df_cursos = pd.read_excel(ARQUIVO_PLANILHA, sheet_name='TURMAS')
    
    # Extrai TurmaId do novo link de sistema (LINK FORMULÁRIO)
    if 'LINK FORMULÁRIO' in df_cursos.columns:
        df_cursos['TURMA_ID'] = df_cursos['LINK FORMULÁRIO'].apply(extrair_turma_id)
    else:
        df_cursos['TURMA_ID'] = None

    # Inicia o navegador com Playwright
    with sync_playwright() as p:
        print("\n[NAVEGADOR] Abrindo navegador...")
        browser = p.chromium.launch(headless=False, slow_mo=400)
        page = browser.new_page()

        # ==========================================
        # ETAPA 1: LOGIN NO PORTAL DE SISTEMAS
        # ==========================================
        print(f"\n[ETAPA 1] Acessando {URL_LOGIN}...")
        page.goto(URL_LOGIN)
        page.wait_for_selector("#Input_UsernameVal")

        print(f"[ETAPA 1] Preenchendo credenciais do usuário '{usuario}' e órgão '{orgao}'...")
        page.fill("#Input_UsernameVal", usuario)
        page.fill("#Input_PasswordVal", senha)
        page.select_option("#DropdownOrgao", label=orgao)
        page.click("button[type=submit]")

        try:
            page.wait_for_url("**/GaleriaSistemas", timeout=15000)
            print("[ETAPA 1] ✅ Login realizado com sucesso!")
        except Exception:
            # Verifica se apareceu mensagem de erro na tela de login
            msg_erro = page.locator(".feedback-message, .ValidationMessage").all_inner_texts()
            detalhe = f" Detalhe: {' | '.join(msg_erro)}" if msg_erro else ""
            print(f"\n❌ [FALHA NO LOGIN] Não foi possível logar no sistema.{detalhe}")
            print("Verifique se seu usuário, senha ou órgão foram digitados corretamente.")
            browser.close()
            return

        # ==========================================
        # ETAPA 2: ACESSO AO SISTEMA EGOV
        # ==========================================
        print("\n[ETAPA 2] Entrando no sistema EGOV...")
        card_egov = page.locator('text="EGOV"')
        card_egov.first.wait_for(timeout=15000)
        card_egov.first.click()

        page.wait_for_url("**/egov/**", timeout=20000)
        page.wait_for_load_state("networkidle")
        print("[ETAPA 2] ✅ Tela de Turmas do EGOV carregada com sucesso!")

        # ==========================================
        # ETAPA 3: APLICAÇÃO DOS FILTROS E BUSCA
        # ==========================================
        print("\n[ETAPA 3] Aplicando filtros de busca...")
        print("   -> Modalidade: EAD")
        page.locator('select:has(option:text-is("EAD"))').select_option(label="EAD")

        print("   -> Situação Abertura Turma: Fechada")
        page.locator('label:has-text("Fechada")').click()

        print("   -> Clicando em 'Pesquisar'...")
        page.click('input[value="Pesquisar"]')
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(3000)
        print("[ETAPA 3] ✅ Turmas pesquisadas e tabela carregada!")

        # ==========================================
        # ETAPA 4: DOWNLOAD DOS ARQUIVOS .CSV
        # ==========================================
        print("\n[ETAPA 4] Iniciando download dos arquivos .csv das turmas...")

        total_baixados = 0
        hora_atual = datetime.now().strftime("%d/%m/%Y %H:%M")

        # Função de busca de turma com paginação
        def localizar_turma_com_paginacao(turma_id, max_paginas=20):
            btn_pag1 = page.locator('a:text-is("1")').first
            if btn_pag1.count() > 0 and btn_pag1.is_visible():
                try:
                    btn_pag1.click()
                    page.wait_for_load_state("networkidle", timeout=5000)
                    page.wait_for_timeout(1500)
                except Exception:
                    pass

            pagina_atual = 1
            while pagina_atual <= max_paginas:
                linha = page.locator(f'tr:has(a[href*="TurmaId={turma_id}" i])').first
                if linha.count() > 0:
                    return linha, pagina_atual

                btn_proximo = page.locator('a, button, [role="button"]').filter(has_text=re.compile(r'pr[oó]ximo', re.IGNORECASE)).first
                prox_num = str(pagina_atual + 1)
                btn_num = page.locator(f'a:text-is("{prox_num}"), button:text-is("{prox_num}")').first

                btn_avancar = None
                if btn_proximo.count() > 0 and btn_proximo.is_visible():
                    classes = (btn_proximo.get_attribute("class") or "").lower()
                    pai = btn_proximo.locator("xpath=..")
                    classes_pai = (pai.get_attribute("class") or "").lower() if pai.count() > 0 else ""
                    if "disabled" not in classes and "disabled" not in classes_pai:
                        btn_avancar = btn_proximo

                if btn_avancar is None and btn_num.count() > 0 and btn_num.is_visible():
                    btn_avancar = btn_num

                if btn_avancar is None:
                    return None, pagina_atual

                btn_avancar.click()
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                page.wait_for_timeout(2000)
                pagina_atual += 1

            return None, pagina_atual

        for index, row in df_cursos.iterrows():
            turma_id = row['TURMA_ID']
            sigla = str(row.get('SIGLA', 'CURSO')).strip()
            turma_num = formatar_turma(row.get('TURMA', ''))
            ano = extrair_ano(row.get('DATA INÍCIO', '2026'))
            nome_curso = str(row.get('Curso', 'Curso')).strip()

            if not turma_id:
                print(f"   [AVISO] Linha {index + 1}: Sem TurmaId válido. Pulando...")
                continue

            nome_arquivo_destino = f"{sigla} {turma_num}.{ano} - Sistema.csv"
            nome_arquivo_destino = re.sub(r'[\\/*?:"<>|]', "", nome_arquivo_destino)
            caminho_salvar = os.path.join(PASTA_DOWNLOADS, nome_arquivo_destino)

            print(f"\n   [CURSO {index + 1}/{len(df_cursos)}] Buscando ID {turma_id}: {nome_curso} ({sigla} - Turma {turma_num})...")

            # Localiza a linha da tabela navegando pelas páginas
            linha_tabela, num_pag = localizar_turma_com_paginacao(turma_id)

            if linha_tabela is None:
                print(f"   -> [AVISO] Turma {turma_id} não encontrada em nenhuma página da tabela.")
                continue

            # Localiza o ícone de folha (Download Relatório CSV) na linha
            botao_download = linha_tabela.locator('a[title*="Download Relat"], a:has(.fa-file-text-o)').first

            if botao_download.count() == 0:
                print(f"   -> [AVISO] Botão de download não encontrado na linha da turma {turma_id}.")
                continue

            print(f"   -> Clicando no ícone de folha para baixar...")
            try:
                with page.expect_download(timeout=25000) as download_info:
                    botao_download.click()

                download = download_info.value
                download.save_as(caminho_salvar)
                
                tamanho_kb = os.path.getsize(caminho_salvar) / 1024
                print(f"   -> [SUCESSO] Salvo como: \"{nome_arquivo_destino}\" ({tamanho_kb:.1f} KB)")

                total_baixados += 1
            except Exception as e:
                print(f"   -> [ERRO] Falha ao baixar turma {turma_id}: {e}")

            page.wait_for_timeout(2000)

        # Atualização no console
        print(f"\n[INFO] Os logs de download foram exibidos acima. O arquivo original Excel '{os.path.basename(ARQUIVO_PLANILHA)}' não será alterado para preservar suas formatações.")

        print("\n" + "=" * 60)
        print(f">> CONCLUÍDO! Total de {total_baixados} de {len(df_cursos)} arquivos baixados.")
        print(f">> Arquivos salvos na pasta: {PASTA_DOWNLOADS}")
        print("=" * 60)

        page.wait_for_timeout(3000)
        browser.close()

if __name__ == "__main__":
    executar_automacao()
