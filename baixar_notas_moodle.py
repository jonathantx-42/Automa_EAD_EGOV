"""
Automação Moodle - Egovvirtual (Escola de Governo do DF)
Download automático de notas da planilha de cursos.

Segurança:
- Solicitação interativa de Usuário e Senha com getpass (sem salvar senhas no código)
- Suporte para uso por qualquer colega de equipe com suas próprias credenciais
"""

import os
import sys
import time
import getpass
import pandas as pd
from playwright.sync_api import sync_playwright
from logger_util import AutoLogger

# Configuração segura para evitar erros de acentuação no terminal do Windows
if hasattr(sys.stdout, 'encoding') and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

def coletar_credenciais():
    """Solicita login e senha com segurança no terminal, sem salvar no código"""
    print("\n" + "=" * 60)
    print("🔐 IDENTIFICAÇÃO DO USUÁRIO (EGOVVIRTUAL / MOODLE)")
    print("=" * 60)
    print("Insira suas credenciais do Egovvirtual para iniciar a automação.\n")

    usuario = input("👤 Digite seu Usuário: ").strip()
    while not usuario:
        print("   ⚠️ O usuário não pode ser vazio!")
        usuario = input("👤 Digite seu Usuário: ").strip()

    # getpass oculta a senha enquanto digita (não aparece na tela)
    senha = getpass.getpass("🔑 Digite sua Senha (a digitação fica invisível): ")
    while not senha:
        print("   ⚠️ A senha não pode ser vazia!")
        senha = getpass.getpass("🔑 Digite sua Senha: ")

    return usuario, senha

def baixar_notas_da_planilha():
    usuario, senha = coletar_credenciais()
    baixar_notas_gui(usuario, senha)

def baixar_notas_gui(usuario, senha, log_fn=print):
    logger = AutoLogger("Notas_Moodle", usuario=usuario, callback_console=log_fn)
    def log(msg):
        logger.log(msg)

    log("=" * 60)
    log(">> AUTOMAÇÃO MOODLE - DOWNLOAD DE NOTAS DOS CURSOS")
    log("=" * 60)

    # 1. Leitura da planilha de cursos (procura com C maiúsculo ou minúsculo)
    pasta_app = os.path.dirname(os.path.abspath(__file__))
    pasta_saida = os.path.join(pasta_app, "Arquivos_Originais_Moodle")
    os.makedirs(pasta_saida, exist_ok=True)

    arquivo_planilha = os.path.join(pasta_app, 'EAD_Controle_Geral.xlsx')

    if not os.path.exists(arquivo_planilha):
        log(f"\n❌ [ERRO] A planilha 'EAD_Controle_Geral.xlsx' não foi encontrada na pasta!")
        log("Certifique-se de que o arquivo 'EAD_Controle_Geral.xlsx' está na pasta do projeto.")
        logger.finalizar(status_geral="Falha: Planilha não encontrada")
        return

    log("\n1. Lendo a planilha de cursos...")
    try:
        tabela_cursos = pd.read_excel(arquivo_planilha, sheet_name='TURMAS')
    except Exception as e:
        log(f"\n❌ [ERRO] Falha ao ler a planilha: {e}")
        logger.finalizar(status_geral=f"Falha ao ler planilha: {e}")
        return

    log(f"   -> {len(tabela_cursos)} cursos encontrados na planilha.")

    with sync_playwright() as p:
        log("\n2. Abrindo o navegador...")
        navegador = p.chromium.launch(headless=False, slow_mo=300)
        pagina = navegador.new_page()

        # --- FAZ O LOGIN ---
        log("3. Fazendo login no Egovvirtual...")
        pagina.goto("https://egovvirtual.df.gov.br/login/index.php")
        pagina.wait_for_selector("#username")

        pagina.fill('#username', usuario)
        pagina.fill('#password', senha)
        pagina.click('#loginbtn')

        pagina.wait_for_load_state("networkidle")
        time.sleep(3)

        # Validação simples de login
        if "login/index.php" in pagina.url and pagina.locator(".alert-danger, #loginerrormessage, .loginerrors").count() > 0:
            log("\n❌ [FALHA NO LOGIN] Usuário ou senha incorretos no Egovvirtual!")
            log("Verifique suas credenciais e tente executar o script novamente.")
            logger.finalizar(status_geral="Falha no Login")
            navegador.close()
            return

        log("✅ Login realizado com sucesso no Egovvirtual!")

        log("\n--- INICIANDO OS DOWNLOADS DAS NOTAS ---")
        total_baixados = 0

        # --- REPETE O PROCESSO PARA CADA LINHA DA PLANILHA ---
        for linha in tabela_cursos.index:
            nome_do_curso = str(tabela_cursos.loc[linha, 'Curso']).strip()
            link_do_curso = tabela_cursos.loc[linha, 'ID MOODLE']
            sigla = str(tabela_cursos.loc[linha, 'SIGLA']).strip() if 'SIGLA' in tabela_cursos.columns else ""
            turma_num = str(tabela_cursos.loc[linha, 'TURMA']).strip() if 'TURMA' in tabela_cursos.columns else ""

            # Pega o link (ex: https://.../view.php?id=123) e extrai só o "123"
            id_do_curso = str(link_do_curso).split('id=')[-1].strip()
            item_identificador = f"Curso {id_do_curso} ({sigla} - Turma {turma_num})" if sigla else f"Curso {id_do_curso} ({nome_do_curso})"

            log(f"\n[{linha + 1}/{len(tabela_cursos)}] Processando: {nome_do_curso}")

            if not id_do_curso or not id_do_curso.isdigit():
                log(f"   ⚠️ Link inválido ou sem ID numérico: {link_do_curso}. Pulando...")
                continue

            logger.iniciar_item(item_identificador)

            # Monta o link direto para a página de exportar notas em Excel
            link_download = f"https://egovvirtual.df.gov.br/grade/export/xls/index.php?id={id_do_curso}"

            log("   -> Acessando a página de notas...")
            pagina.goto(link_download)
            pagina.wait_for_load_state("networkidle")
            time.sleep(2)

            # Verifica se o botão de download existe na página
            botao_download = pagina.locator('#id_submitbutton')
            if botao_download.count() == 0:
                log("   ⚠️ Botão de download (#id_submitbutton) não encontrado nesta turma. Pulando...")
                logger.finalizar_item(item_identificador, status="Sem Botão Download", detalhes="Botão #id_submitbutton não encontrado")
                continue

            log("   -> Baixando a planilha de notas...")
            try:
                with pagina.expect_download(timeout=25000) as download_info:
                    botao_download.click()

                download = download_info.value
                nome_original = download.suggested_filename
                caminho_salvar = os.path.join(pasta_saida, nome_original)

                # Salva o arquivo na pasta Arquivos_Originais_Moodle
                download.save_as(caminho_salvar)
                logger.finalizar_item(item_identificador, status="Sucesso", detalhes=f"Salvo: {nome_original}")
                log(f"   ✅ SUCESSO! Salvo em: Arquivos_Originais_Moodle/{nome_original}")
                total_baixados += 1
            except Exception as e:
                logger.finalizar_item(item_identificador, status="Falha", detalhes=str(e))
                log(f"   ❌ Erro ao baixar notas do curso {nome_do_curso}: {e}")

            time.sleep(1)

        status_geral = f"{total_baixados} de {len(tabela_cursos)} arquivos de notas baixados"
        logger.finalizar(status_geral=status_geral)

        log("\n" + "=" * 60)
        log(f"🎉 CONCLUÍDO! Total de {total_baixados} de {len(tabela_cursos)} arquivos de notas baixados.")
        log("=" * 60)

        time.sleep(2)
        navegador.close()

if __name__ == "__main__":
    baixar_notas_da_planilha()
