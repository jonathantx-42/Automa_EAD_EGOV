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
    print("=" * 60)
    print(">> AUTOMAÇÃO MOODLE - DOWNLOAD DE NOTAS DOS CURSOS")
    print("=" * 60)

    # 1. Leitura da planilha de cursos (procura com C maiúsculo ou minúsculo)
    pasta_app = os.path.dirname(os.path.abspath(__file__))
    pasta_saida = os.path.join(pasta_app, "Arquivos_Originais_Moodle")
    os.makedirs(pasta_saida, exist_ok=True)

    arquivo_planilha = os.path.join(pasta_app, 'EAD_Controle_Geral.xlsx')

    if not os.path.exists(arquivo_planilha):
        print(f"\n❌ [ERRO] A planilha 'EAD_Controle_Geral.xlsx' não foi encontrada na pasta!")
        print("Certifique-se de que o arquivo 'EAD_Controle_Geral.xlsx' está na pasta do projeto.")
        return

    print("\n1. Lendo a planilha de cursos...")
    tabela_cursos = pd.read_excel(arquivo_planilha, sheet_name='TURMAS')
    print(f"   -> {len(tabela_cursos)} cursos encontrados na planilha.")

    with sync_playwright() as p:
        print("\n2. Abrindo o navegador...")
        navegador = p.chromium.launch(headless=False, slow_mo=300)
        pagina = navegador.new_page()

        # --- FAZ O LOGIN ---
        print("3. Fazendo login no Egovvirtual...")
        pagina.goto("https://egovvirtual.df.gov.br/login/index.php")
        pagina.wait_for_selector("#username")

        pagina.fill('#username', usuario)
        pagina.fill('#password', senha)
        pagina.click('#loginbtn')

        pagina.wait_for_load_state("networkidle")
        time.sleep(3)

        # Validação simples de login
        if "login/index.php" in pagina.url and pagina.locator(".alert-danger, #loginerrormessage, .loginerrors").count() > 0:
            print("\n❌ [FALHA NO LOGIN] Usuário ou senha incorretos no Egovvirtual!")
            print("Verifique suas credenciais e tente executar o script novamente.")
            navegador.close()
            return

        print("✅ Login realizado com sucesso no Egovvirtual!")

        print("\n--- INICIANDO OS DOWNLOADS DAS NOTAS ---")
        total_baixados = 0

        # --- REPETE O PROCESSO PARA CADA LINHA DA PLANILHA ---
        for linha in tabela_cursos.index:
            nome_do_curso = tabela_cursos.loc[linha, 'Curso']
            link_do_curso = tabela_cursos.loc[linha, 'ID MOODLE']

            print(f"\n[{linha + 1}/{len(tabela_cursos)}] Processando: {nome_do_curso}")

            # Pega o link (ex: https://.../view.php?id=123) e extrai só o "123"
            id_do_curso = str(link_do_curso).split('id=')[-1].strip()

            if not id_do_curso or not id_do_curso.isdigit():
                print(f"   ⚠️ Link inválido ou sem ID numérico: {link_do_curso}. Pulando...")
                continue

            # Monta o link direto para a página de exportar notas em Excel
            link_download = f"https://egovvirtual.df.gov.br/grade/export/xls/index.php?id={id_do_curso}"

            print("   -> Acessando a página de notas...")
            pagina.goto(link_download)
            pagina.wait_for_load_state("networkidle")
            time.sleep(2)

            # Verifica se o botão de download existe na página
            botao_download = pagina.locator('#id_submitbutton')
            if botao_download.count() == 0:
                print("   ⚠️ Botão de download (#id_submitbutton) não encontrado nesta turma. Pulando...")
                continue

            print("   -> Baixando a planilha de notas...")
            try:
                with pagina.expect_download(timeout=25000) as download_info:
                    botao_download.click()

                download = download_info.value
                nome_original = download.suggested_filename
                caminho_salvar = os.path.join(pasta_saida, nome_original)

                # Salva o arquivo na pasta Arquivos_Originais_Moodle
                download.save_as(caminho_salvar)
                print(f"   ✅ SUCESSO! Salvo em: Arquivos_Originais_Moodle/{nome_original}")
                total_baixados += 1
            except Exception as e:
                print(f"   ❌ Erro ao baixar notas do curso {nome_do_curso}: {e}")

            time.sleep(1)

        print("\n" + "=" * 60)
        print(f"🎉 CONCLUÍDO! Total de {total_baixados} de {len(tabela_cursos)} arquivos de notas baixados.")
        print("=" * 60)

        time.sleep(2)
        navegador.close()

if __name__ == "__main__":
    baixar_notas_da_planilha()
