"""
Automação EGOV - Fechamento de Turmas EAD no Portal de Sistemas do GDF

O que este script faz:
1. Faz login seguro no Portal de Sistemas do GDF (https://sistemas.df.gov.br/PortalDeServicos/Login).
2. Acessa a Galeria de Sistemas e entra no sistema EGOV (https://sistemas.df.gov.br/egov/).
3. Aplica os filtros de busca de turmas:
   - Modalidade: EAD
   - Situação Abertura Turma: Aberta
4. Realiza a pesquisa para listar todas as turmas abertas.
5. Lê a planilha 'EAD_Controle_Geral.xlsx' (aba 'TURMAS') para obter os IDs e nomes das turmas que devem ser fechadas.
6. Para cada turma indicada:
   - Localiza a linha correspondente pelo TurmaId.
   - Clica no ícone de "cadeado aberto" (PopUp_FecharTurma.aspx?NumeroVagas=...&TurmaId=...).
   - Localiza a caixa de confirmação e clica em "Sim" (botão vermelho).
   - Aguarda o processamento do fechamento (tempo variável conforme o número de inscritos).
   - Confirma a mensagem oficial: "Esta turma foi fechada para realização. A listagem de alunos não poderá ser modificada".
   - Fecha a caixa de confirmação e atualiza a lista de turmas para a próxima ação.
7. Exibe um relatório detalhado das turmas fechadas com sucesso.
"""

import os
import sys
import re
import time
import getpass
import pandas as pd
from datetime import datetime
from playwright.sync_api import sync_playwright
from logger_util import AutoLogger

# --- CONFIGURAÇÃO DE DIRETÓRIOS E ARQUIVOS ---
DIRETORIO_ATUAL = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_PLANILHA = os.path.join(DIRETORIO_ATUAL, "EAD_Controle_Geral.xlsx")
if not os.path.exists(ARQUIVO_PLANILHA):
    copia_planilha = os.path.join(DIRETORIO_ATUAL, "EAD_Controle_Geral - Copia.xlsx")
    if os.path.exists(copia_planilha):
        ARQUIVO_PLANILHA = copia_planilha

URL_LOGIN = "https://sistemas.df.gov.br/PortalDeServicos/Login"
URL_EGOV = "https://sistemas.df.gov.br/egov/"

def extrair_turma_id(url_sistema):
    """Extrai o número do TurmaId da URL do formulário (ex: TurmaId=5339 -> 5339)"""
    if pd.isna(url_sistema) or not url_sistema:
        return None
    match = re.search(r'TurmaId=(\d+)', str(url_sistema), re.IGNORECASE)
    return match.group(1) if match else None

def coletar_credenciais():
    """Solicita login e senha com segurança no terminal"""
    print("\n" + "=" * 60)
    print("🔐 IDENTIFICAÇÃO DO USUÁRIO - FECHAMENTO DE TURMAS GDF")
    print("=" * 60)
    print("Insira suas credenciais corporativas para iniciar a automação.\n")

    usuario = input("👤 Digite seu Usuário (ex: nome.sobrenome): ").strip()
    while not usuario:
        print("   ⚠️ O usuário não pode ser vazio!")
        usuario = input("👤 Digite seu Usuário: ").strip()

    senha = getpass.getpass("🔑 Digite sua Senha (a digitação fica oculta): ")
    while not senha:
        print("   ⚠️ A senha não pode ser vazia!")
        senha = getpass.getpass("🔑 Digite sua Senha: ")

    orgao = input("🏢 Órgão [pressione ENTER para 'SEEC']: ").strip()
    if not orgao:
        orgao = "SEEC"

    return usuario, senha, orgao

def executar_fechamento():
    """Ponto de entrada para execução via terminal"""
    usuario, senha, orgao = coletar_credenciais()
    fechar_turmas_gui(usuario, senha, orgao)

def fechar_turmas_gui(usuario, senha, orgao="SEEC", log_fn=print):
    """
    Função principal de fechamento de turmas, compatível tanto com o terminal
    quanto com a interface gráfica (app.py) via log_fn.
    """
    logger = AutoLogger("Fechar_Turmas_EAD", usuario=usuario, callback_console=log_fn)
    def log(msg):
        logger.log(msg)

    log("=" * 65)
    log(">> AUTOMAÇÃO EGOV - FECHAMENTO DE TURMAS EAD NO SISTEMA")
    log("=" * 65)

    # 1. Carrega a planilha de turmas
    if not os.path.exists(ARQUIVO_PLANILHA):
        log(f"\n❌ [ERRO] Planilha '{os.path.basename(ARQUIVO_PLANILHA)}' não encontrada no diretório!")
        return

    try:
        df_cursos = pd.read_excel(ARQUIVO_PLANILHA, sheet_name='TURMAS')
    except Exception as e:
        log(f"\n❌ [ERRO] Falha ao ler a aba 'TURMAS' da planilha: {e}")
        return

    if 'LINK FORMULÁRIO' not in df_cursos.columns:
        log("\n❌ [ERRO] Coluna 'LINK FORMULÁRIO' não encontrada na aba 'TURMAS'.")
        return

    # Filtra apenas linhas com link válido
    df_cursos['TURMA_ID'] = df_cursos['LINK FORMULÁRIO'].apply(extrair_turma_id)
    df_turmas = df_cursos[df_cursos['TURMA_ID'].notna()].copy()

    if len(df_turmas) == 0:
        log("\n⚠️ [AVISO] Nenhuma turma válida encontrada com 'TurmaId' na planilha para fechar.")
        return

    log(f"\n[INFO] Planilha carregada: {len(df_turmas)} turma(s) mapeada(s) para fechamento.")

    with sync_playwright() as p:
        log("\n[NAVEGADOR] Abrindo navegador...")
        browser = p.chromium.launch(headless=False, slow_mo=300)
        context = browser.new_context()
        page = context.new_page()

        # ==========================================
        # ETAPA 1: LOGIN NO PORTAL DE SISTEMAS
        # ==========================================
        log(f"\n[ETAPA 1] Acessando {URL_LOGIN}...")
        page.goto(URL_LOGIN)
        page.wait_for_selector("#Input_UsernameVal", timeout=20000)

        log(f"[ETAPA 1] Autenticando com usuário '{usuario}' no órgão '{orgao}'...")
        page.fill("#Input_UsernameVal", usuario)
        page.fill("#Input_PasswordVal", senha)
        page.select_option("#DropdownOrgao", label=orgao)
        page.click("button[type=submit]")

        try:
            page.wait_for_url("**/GaleriaSistemas", timeout=15000)
            log("[ETAPA 1] ✅ Login realizado com sucesso!")
        except Exception:
            msg_erro = page.locator(".feedback-message, .ValidationMessage").all_inner_texts()
            detalhe = f" Detalhe: {' | '.join(msg_erro)}" if msg_erro else ""
            log(f"\n❌ [FALHA NO LOGIN] Não foi possível autenticar no sistema.{detalhe}")
            browser.close()
            return

        # ==========================================
        # ETAPA 2: ACESSO AO SISTEMA EGOV
        # ==========================================
        log("\n[ETAPA 2] Acessando o sistema EGOV...")
        card_egov = page.locator('text="EGOV"')
        card_egov.first.wait_for(timeout=15000)
        card_egov.first.click()

        page.wait_for_url("**/egov/**", timeout=25000)
        page.wait_for_load_state("networkidle")
        log("[ETAPA 2] ✅ Sistema EGOV carregado!")

        # Função interna para aplicar os filtros e pesquisar
        def aplicar_pesquisa_aberta():
            log("   -> Aplicando filtros (Modalidade: EAD | Situação: Aberta)...")
            
            # Aguarda qualquer overlay do OutSystems sumir
            try:
                page.locator('#divWait, .Feedback_AjaxWait').wait_for(state='hidden', timeout=15000)
            except Exception:
                pass

            # Modalidade EAD
            sel_mod = page.locator('select:has(option:text-is("EAD"))')
            if sel_mod.count() > 0:
                sel_mod.first.select_option(label="EAD")

            # Situação Abertura Turma: Aberta (só clica se já não estiver ativa)
            rad_aberta = page.locator('label:text-is("Aberta"), label:has-text("Aberta")').first
            if rad_aberta.count() > 0:
                classes = rad_aberta.get_attribute("class") or ""
                if "active" not in classes.lower():
                    rad_aberta.click()

            # Clicar em Pesquisar
            btn_pesq = page.locator('input[value="Pesquisar"], button:has-text("Pesquisar")').first
            btn_pesq.click()
            page.wait_for_load_state("networkidle")
            page.wait_for_timeout(2000)

        # ==========================================
        # ETAPA 3: APLICAÇÃO DOS FILTROS INICIAIS
        # ==========================================
        log("\n[ETAPA 3] Buscando turmas abertas...")
        aplicar_pesquisa_aberta()
        log("[ETAPA 3] ✅ Pesquisa de turmas abertas concluída!")

        # ==========================================
        # ETAPA 4: FECHAMENTO DE CADA TURMA
        # ==========================================
        log("\n[ETAPA 4] Iniciando processo de fechamento das turmas...")
        total_fechadas = 0

        # Rastreamento de popup de janela ou diálogos nativos
        popup_window = None
        def on_page_opened(new_page):
            nonlocal popup_window
            popup_window = new_page

        context.on("page", on_page_opened)

        dialogos_recebidos = []
        def on_dialog(dialog):
            msg_diag = dialog.message
            log(f"   [AVISO DO SISTEMA]: {msg_diag}")
            dialogos_recebidos.append(msg_diag)
            dialog.accept()

        page.on("dialog", on_dialog)

        # Função para buscar a turma navegando pelas páginas da tabela se necessário
        def localizar_turma_com_paginacao(turma_id, max_paginas=20):
            # Garante que começamos a busca a partir da página 1
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
                # 1. Checa se a turma está visível na página atual
                linha = page.locator(f'tr:has(a[href*="TurmaId={turma_id}" i])').first

                if linha.count() > 0:
                    icone = page.locator(f'a[href*="PopUp_FecharTurma" i][href*="TurmaId={turma_id}" i]').first
                    if icone.count() == 0:
                        icone = linha.locator('a[href*="FecharTurma" i], a[title*="Fechar" i], a:has(.fa-unlock), a:has(.fa-lock)').first
                    log(f"   -> ✅ Turma {turma_id} localizada na página {pagina_atual} da tabela!")
                    return linha, icone

                # 2. Se não encontrou, busca botão para avançar para a próxima página
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
                    log(f"   ⚠️ Percorridas {pagina_atual} páginas e a turma {turma_id} não foi encontrada.")
                    return None, None

                # Captura texto do rodapé (ex: '1 a 10 de 56 registros') antes de avançar
                texto_rodape_antes = ""
                elem_rodape = page.locator('text=/\\d+\\s+a\\s+\\d+\\s+de\\s+\\d+/i').first
                if elem_rodape.count() > 0:
                    texto_rodape_antes = elem_rodape.inner_text().strip()

                log(f"   -> Não encontrada na página {pagina_atual}. Avançando para a página {pagina_atual + 1}...")
                btn_avancar.click()

                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                page.wait_for_timeout(2000)

                if texto_rodape_antes:
                    try:
                        page.wait_for_function(
                            "antes => !document.body.innerText.includes(antes)",
                            arg=texto_rodape_antes,
                            timeout=5000
                        )
                        page.wait_for_timeout(1000)
                    except Exception:
                        pass

                pagina_atual += 1

            return None, None

        for idx, (_, row) in enumerate(df_turmas.iterrows(), start=1):
            turma_id = str(row['TURMA_ID']).strip()
            nome_curso = str(row.get('Curso', 'Curso')).strip()
            sigla = str(row.get('SIGLA', '')).strip()
            num_turma = str(row.get('TURMA', '')).strip()
            item_identificador = f"Turma {turma_id} ({sigla} - Turma {num_turma})"

            logger.iniciar_item(item_identificador)

            log(f"\n[{idx}/{len(df_turmas)}] Processando ID {turma_id}: {nome_curso} ({sigla} - Turma {num_turma})...")

            # Remove qualquer mensagem residual da turma anterior
            try:
                page.evaluate("() => document.querySelectorAll('.Feedback_Message_Success, .alert, .Feedback_Message_Wrapper, div[class*=\"Feedback\"]').forEach(el => el.remove())")
            except Exception:
                pass

            # 1. Localiza a linha e o ícone de cadeado navegando pelas páginas se necessário
            linha_tabela, icone_cadeado = localizar_turma_com_paginacao(turma_id)

            if linha_tabela is None:
                log(f"   ⚠️ Turma {turma_id} não foi encontrada em nenhuma página de turmas abertas (pode já estar fechada).")
                logger.finalizar_item(item_identificador, status="Não Encontrada", detalhes="Não localizada na tabela de turmas abertas")
                continue

            if icone_cadeado is None or icone_cadeado.count() == 0:
                log(f"   ⚠️ Turma {turma_id} localizada na tabela, mas não possui ícone de cadeado aberto (pode já estar fechada).")
                logger.finalizar_item(item_identificador, status="Sem Cadeado Aberto", detalhes="Turma não possui ícone de cadeado aberto")
                continue

            popup_window = None
            dialogos_recebidos.clear()

            log("   -> Clicando no ícone de cadeado aberto...")
            icone_cadeado.click()
            page.wait_for_timeout(2000)

            # Funções auxiliares para localização e clique nas confirmações
            def extrair_info_modal(container, etapa_nome):
                try:
                    texto = ""
                    if hasattr(container, "inner_text"):
                        texto = container.inner_text("body")
                    elif hasattr(container, "locator"):
                        texto = container.locator("body").inner_text()
                    linhas = [l.strip() for l in texto.splitlines() if any(k in l.lower() for k in ["vaga", "concluir", "realmente", "fechar"])]
                    if linhas:
                        log(f"   -> [{etapa_nome}] {linhas[0]}")
                except Exception:
                    pass

            def clicar_com_fallback(el):
                try:
                    el.click(timeout=8000)
                except Exception:
                    try:
                        el.evaluate("e => e.click()")
                    except Exception as e:
                        log(f"   ⚠️ Falha ao clicar: {e}")

            def buscar_e_clicar_sim(etapa_nome, timeout=12):
                fim = time.time() + timeout
                while time.time() < fim:
                    # 1. Se abriu janela popup independente
                    if popup_window is not None and not popup_window.is_closed():
                        try:
                            c = popup_window.locator('input[value="Sim" i], button:has-text("Sim"), a:has-text("Sim"), input[type="submit"][value*="Sim" i]').first
                            if c.count() > 0 and c.is_visible():
                                extrair_info_modal(popup_window, etapa_nome)
                                clicar_com_fallback(c)
                                return True, popup_window
                        except Exception:
                            pass

                    # 2. Varre todos os frames carregados
                    for f in page.frames:
                        if f != page.main_frame:
                            try:
                                c = f.locator('input[value="Sim" i], button:has-text("Sim"), a:has-text("Sim"), input[type="submit"][value*="Sim" i]').first
                                if c.count() > 0 and c.is_visible():
                                    extrair_info_modal(f, etapa_nome)
                                    clicar_com_fallback(c)
                                    return True, f
                            except Exception:
                                pass

                    # 3. Varre usando frame_locator
                    for s_frame in ['iframe[src*="FecharTurma" i]', 'iframe[src*="PopUp" i]', 'iframe']:
                        try:
                            fl = page.frame_locator(s_frame).first
                            c = fl.locator('input[value="Sim" i], button:has-text("Sim"), a:has-text("Sim"), input[type="submit"][value*="Sim" i]').first
                            if c.count() > 0 and c.is_visible():
                                extrair_info_modal(fl, etapa_nome)
                                clicar_com_fallback(c)
                                return True, fl
                        except Exception:
                            pass

                    # 4. Checa diretamente na página principal
                    try:
                        c = page.locator('div[role="dialog"] input[value="Sim" i], div[role="dialog"] button:has-text("Sim"), div[role="dialog"] a:has-text("Sim"), .modal input[value="Sim" i], .modal button:has-text("Sim"), input[value="Sim" i], button:has-text("Sim"), a:has-text("Sim")').first
                        if c.count() > 0 and c.is_visible():
                            extrair_info_modal(page, etapa_nome)
                            clicar_com_fallback(c)
                            return True, page
                    except Exception:
                        pass

                    time.sleep(0.3)

                return False, None

            # 2.1. Primeira confirmação (ex: "Há 997 vaga(s) abertas na Turma...")
            log("   -> Aguardando 1ª caixa de confirmação (vagas abertas)...")
            sucesso_1, alvo_1 = buscar_e_clicar_sim("1ª Confirmação", timeout=15)

            if not sucesso_1:
                log(f"   ❌ [ERRO] Botão 'Sim' da 1ª confirmação não foi encontrado para a turma {turma_id}.")
                logger.finalizar_item(item_identificador, status="Falha", detalhes="Botão Sim da 1ª confirmação não encontrado")
                frames_info = [f.url for f in page.frames]
                log(f"   [DIAGNÓSTICO] Frames detectados na tela: {frames_info}")
                if popup_window and not popup_window.is_closed():
                    popup_window.close()
                continue

            log("   -> 1ª Confirmação realizada com sucesso!")
            page.wait_for_timeout(1000)

            # 2.2. Segunda confirmação (ex: "Deseja realmente fechar a turma?")
            log("   -> Aguardando 2ª caixa de confirmação ('Deseja realmente fechar a turma?')...")
            sucesso_2, alvo_2 = buscar_e_clicar_sim("2ª Confirmação", timeout=8)

            alvo = alvo_2 if alvo_2 is not None else alvo_1

            if sucesso_2:
                log("   -> 2ª Confirmação realizada com sucesso!")
            else:
                log("   -> (Segunda confirmação não solicitada ou já processada).")

            # 4. Aguarda a confirmação de fechamento (caixa verde do sistema)
            log("   -> Aguardando confirmação do sistema...")
            sucesso = False
            tempo_inicio = time.time()
            timeout_max = 180  # até 3 minutos no pior caso

            while time.time() - tempo_inicio < timeout_max:
                # 4.1. Detecção instantânea e infalível da caixa verde na página principal
                try:
                    texto_pagina = page.locator("body").inner_text()
                    if "fechada para realiza" in texto_pagina or "listagem de alunos" in texto_pagina:
                        sucesso = True
                        break
                except Exception:
                    pass

                # 4.2. Detecção com get_by_text oficial
                try:
                    if page.get_by_text("Esta turma foi fechada para realização").count() > 0:
                        sucesso = True
                        break
                except Exception:
                    pass

                # 4.3. Detecção em diálogos nativos (alerts)
                if any("fechada para realiza" in d.lower() or "listagem de alunos" in d.lower() for d in dialogos_recebidos):
                    sucesso = True
                    break

                # 4.4. Detecção em frames remanescentes
                for f in page.frames:
                    try:
                        if "fechada para realiza" in f.locator("body").inner_text():
                            sucesso = True
                            break
                    except Exception:
                        pass
                if sucesso:
                    break

                # 4.5. Se a popup fechou sozinha após salvar
                if popup_window and popup_window.is_closed():
                    sucesso = True
                    break

                time.sleep(0.4)

            if sucesso:
                tempo_total = time.time() - tempo_inicio
                logger.finalizar_item(item_identificador, status="Sucesso", detalhes=f"Confirmado em {tempo_total:.1f}s")
                log(f"   ✅ [SUCESSO] Turma {turma_id} ({nome_curso}) fechada com sucesso! (Confirmado em {tempo_total:.1f}s)")
                total_fechadas += 1

                # Pausa de 10 segundos solicitada para verificação das informações na tela
                log("   ⏳ Aguardando 10 segundos para conferência das informações na tela...")
                page.wait_for_timeout(10000)

                # Fecha a notificação verde se houver botão 'x' para deixar a tela limpa
                try:
                    btn_fechar_verde = page.locator('.alert .close, .alert a:has-text("x"), .alert button, a:has-text("✖")').first
                    if btn_fechar_verde.count() > 0 and btn_fechar_verde.is_visible():
                        btn_fechar_verde.click()
                except Exception:
                    pass

                if popup_window and not popup_window.is_closed():
                    try:
                        popup_window.close()
                    except Exception:
                        pass
            else:
                logger.finalizar_item(item_identificador, status="Falha", detalhes="Tempo limite excedido aguardando confirmação")
                log(f"   ⚠️ [ATENÇÃO] Tempo limite atingido ao aguardar confirmação da turma {turma_id}.")

            page.wait_for_timeout(1000)

            # Atualiza a pesquisa de turmas abertas para a próxima iteração
            if idx < len(df_turmas):
                log("   -> Atualizando tabela de turmas para o próximo item...")
                try:
                    aplicar_pesquisa_aberta()
                except Exception as e:
                    log(f"   ⚠️ Aviso ao atualizar pesquisa: {e}")

        status_geral = f"{total_fechadas} de {len(df_turmas)} turmas fechadas com sucesso"
        logger.finalizar(status_geral=status_geral)

        log("\n" + "=" * 65)
        log(f">> CONCLUÍDO! Total de {total_fechadas} de {len(df_turmas)} turmas fechadas com sucesso.")
        log("=" * 65)

        page.wait_for_timeout(3000)
        browser.close()

if __name__ == "__main__":
    executar_fechamento()
