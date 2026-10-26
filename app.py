"""
Central de Automações EAD - EGOV (Escola de Governo do DF)
Interface Gráfica Moderna (Desktop GUI) desenvolvida com CustomTkinter.

Funcionalidades:
1. Download de Inscritos do Portal de Sistemas GDF
2. Tratamento dos arquivos CSV para o Moodle
3. Download de Notas do Egovvirtual (Moodle)
4. Formatação e Avaliação Final de Notas (SEDOC)
"""

import os
import sys
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from tkcalendar import DateEntry

# Configurações visuais do CustomTkinter
ctk.set_appearance_mode("System")  # Segue o tema do Windows (Dark/Light)
ctk.set_default_color_theme("blue")

DIRETORIO_APP = os.path.dirname(os.path.abspath(__file__))

class TextRedirector:
    """Redireciona o stdout/print para uma caixa de texto do CustomTkinter"""
    def __init__(self, textbox):
        self.textbox = textbox
        self.encoding = "utf-8"

    def write(self, text):
        def append():
            try:
                self.textbox.configure(state="normal")
                self.textbox.insert("end", text)
                self.textbox.see("end")
                self.textbox.configure(state="disabled")
            except Exception:
                pass
        try:
            self.textbox.after(0, append)
        except Exception:
            pass

    def flush(self):
        pass

    def isatty(self):
        return False

    def reconfigure(self, **kwargs):
        if "encoding" in kwargs:
            self.encoding = kwargs["encoding"]

class AppEGOV(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Central de Automações EAD - Escola de Governo do DF")
        self.geometry("980 x 740")
        self.minsize(850, 650)
        self.executando = False

        # Configura layout de grid principal (2 linhas: Conteúdo e Terminal de Logs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self._criar_abas()
        self._criar_area_logs()

        # Redireciona prints do sistema para a caixa de logs
        self.redirector = TextRedirector(self.log_textbox)
        sys.stdout = self.redirector

        self.log("🚀 Central de Automações da EGOV carregada com sucesso!")
        self.log("Selecione uma das etapas acima e clique no botão correspondente para executar.\n")

    def log(self, mensagem):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", mensagem + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def _criar_abas(self):
        # Abas principais no topo
        self.tabview = ctk.CTkTabview(self, corner_radius=10)
        self.tabview.grid(row=0, column=0, padx=20, pady=(15, 10), sticky="nsew")

        self.tab_inscritos = self.tabview.add("1. Inscritos GDF")
        self.tab_tratar = self.tabview.add("2. Tratar p/ Moodle")
        self.tab_notas = self.tabview.add("3. Notas Moodle")
        self.tab_sedoc = self.tabview.add("4. Formatar SEDOC")
        self.tab_mover = self.tabview.add("5. Mover Arquivos Finais")
        self.tab_fechar = self.tabview.add("6. Fechar Turmas")

        self._montar_aba_inscritos()
        self._montar_aba_tratar()
        self._montar_aba_notas()
        self._montar_aba_sedoc()
        self._montar_aba_mover()
        self._montar_aba_fechar()

    # =========================================================================
    # ABA 1: BAIXAR INSCRITOS DO GDF
    # =========================================================================
    def _montar_aba_inscritos(self):
        tab = self.tab_inscritos
        
        lbl_titulo = ctk.CTkLabel(
            tab, text="📥 Download de Inscritos no Portal de Sistemas (EGOV)", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), anchor="w")

        lbl_desc = ctk.CTkLabel(
            tab, 
            text="Faz login seguro no Portal de Sistemas do GDF, aplica os filtros (EAD / Fechada) e baixa\n"
                 "os CSVs dos cursos listados em 'EAD_Controle_Geral.xlsx', salvando na pasta 'Arquivos_Originais_Sistema/'.",
            text_color="gray", justify="left"
        )
        lbl_desc.pack(pady=(0, 15), anchor="w")

        # Frame de credenciais
        frame_cred = ctk.CTkFrame(tab)
        frame_cred.pack(fill="x", padx=10, pady=10)

        # Usuário
        ctk.CTkLabel(frame_cred, text="Usuário GDF:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        self.ent_user_gdf = ctk.CTkEntry(frame_cred, width=280, placeholder_text="ex: jonathan.lima")
        self.ent_user_gdf.grid(row=0, column=1, padx=15, pady=10, sticky="w")

        # Senha
        ctk.CTkLabel(frame_cred, text="Senha GDF:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.ent_pass_gdf = ctk.CTkEntry(frame_cred, width=280, show="•", placeholder_text="Sua senha corporativa")
        self.ent_pass_gdf.grid(row=1, column=1, padx=15, pady=10, sticky="w")

        # Checkbox mostrar senha
        self.chk_show_pass_gdf = ctk.CTkCheckBox(frame_cred, text="Mostrar senha", command=self._toggle_pass_gdf)
        self.chk_show_pass_gdf.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # Órgão
        ctk.CTkLabel(frame_cred, text="Órgão:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=15, pady=10, sticky="w")
        self.ent_orgao_gdf = ctk.CTkEntry(frame_cred, width=150)
        self.ent_orgao_gdf.insert(0, "SEEC")
        self.ent_orgao_gdf.grid(row=2, column=1, padx=15, pady=10, sticky="w")

        # Botões de Ação
        frame_botoes = ctk.CTkFrame(tab, fg_color="transparent")
        frame_botoes.pack(fill="x", padx=10, pady=(20, 10))

        self.btn_exec_inscritos = ctk.CTkButton(
            frame_botoes, text="▶ Iniciar Download de Inscritos", 
            font=ctk.CTkFont(size=14, weight="bold"), height=40,
            command=self._iniciar_thread_inscritos
        )
        self.btn_exec_inscritos.pack(side="left", padx=(0, 15))

        btn_abrir_originais = ctk.CTkButton(
            frame_botoes, text="📂 Abrir Arquivos_Originais_Sistema", 
            fg_color="#4A5568", hover_color="#2D3748", height=40,
            command=lambda: self._abrir_pasta("Arquivos_Originais_Sistema")
        )
        btn_abrir_originais.pack(side="left")

    def _toggle_pass_gdf(self):
        if self.chk_show_pass_gdf.get():
            self.ent_pass_gdf.configure(show="")
        else:
            self.ent_pass_gdf.configure(show="•")

    # =========================================================================
    # ABA 2: TRATAR CSVs PARA MOODLE
    # =========================================================================
    def _montar_aba_tratar(self):
        tab = self.tab_tratar

        lbl_titulo = ctk.CTkLabel(
            tab, text="🛠️ Tratamento e Padronização de CSV para o Moodle", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), anchor="w")

        lbl_desc = ctk.CTkLabel(
            tab, 
            text="Formata os arquivos CSV da pasta 'Arquivos_Originais_Sistema/':\n"
                 "• Padroniza Nomes e Sobrenomes (Primeira Letra Maiúscula);\n"
                 "• Formata o CPF (username) com 11 dígitos e zeros à esquerda;\n"
                 "• Converte e-mails para minúsculas;\n"
                 "• Salva na pasta 'Arquivos_CARREGAR/' com o sufixo '... - CARREGAR.csv'.",
            text_color="gray", justify="left"
        )
        lbl_desc.pack(pady=(0, 20), anchor="w")

        frame_botoes = ctk.CTkFrame(tab, fg_color="transparent")
        frame_botoes.pack(fill="x", padx=10, pady=10)

        self.btn_exec_tratar = ctk.CTkButton(
            frame_botoes, text="▶ Executar Tratamento dos CSVs", 
            font=ctk.CTkFont(size=14, weight="bold"), height=40,
            command=self._iniciar_thread_tratar
        )
        self.btn_exec_tratar.pack(side="left", padx=(0, 15))

        btn_abrir_carregar = ctk.CTkButton(
            frame_botoes, text="📂 Abrir Pasta Arquivos_CARREGAR", 
            fg_color="#4A5568", hover_color="#2D3748", height=40,
            command=lambda: self._abrir_pasta("Arquivos_CARREGAR")
        )
        btn_abrir_carregar.pack(side="left")

    # =========================================================================
    # ABA 3: BAIXAR NOTAS DO MOODLE
    # =========================================================================
    def _montar_aba_notas(self):
        tab = self.tab_notas

        lbl_titulo = ctk.CTkLabel(
            tab, text="🎓 Download de Notas no Egovvirtual (Moodle)", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), anchor="w")

        lbl_desc = ctk.CTkLabel(
            tab, 
            text="Faz login automático no Egovvirtual e baixa as planilhas de notas brutas (.xlsx)\n"
                 "de cada turma configurada na planilha 'EAD_Controle_Geral.xlsx', salvando em 'Arquivos_Originais_Moodle/'.",
            text_color="gray", justify="left"
        )
        lbl_desc.pack(pady=(0, 15), anchor="w")

        # Frame de credenciais do Moodle
        frame_cred = ctk.CTkFrame(tab)
        frame_cred.pack(fill="x", padx=10, pady=10)

        # Usuário Moodle
        ctk.CTkLabel(frame_cred, text="Usuário Moodle:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        self.ent_user_moodle = ctk.CTkEntry(frame_cred, width=280, placeholder_text="ex: jonathan")
        self.ent_user_moodle.grid(row=0, column=1, padx=15, pady=10, sticky="w")

        # Senha Moodle
        ctk.CTkLabel(frame_cred, text="Senha Moodle:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.ent_pass_moodle = ctk.CTkEntry(frame_cred, width=280, show="•", placeholder_text="Sua senha do Egovvirtual")
        self.ent_pass_moodle.grid(row=1, column=1, padx=15, pady=10, sticky="w")

        # Checkbox mostrar senha
        self.chk_show_pass_moodle = ctk.CTkCheckBox(frame_cred, text="Mostrar senha", command=self._toggle_pass_moodle)
        self.chk_show_pass_moodle.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # Botão de Ação
        frame_botoes = ctk.CTkFrame(tab, fg_color="transparent")
        frame_botoes.pack(fill="x", padx=10, pady=(20, 10))

        self.btn_exec_notas = ctk.CTkButton(
            frame_botoes, text="▶ Iniciar Download de Notas", 
            font=ctk.CTkFont(size=14, weight="bold"), height=40,
            command=self._iniciar_thread_notas
        )
        self.btn_exec_notas.pack(side="left", padx=(0, 15))

        btn_abrir_moodle_orig = ctk.CTkButton(
            frame_botoes, text="📂 Abrir Arquivos_Originais_Moodle", 
            fg_color="#4A5568", hover_color="#2D3748", height=40,
            command=lambda: self._abrir_pasta("Arquivos_Originais_Moodle")
        )
        btn_abrir_moodle_orig.pack(side="left")

    def _toggle_pass_moodle(self):
        if self.chk_show_pass_moodle.get():
            self.ent_pass_moodle.configure(show="")
        else:
            self.ent_pass_moodle.configure(show="•")

    # =========================================================================
    # ABA 4: FORMATAR NOTAS SEDOC
    # =========================================================================
    def _montar_aba_sedoc(self):
        tab = self.tab_sedoc

        lbl_titulo = ctk.CTkLabel(
            tab, text="📊 Avaliação e Formatação Final de Notas (SEDOC)", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), anchor="w")

        lbl_desc = ctk.CTkLabel(
            tab, 
            text="Lê as notas da pasta 'Arquivos_Originais_Moodle/', aplica a regra de aprovação (60% RPIM / 70% demais),\n"
                 "constrói a tabela oficial com base legal e salva em 'Arquivos_Finais_SEDOC/' com sufixo '_SEDOC.xlsx'.",
            text_color="gray", justify="left"
        )
        lbl_desc.pack(pady=(0, 15), anchor="w")

        # Frame dos Calendários de Seleção de Data
        frame_input = ctk.CTkFrame(tab)
        frame_input.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(
            frame_input, text="Período de Realização do Curso:", 
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, columnspan=4, padx=15, pady=(12, 10), sticky="w")

        # 1. Data Início
        ctk.CTkLabel(frame_input, text="Data Início:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=(15, 8), pady=10, sticky="w")
        self.cal_inicio = DateEntry(
            frame_input, width=13, font=("Verdana", 10),
            background="#1f538d", foreground="white",
            headersbackground="#14375e", headersforeground="white",
            selectbackground="#3b8ed0", selectforeground="white",
            borderwidth=2, date_pattern="dd/mm/yyyy", locale="pt_BR"
        )
        self.cal_inicio.grid(row=1, column=1, padx=(0, 25), pady=10, sticky="w")

        # 2. Data Fim
        ctk.CTkLabel(frame_input, text="Data Fim:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=2, padx=(10, 8), pady=10, sticky="w")
        self.cal_fim = DateEntry(
            frame_input, width=13, font=("Verdana", 10),
            background="#1f538d", foreground="white",
            headersbackground="#14375e", headersforeground="white",
            selectbackground="#3b8ed0", selectforeground="white",
            borderwidth=2, date_pattern="dd/mm/yyyy", locale="pt_BR"
        )
        self.cal_fim.grid(row=1, column=3, padx=(0, 15), pady=10, sticky="w")

        # 3. Preview do formato oficial gerado
        self.lbl_preview_data = ctk.CTkLabel(
            frame_input, 
            text=f"📅 Período gerado para a planilha: {self.cal_inicio.get()} a {self.cal_fim.get()}",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#3B8ED0"
        )
        self.lbl_preview_data.grid(row=2, column=0, columnspan=4, padx=15, pady=(5, 12), sticky="w")

        def _atualizar_preview(event=None):
            self.lbl_preview_data.configure(
                text=f"📅 Período gerado para a planilha: {self.cal_inicio.get()} a {self.cal_fim.get()}"
            )

        self.cal_inicio.bind("<<DateEntrySelected>>", _atualizar_preview)
        self.cal_fim.bind("<<DateEntrySelected>>", _atualizar_preview)

        # Botões de Ação
        frame_botoes = ctk.CTkFrame(tab, fg_color="transparent")
        frame_botoes.pack(fill="x", padx=10, pady=(20, 10))

        self.btn_exec_sedoc = ctk.CTkButton(
            frame_botoes, text="▶ Gerar Planilhas SEDOC", 
            font=ctk.CTkFont(size=14, weight="bold"), height=40,
            command=self._iniciar_thread_sedoc
        )
        self.btn_exec_sedoc.pack(side="left", padx=(0, 15))

        btn_abrir_sedoc = ctk.CTkButton(
            frame_botoes, text="📂 Abrir Pasta Arquivos_Finais_SEDOC", 
            fg_color="#4A5568", hover_color="#2D3748", height=40,
            command=lambda: self._abrir_pasta("Arquivos_Finais_SEDOC")
        )
        btn_abrir_sedoc.pack(side="left")

    # =========================================================================
    # ABA 5: MOVER / COPIAR ARQUIVOS FINAIS PARA O SERVIDOR
    # =========================================================================
    def _montar_aba_mover(self):
        tab = self.tab_mover

        lbl_titulo = ctk.CTkLabel(
            tab, text="📁 Copiar Arquivos Finais para o Servidor de Arquivos", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), anchor="w")

        lbl_desc = ctk.CTkLabel(
            tab, 
            text="Organiza e cria uma cópia dos arquivos finais gerados nas suas respectivas pastas no servidor de rede.\n"
                 "• Origens locais: 'Arquivos_CARREGAR/' (.csv) e 'Arquivos_Finais_SEDOC/' (.xlsx);\n"
                 "• Identifica automaticamente a Sigla do Curso e a Turma (ex: PROAMIS_AFG -> #PROAMIS\\AFG\\Turma 06.2026);\n"
                 "• Seus arquivos locais continuam preservados no computador;\n"
                 "• Se o arquivo já existir no destino, você será avisado para autorizar ou não a substituição.",
            text_color="gray", justify="left"
        )
        lbl_desc.pack(pady=(0, 15), anchor="w")

        # Frame com o caminho de destino no servidor
        frame_caminho = ctk.CTkFrame(tab)
        frame_caminho.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(
            frame_caminho, text="📍 Diretório de Destino no Servidor:", 
            font=ctk.CTkFont(weight="bold")
        ).grid(row=0, column=0, padx=15, pady=(12, 5), sticky="w")

        self.ent_caminho_servidor = ctk.CTkEntry(frame_caminho, width=680, font=ctk.CTkFont(size=12))
        self.ent_caminho_servidor.insert(0, r"D:\SEGEA\RECENTES\EGOV\CEAD\Cursos Autoinstrucionais\# 2026")
        self.ent_caminho_servidor.grid(row=1, column=0, padx=15, pady=(0, 12), sticky="ew")

        # Botões de Ação
        frame_botoes = ctk.CTkFrame(tab, fg_color="transparent")
        frame_botoes.pack(fill="x", padx=10, pady=(20, 10))

        self.btn_exec_mover = ctk.CTkButton(
            frame_botoes, text="▶ Copiar Arquivos para o Servidor", 
            font=ctk.CTkFont(size=14, weight="bold"), height=40,
            command=self._iniciar_thread_mover
        )
        self.btn_exec_mover.pack(side="left", padx=(0, 15))

        btn_abrir_serv = ctk.CTkButton(
            frame_botoes, text="📂 Abrir Pasta no Servidor (D:)", 
            fg_color="#4A5568", hover_color="#2D3748", height=40,
            command=self._abrir_pasta_servidor
        )
        btn_abrir_serv.pack(side="left")

    # =========================================================================
    # ABA 6: FECHAR TURMAS NO SISTEMA (EGOV)
    # =========================================================================
    def _montar_aba_fechar(self):
        tab = self.tab_fechar

        lbl_titulo = ctk.CTkLabel(
            tab, text="🔒 Fechamento Automático de Turmas no Sistema (EGOV)", 
            font=ctk.CTkFont(size=18, weight="bold")
        )
        lbl_titulo.pack(pady=(15, 5), anchor="w")

        lbl_desc = ctk.CTkLabel(
            tab, 
            text="Faz login no Portal de Sistemas do GDF, aplica os filtros (Modalidade: EAD | Situação: Aberta) e fecha\n"
                 "as turmas cadastradas na planilha 'EAD_Controle_Geral.xlsx', confirmando o fechamento de cada uma no sistema.",
            text_color="gray", justify="left"
        )
        lbl_desc.pack(pady=(0, 15), anchor="w")

        # Frame de credenciais
        frame_cred = ctk.CTkFrame(tab)
        frame_cred.pack(fill="x", padx=10, pady=10)

        # Usuário
        ctk.CTkLabel(frame_cred, text="Usuário GDF:", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=15, pady=10, sticky="w")
        self.ent_user_fechar = ctk.CTkEntry(frame_cred, width=280, placeholder_text="ex: jonathan.lima")
        self.ent_user_fechar.grid(row=0, column=1, padx=15, pady=10, sticky="w")

        # Senha
        ctk.CTkLabel(frame_cred, text="Senha GDF:", font=ctk.CTkFont(weight="bold")).grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.ent_pass_fechar = ctk.CTkEntry(frame_cred, width=280, show="•", placeholder_text="Sua senha corporativa")
        self.ent_pass_fechar.grid(row=1, column=1, padx=15, pady=10, sticky="w")

        # Checkbox mostrar senha
        self.chk_show_pass_fechar = ctk.CTkCheckBox(frame_cred, text="Mostrar senha", command=self._toggle_pass_fechar)
        self.chk_show_pass_fechar.grid(row=1, column=2, padx=10, pady=10, sticky="w")

        # Órgão
        ctk.CTkLabel(frame_cred, text="Órgão:", font=ctk.CTkFont(weight="bold")).grid(row=2, column=0, padx=15, pady=10, sticky="w")
        self.ent_orgao_fechar = ctk.CTkEntry(frame_cred, width=150)
        self.ent_orgao_fechar.insert(0, "SEEC")
        self.ent_orgao_fechar.grid(row=2, column=1, padx=15, pady=10, sticky="w")

        # Botões de Ação
        frame_botoes = ctk.CTkFrame(tab, fg_color="transparent")
        frame_botoes.pack(fill="x", padx=10, pady=(20, 10))

        self.btn_exec_fechar = ctk.CTkButton(
            frame_botoes, text="▶ Iniciar Fechamento de Turmas", 
            font=ctk.CTkFont(size=14, weight="bold"), height=40,
            command=self._iniciar_thread_fechar
        )
        self.btn_exec_fechar.pack(side="left", padx=(0, 15))

        btn_abrir_planilha = ctk.CTkButton(
            frame_botoes, text="📂 Abrir Planilha EAD_Controle_Geral", 
            fg_color="#4A5568", hover_color="#2D3748", height=40,
            command=self._abrir_planilha_controle
        )
        btn_abrir_planilha.pack(side="left")

    def _toggle_pass_fechar(self):
        if self.chk_show_pass_fechar.get():
            self.ent_pass_fechar.configure(show="")
        else:
            self.ent_pass_fechar.configure(show="•")

    def _abrir_planilha_controle(self):
        caminho = os.path.join(DIRETORIO_APP, "EAD_Controle_Geral.xlsx")
        if not os.path.exists(caminho):
            caminho = os.path.join(DIRETORIO_APP, "EAD_Controle_Geral - Copia.xlsx")
        if os.path.exists(caminho):
            try:
                os.startfile(caminho)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")
        else:
            messagebox.showwarning("Não Encontrado", "Planilha EAD_Controle_Geral.xlsx não encontrada na pasta.")

    # =========================================================================
    # ÁREA DE LOGS / CONSOLE EM TEMPO REAL
    # =========================================================================
    def _criar_area_logs(self):
        frame_logs = ctk.CTkFrame(self, corner_radius=10)
        frame_logs.grid(row=1, column=0, padx=20, pady=(0, 15), sticky="nsew")
        frame_logs.grid_columnconfigure(0, weight=1)

        # Cabeçalho dos logs
        header = ctk.CTkFrame(frame_logs, fg_color="transparent")
        header.pack(fill="x", padx=15, pady=(8, 4))

        lbl_log = ctk.CTkLabel(header, text="💻 Console de Execução (Tempo Real)", font=ctk.CTkFont(weight="bold"))
        lbl_log.pack(side="left")

        btn_limpar = ctk.CTkButton(header, text="Limpar Logs", width=90, height=24, font=ctk.CTkFont(size=12), command=self._limpar_logs)
        btn_limpar.pack(side="right")

        # Caixa de texto para logs
        self.log_textbox = ctk.CTkTextbox(
            frame_logs, height=210, font=ctk.CTkFont(family="Consolas", size=12), 
            corner_radius=8, state="disabled"
        )
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _limpar_logs(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def _abrir_pasta(self, nome_pasta):
        caminho = os.path.join(DIRETORIO_APP, nome_pasta)
        os.makedirs(caminho, exist_ok=True)
        try:
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a pasta:\n{e}")

    # =========================================================================
    # CONTROLE DE THREADS (RODA AS AUTOMAÇÕES SEM TRAVAR A JANELA)
    # =========================================================================
    def _esta_ocupado(self):
        if getattr(self, 'executando', False):
            messagebox.showwarning("Em Andamento", "Já existe uma automação em execução. Aguarde a finalização antes de iniciar outra!")
            return True
        return False

    def _finalizar_tarefa(self, botao, texto_original):
        def reset():
            self.executando = False
            botao.configure(text=texto_original)
        self.after(0, reset)

    # 1. Thread Inscritos
    def _iniciar_thread_inscritos(self):
        if self._esta_ocupado():
            return

        usuario = self.ent_user_gdf.get().strip()
        senha = self.ent_pass_gdf.get().strip()
        orgao = self.ent_orgao_gdf.get().strip() or "SEEC"

        if not usuario or not senha:
            messagebox.showwarning("Campos Obrigatórios", "Por favor, preencha o Usuário e a Senha do GDF.")
            return

        self.executando = True
        self.btn_exec_inscritos.configure(text="⏳ Executando...")
        threading.Thread(target=self._exec_inscritos, args=(usuario, senha, orgao), daemon=True).start()

    def _exec_inscritos(self, usuario, senha, orgao):
        try:
            from automacao import executar_automacao_gui
            executar_automacao_gui(usuario, senha, orgao, self.log)
        except Exception as e:
            self.log(f"\n❌ Erro durante o processo: {e}")
        finally:
            self._finalizar_tarefa(self.btn_exec_inscritos, "▶ Iniciar Download de Inscritos")

    # 2. Thread Tratar
    def _iniciar_thread_tratar(self):
        if self._esta_ocupado():
            return

        self.executando = True
        self.btn_exec_tratar.configure(text="⏳ Processando...")
        threading.Thread(target=self._exec_tratar, daemon=True).start()

    def _exec_tratar(self):
        try:
            from tratar_csv_moodle import tratar_arquivos_csv
            tratar_arquivos_csv()
        except Exception as e:
            self.log(f"\n❌ Erro durante o tratamento: {e}")
        finally:
            self._finalizar_tarefa(self.btn_exec_tratar, "▶ Executar Tratamento dos CSVs")

    # 3. Thread Notas
    def _iniciar_thread_notas(self):
        if self._esta_ocupado():
            return

        usuario = self.ent_user_moodle.get().strip()
        senha = self.ent_pass_moodle.get().strip()

        if not usuario or not senha:
            messagebox.showwarning("Campos Obrigatórios", "Por favor, preencha o Usuário e a Senha do Moodle.")
            return

        self.executando = True
        self.btn_exec_notas.configure(text="⏳ Baixando Notas...")
        threading.Thread(target=self._exec_notas, args=(usuario, senha), daemon=True).start()

    def _exec_notas(self, usuario, senha):
        try:
            from baixar_notas_moodle import baixar_notas_gui
            baixar_notas_gui(usuario, senha, self.log)
        except Exception as e:
            self.log(f"\n❌ Erro ao baixar notas: {e}")
        finally:
            self._finalizar_tarefa(self.btn_exec_notas, "▶ Iniciar Download de Notas")

    # 4. Thread SEDOC
    def _iniciar_thread_sedoc(self):
        if self._esta_ocupado():
            return

        data_ini = self.cal_inicio.get().strip()
        data_fim = self.cal_fim.get().strip()

        if not data_ini or not data_fim:
            messagebox.showwarning("Campos Obrigatórios", "Por favor, selecione as datas de início e fim do curso.")
            return

        periodo = f"{data_ini} a {data_fim}"

        self.executando = True
        self.btn_exec_sedoc.configure(text="⏳ Gerando SEDOC...")
        threading.Thread(target=self._exec_sedoc, args=(periodo,), daemon=True).start()

    def _exec_sedoc(self, periodo):
        try:
            from formatar_notas_moodle import formatar_notas_gui
            formatar_notas_gui(periodo, self.log)
        except Exception as e:
            self.log(f"\n❌ Erro ao formatar SEDOC: {e}")
        finally:
            self._finalizar_tarefa(self.btn_exec_sedoc, "▶ Gerar Planilhas SEDOC")

    # 5. Thread Mover / Copiar Arquivos para o Servidor
    def _iniciar_thread_mover(self):
        if self._esta_ocupado():
            return

        caminho_dest = self.ent_caminho_servidor.get().strip()
        if not caminho_dest:
            messagebox.showwarning("Caminho Obrigatório", "Por favor, informe o caminho de destino no servidor.")
            return

        self.executando = True
        self.btn_exec_mover.configure(text="⏳ Copiando para Servidor...")
        threading.Thread(target=self._exec_mover, args=(caminho_dest,), daemon=True).start()

    def _exec_mover(self, caminho_dest):
        try:
            from copiar_para_servidor import copiar_arquivos_para_servidor
            copiar_arquivos_para_servidor(
                caminho_servidor=caminho_dest,
                callback_confirmar=self._pedir_confirmacao_sobrescrita,
                log_fn=self.log
            )
        except Exception as e:
            self.log(f"\n❌ Erro durante a cópia para o servidor: {e}")
        finally:
            self._finalizar_tarefa(self.btn_exec_mover, "▶ Copiar Arquivos para o Servidor")

    def _abrir_pasta_servidor(self):
        caminho = self.ent_caminho_servidor.get().strip()
        if not os.path.exists(caminho):
            messagebox.showwarning("Pasta Inacessível", f"O caminho do servidor não está acessível no momento:\n{caminho}")
            return
        try:
            os.startfile(caminho)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o diretório do servidor:\n{e}")

    # 6. Thread Fechar Turmas GDF
    def _iniciar_thread_fechar(self):
        if self._esta_ocupado():
            return

        usuario = self.ent_user_fechar.get().strip() or self.ent_user_gdf.get().strip()
        senha = self.ent_pass_fechar.get().strip() or self.ent_pass_gdf.get().strip()
        orgao = self.ent_orgao_fechar.get().strip() or self.ent_orgao_gdf.get().strip() or "SEEC"

        if not usuario or not senha:
            messagebox.showwarning("Campos Obrigatórios", "Por favor, preencha o Usuário e a Senha do GDF.")
            return

        # Sincroniza campos se um deles estava preenchido
        if not self.ent_user_fechar.get().strip():
            self.ent_user_fechar.insert(0, usuario)
        if not self.ent_pass_fechar.get().strip():
            self.ent_pass_fechar.insert(0, senha)

        self.executando = True
        self.btn_exec_fechar.configure(text="⏳ Fechando Turmas...")
        threading.Thread(target=self._exec_fechar, args=(usuario, senha, orgao), daemon=True).start()

    def _exec_fechar(self, usuario, senha, orgao):
        try:
            from fechar_turmas_sistema import fechar_turmas_gui
            fechar_turmas_gui(usuario, senha, orgao, self.log)
        except Exception as e:
            self.log(f"\n❌ Erro durante o fechamento de turmas: {e}")
        finally:
            self._finalizar_tarefa(self.btn_exec_fechar, "▶ Iniciar Fechamento de Turmas")

    def _pedir_confirmacao_sobrescrita(self, nome_arquivo, pasta_destino):
        """
        Exibe janela modal no CustomTkinter para decidir sobre a substituição de arquivos existentes.
        Retorna: 'sim', 'todos', 'nao', 'pular_todos', 'cancelar'
        """
        resultado = ["nao"]
        evento = threading.Event()

        def mostrar_dialogo():
            janela = ctk.CTkToplevel(self)
            janela.title("Aviso: Arquivo Já Existe")
            janela.geometry("540x260")
            janela.resizable(False, False)
            janela.transient(self)
            janela.grab_set()

            def fechar(resp):
                resultado[0] = resp
                try:
                    janela.grab_release()
                    janela.destroy()
                except Exception:
                    pass
                evento.set()

            janela.protocol("WM_DELETE_WINDOW", lambda: fechar("cancelar"))

            lbl_msg = ctk.CTkLabel(
                janela, 
                text="⚠️ O arquivo já existe na pasta de destino!", 
                font=ctk.CTkFont(size=15, weight="bold"),
                text_color="#E53E3E"
            )
            lbl_msg.pack(pady=(15, 6), padx=20, anchor="w")

            lbl_detalhes = ctk.CTkLabel(
                janela, 
                text=f"Arquivo:\n{nome_arquivo}\n\nDestino:\n{pasta_destino}", 
                font=ctk.CTkFont(size=12),
                justify="left",
                wraplength=500
            )
            lbl_detalhes.pack(pady=(0, 12), padx=20, anchor="w")

            lbl_pergunta = ctk.CTkLabel(
                janela, 
                text="Deseja substituir o arquivo existente?", 
                font=ctk.CTkFont(weight="bold")
            )
            lbl_pergunta.pack(pady=(0, 15), padx=20, anchor="w")

            frame_btn = ctk.CTkFrame(janela, fg_color="transparent")
            frame_btn.pack(fill="x", padx=15, pady=(0, 10))

            btn_sim = ctk.CTkButton(frame_btn, text="Sim", width=75, command=lambda: fechar("sim"))
            btn_sim.pack(side="left", padx=4)

            btn_todos = ctk.CTkButton(frame_btn, text="Sim para Todos", width=110, command=lambda: fechar("todos"))
            btn_todos.pack(side="left", padx=4)

            btn_nao = ctk.CTkButton(frame_btn, text="Não", width=75, fg_color="#718096", hover_color="#4A5568", command=lambda: fechar("nao"))
            btn_nao.pack(side="left", padx=4)

            btn_pular_todos = ctk.CTkButton(frame_btn, text="Pular Todos", width=95, fg_color="#718096", hover_color="#4A5568", command=lambda: fechar("pular_todos"))
            btn_pular_todos.pack(side="left", padx=4)

            btn_cancelar = ctk.CTkButton(frame_btn, text="Cancelar", width=80, fg_color="#E53E3E", hover_color="#C53030", command=lambda: fechar("cancelar"))
            btn_cancelar.pack(side="left", padx=4)

            janela.update_idletasks()
            x = self.winfo_x() + (self.winfo_width() // 2) - (janela.winfo_width() // 2)
            y = self.winfo_y() + (self.winfo_height() // 2) - (janela.winfo_height() // 2)
            janela.geometry(f"+{x}+{y}")
            janela.focus()

        self.after(0, mostrar_dialogo)
        evento.wait()
        return resultado[0]

if __name__ == "__main__":
    app = AppEGOV()
    app.mainloop()
