"""
Módulo Utilitário de Logs para a Central de Automações EGOV

Funcionalidades:
1. Gravação automática de arquivos de log em formato .txt na pasta 'logs/'.
2. Registro detalhado de:
   - Nome da automação / macro
   - Usuário do sistema e usuário da máquina Windows
   - Data e hora de início e término
   - Cronometragem individual de cada etapa / turma / arquivo
   - Resumo estatístico e tempo total de execução
3. Suporte simultâneo a exibição no console visual do CustomTkinter e no terminal.
"""

import os
import sys
import time
import getpass
from datetime import datetime

DIRETORIO_APP = os.path.dirname(os.path.abspath(__file__))
PASTA_LOGS = os.path.join(DIRETORIO_APP, "logs")

def formatar_duracao(segundos):
    """Converte segundos em formato amigável (ex: '01 min 23 seg (83.5s)' ou '14.2s')"""
    if segundos is None:
        return "N/A"
    if segundos < 60:
        return f"{segundos:.1f}s"
    minutos = int(segundos // 60)
    segs = segundos % 60
    return f"{minutos:02d} min {segs:04.1f} seg ({segundos:.1f}s)"

class AutoLogger:
    def __init__(self, nome_macro, usuario=None, callback_console=print):
        self.nome_macro = nome_macro
        self.usuario_sistema = usuario or getpass.getuser()
        self.usuario_windows = getpass.getuser()
        self.callback_console = callback_console
        
        # Garante que a pasta 'logs' exista
        os.makedirs(PASTA_LOGS, exist_ok=True)
        
        self.inicio_dt = datetime.now()
        self.tempo_inicio = time.time()
        
        # Nome do arquivo de log: log_<nome_macro>_YYYY-MM-DD_HH-MM-SS.txt
        slug_macro = (
            nome_macro.lower()
            .replace(" ", "_")
            .replace("/", "_")
            .replace("\\", "_")
            .replace(":", "")
        )
        timestamp_str = self.inicio_dt.strftime("%Y-%m-%d_%H-%M-%S")
        self.nome_arquivo = f"log_{slug_macro}_{timestamp_str}.txt"
        self.caminho_arquivo = os.path.join(PASTA_LOGS, self.nome_arquivo)
        
        # Dicionários de cronometragem
        self.cronometros = {}
        self.historico_etapas = []  # Lista de registros individuais
        
        # Escreve o cabeçalho no arquivo de log
        self._escrever_cabecalho()

    def _escrever_cabecalho(self):
        cabecalho = [
            "=" * 85,
            "RELATÓRIO DE EXECUÇÃO - CENTRAL DE AUTOMAÇÕES EGOV",
            "=" * 85,
            f"Processo / Automação : {self.nome_macro}",
            f"Usuário do Sistema   : {self.usuario_sistema}",
            f"Usuário do Windows   : {self.usuario_windows}",
            f"Início da Execução   : {self.inicio_dt.strftime('%d/%m/%Y às %H:%M:%S')}",
            f"Arquivo de Log       : {self.caminho_arquivo}",
            "=" * 85,
            ""
        ]
        with open(self.caminho_arquivo, "w", encoding="utf-8") as f:
            f.write("\n".join(cabecalho) + "\n")

    def log(self, mensagem=""):
        """Exibe a mensagem no console e registra no arquivo .txt"""
        if self.callback_console:
            try:
                self.callback_console(mensagem)
            except Exception:
                pass
            
        hora_atual = datetime.now().strftime("%H:%M:%S")
        with open(self.caminho_arquivo, "a", encoding="utf-8") as f:
            if not mensagem or mensagem.startswith("="):
                f.write(f"{mensagem}\n")
            else:
                f.write(f"[{hora_atual}] {mensagem}\n")

    def iniciar_item(self, identificador):
        """Inicia a cronometragem para uma turma ou etapa específica"""
        self.cronometros[identificador] = time.time()

    def finalizar_item(self, identificador, status="Sucesso", detalhes=""):
        """Finaliza a cronometragem para uma turma ou etapa e guarda o registro"""
        t_inicio = self.cronometros.get(identificador, time.time())
        duracao = time.time() - t_inicio
        registro = {
            'identificador': identificador,
            'duracao': duracao,
            'status': status,
            'detalhes': detalhes,
            'hora_fim': datetime.now().strftime("%H:%M:%S")
        }
        self.historico_etapas.append(registro)
        return duracao

    def finalizar(self, status_geral="Concluído com Sucesso"):
        """Gera o sumário estatístico final e fecha o relatório"""
        fim_dt = datetime.now()
        tempo_total = time.time() - self.tempo_inicio
        
        total_itens = len(self.historico_etapas)
        sucessos = sum(1 for e in self.historico_etapas if "sucesso" in str(e['status']).lower() or "ok" in str(e['status']).lower())
        falhas = total_itens - sucessos
        
        resumo = [
            "",
            "=" * 85,
            "RESUMO E ESTATÍSTICAS DA EXECUÇÃO",
            "=" * 85,
            f"Status Geral                 : {status_geral}",
            f"Total de Itens Processados   : {total_itens}",
            f"Itens com Sucesso            : {sucessos}",
            f"Itens com Falha / Atenção    : {falhas}",
            ""
        ]
        
        if self.historico_etapas:
            resumo.append("DETALHAMENTO DE TEMPO POR ETAPA / ITEM:")
            for item in self.historico_etapas:
                dur_str = formatar_duracao(item['duracao'])
                detalhe_str = f" - {item['detalhes']}" if item['detalhes'] else ""
                resumo.append(f"  • {item['identificador']}: {dur_str} [{item['status']}]{detalhe_str}")
            resumo.append("")
            
        resumo.extend([
            f"Término da Execução          : {fim_dt.strftime('%d/%m/%Y às %H:%M:%S')}",
            f"Tempo Total Decorrido        : {formatar_duracao(tempo_total)}",
            "=" * 85,
            ""
        ])
        
        texto_resumo = "\n".join(resumo)
        with open(self.caminho_arquivo, "a", encoding="utf-8") as f:
            f.write(texto_resumo)
            
        msg_final = f"\n📄 [LOG GRAVADO] Arquivo salvo em: logs/{self.nome_arquivo} (Tempo Total: {formatar_duracao(tempo_total)})"
        if self.callback_console:
            try:
                self.callback_console(msg_final)
            except Exception:
                pass
            
        return self.caminho_arquivo
