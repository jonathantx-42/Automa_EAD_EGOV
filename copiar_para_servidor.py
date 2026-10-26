"""
Automação EGOV - Copiar Arquivos Finais para a Pasta do Servidor
Caminho de Destino Padrão: D:\\SEGEA\\RECENTES\\EGOV\\CEAD\\Cursos Autoinstrucionais\\# 2026\\

O que este script faz:
1. Varre os arquivos finais das pastas:
   - 'Arquivos_CARREGAR/' (arquivos .csv prontos para importação no Moodle)
   - 'Arquivos_Finais_SEDOC/' (arquivos .xlsx com formatação e notas oficiais)
2. Mantém os arquivos intactos no computador local e copia para as pastas de rede/servidor.
3. Identifica a Sigla do Curso e a Turma (ex: PROAMIS_AFG 06.2026 -> #PROAMIS\\AFG\\Turma 06.2026).
4. Se o arquivo já existir na pasta de destino, solicita confirmação para substituir.
"""

import os
import sys
import glob
import re
import shutil

# Configuração segura de codificação para o Windows
if hasattr(sys.stdout, 'encoding') and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

CAMINHO_SERVIDOR_PADRAO = r"D:\SEGEA\RECENTES\EGOV\CEAD\Cursos Autoinstrucionais\# 2026"

def extrair_info_arquivo(nome_arquivo):
    """
    Extrai a sigla, número da turma e ano a partir do nome do arquivo.
    Exemplos:
    - 'PROAMIS_AFG 06.2026 - CARREGAR.csv' -> ('PROAMIS_AFG', 6, '2026')
    - 'AGLP 01.2026 Notas_SEDOC.xlsx'      -> ('AGLP', 1, '2026')
    - 'ASSEDIO_MSD 01.2026 - CARREGAR.csv'  -> ('ASSEDIO_MSD', 1, '2026')
    """
    # Remove extensão temporariamente para análise
    nome_sem_ext, _ = os.path.splitext(nome_arquivo)
    
    padrao = re.compile(r'^([A-Za-z0-9_]+)\s+(?:Turma\s+)?(\d{1,2})(?:\.(\d{4}))?', re.IGNORECASE)
    match = padrao.match(nome_sem_ext.strip())
    if not match:
        return None, None, None

    sigla = match.group(1).strip()
    turma_num = int(match.group(2))
    ano = match.group(3) if match.group(3) else "2026"
    return sigla, turma_num, ano

def encontrar_pasta_curso(caminho_base, sigla):
    """
    Localiza ou define a pasta de destino do curso:
    - Se PROAMIS_*, procura dentro da subpasta '#PROAMIS'
    - Se outros, procura diretamente na raiz do ano
    """
    if sigla.upper().startswith("PROAMIS_"):
        subsigla = sigla.split("_", 1)[1].strip()
        pasta_proamis = os.path.join(caminho_base, "#PROAMIS")
        os.makedirs(pasta_proamis, exist_ok=True)
        
        # Busca sem diferenciar maiúsculas/minúsculas
        subpastas = [d for d in os.listdir(pasta_proamis) if os.path.isdir(os.path.join(pasta_proamis, d))]
        for sub in subpastas:
            if sub.lower() == subsigla.lower():
                return os.path.join(pasta_proamis, sub)
        
        # Se não encontrou subpasta existente, define a nova com o nome da subsigla
        return os.path.join(pasta_proamis, subsigla)
    else:
        pastas_raiz = [d for d in os.listdir(caminho_base) if os.path.isdir(os.path.join(caminho_base, d))]
        # Busca exata case-insensitive
        for p in pastas_raiz:
            if p.lower() == sigla.lower():
                return os.path.join(caminho_base, p)
        # Busca ignorando underscore inicial (ex: _ASSEDIO vs ASSEDIO)
        for p in pastas_raiz:
            if p.lower().lstrip('_') == sigla.lower().lstrip('_'):
                return os.path.join(caminho_base, p)
        
        # Se não encontrou, cria com o nome da sigla
        return os.path.join(caminho_base, sigla)

def encontrar_ou_criar_pasta_turma(pasta_curso, turma_num, ano):
    """
    Localiza a pasta da turma existente (mesmo com variações de digitação, ex 'Turma 06..2026')
    ou cria a pasta padronizada 'Turma XX.YYYY'.
    """
    os.makedirs(pasta_curso, exist_ok=True)
    pastas_existentes = [d for d in os.listdir(pasta_curso) if os.path.isdir(os.path.join(pasta_curso, d))]

    padrao = re.compile(rf'Turma\s*0?{turma_num}\b.*{ano}', re.IGNORECASE)
    for p in pastas_existentes:
        if padrao.search(p):
            return os.path.join(pasta_curso, p)

    # Se não existe, cria a pasta padrão
    pasta_padrao = os.path.join(pasta_curso, f"Turma {turma_num:02d}.{ano}")
    os.makedirs(pasta_padrao, exist_ok=True)
    return pasta_padrao

def copiar_arquivos_para_servidor(caminho_servidor=CAMINHO_SERVIDOR_PADRAO, callback_confirmar=None, log_fn=print):
    """
    Varre os arquivos locais e copia para o servidor/diretório de rede.
    
    callback_confirmar: função(nome_arquivo, pasta_destino) -> 'sim', 'todos', 'nao', 'pular_todos', 'cancelar'
    """
    pasta_base_local = os.path.dirname(os.path.abspath(__file__))
    pasta_carregar = os.path.join(pasta_base_local, "Arquivos_CARREGAR")
    pasta_sedoc = os.path.join(pasta_base_local, "Arquivos_Finais_SEDOC")

    log_fn("=" * 65)
    log_fn("📂 CÓPIA DE ARQUIVOS FINAIS PARA O SERVIDOR DE ARQUIVOS")
    log_fn("=" * 65)
    log_fn(f"Destino Servidor: {caminho_servidor}")
    log_fn("-" * 65)

    if not os.path.exists(caminho_servidor):
        log_fn(f"❌ ERRO: O caminho do servidor não foi encontrado ou não está acessível:")
        log_fn(f"   {caminho_servidor}")
        log_fn("   Verifique se o disco de rede (D:) está conectado.")
        return {
            "sucesso": False,
            "motivo": "Caminho do servidor inacessível",
            "copiados": 0,
            "ignorados": 0
        }

    # Coleta arquivos de ambas as pastas
    arquivos_fila = []
    
    if os.path.exists(pasta_carregar):
        for f in glob.glob(os.path.join(pasta_carregar, "*.csv")):
            arquivos_fila.append((f, "CARREGAR"))
            
    if os.path.exists(pasta_sedoc):
        for f in glob.glob(os.path.join(pasta_sedoc, "*.*")):
            if not os.path.basename(f).startswith("~$"):
                arquivos_fila.append((f, "SEDOC"))

    if not arquivos_fila:
        log_fn("⚠️ Nenhum arquivo encontrado nas pastas:")
        log_fn(f"   • {pasta_carregar}")
        log_fn(f"   • {pasta_sedoc}")
        log_fn("Execute primeiro as etapas 2 (Tratar CSV) e/ou 4 (Formatar SEDOC).")
        return {"sucesso": True, "copiados": 0, "ignorados": 0}

    log_fn(f"Foram encontrados {len(arquivos_fila)} arquivo(s) prontos para copiar.\n")

    modo_sobrescrita = "perguntar"  # 'perguntar', 'todos', 'pular_todos'
    total_copiados = 0
    total_ignorados = 0
    total_erros = 0

    for idx, (caminho_origem, tipo) in enumerate(arquivos_fila, start=1):
        nome_arquivo = os.path.basename(caminho_origem)
        sigla, turma_num, ano = extrair_info_arquivo(nome_arquivo)

        if not sigla or turma_num is None:
            log_fn(f"[{idx}/{len(arquivos_fila)}] ⚠️ Não foi possível identificar Curso/Turma no arquivo: {nome_arquivo}")
            total_erros += 1
            continue

        try:
            pasta_curso = encontrar_pasta_curso(caminho_servidor, sigla)
            pasta_turma = encontrar_ou_criar_pasta_turma(pasta_curso, turma_num, ano)
            caminho_destino = os.path.join(pasta_turma, nome_arquivo)

            # Verifica se já existe no destino
            if os.path.exists(caminho_destino):
                log_fn(f"[{idx}/{len(arquivos_fila)}] ⚠️ Arquivo já existe no destino: {nome_arquivo}")
                log_fn(f"   📁 Destino: {pasta_turma}")

                if modo_sobrescrita == "pular_todos":
                    log_fn("   ⏩ Pulado (opção Pular Todos ativa).")
                    total_ignorados += 1
                    continue
                elif modo_sobrescrita == "todos":
                    log_fn("   🔄 Substituindo (opção Substituir Todos ativa)...")
                else:
                    # Pergunta se pode sobrescrever
                    if callback_confirmar:
                        resp = callback_confirmar(nome_arquivo, pasta_turma)
                    else:
                        # Modo terminal
                        print(f"\nO arquivo '{nome_arquivo}' já existe em:\n{pasta_turma}")
                        escolha = input("Deseja substituir? [s]im / [t]odos / [n]ão / [p]ular todos / [c]ancelar: ").strip().lower()
                        if escolha in ['s', 'sim']:
                            resp = 'sim'
                        elif escolha in ['t', 'todos']:
                            resp = 'todos'
                        elif escolha in ['p', 'pular todos', 'pt']:
                            resp = 'pular_todos'
                        elif escolha in ['c', 'cancelar']:
                            resp = 'cancelar'
                        else:
                            resp = 'nao'

                    if resp == 'cancelar':
                        log_fn("\n🛑 Operação cancelada pelo usuário.")
                        break
                    elif resp == 'todos':
                        modo_sobrescrita = 'todos'
                        log_fn("   🔄 Substituindo este e todos os próximos...")
                    elif resp == 'pular_todos':
                        modo_sobrescrita = 'pular_todos'
                        log_fn("   ⏩ Pulando este e todos os próximos existentes.")
                        total_ignorados += 1
                        continue
                    elif resp == 'nao':
                        log_fn("   ⏩ Arquivo ignorado por decisão do usuário.")
                        total_ignorados += 1
                        continue
                    else:
                        log_fn("   🔄 Substituindo arquivo...")

            else:
                log_fn(f"[{idx}/{len(arquivos_fila)}] 📋 Copiando: {nome_arquivo}")
                log_fn(f"   -> Destino: {pasta_turma}")

            # Efetua a cópia mantendo os atributos originais
            shutil.copy2(caminho_origem, caminho_destino)
            log_fn("   ✅ Copiado com sucesso!")
            total_copiados += 1

        except Exception as e:
            log_fn(f"   ❌ ERRO ao copiar '{nome_arquivo}': {e}")
            total_erros += 1

    log_fn("\n" + "=" * 65)
    log_fn("🎉 RESUMO DA SINCRONIZAÇÃO COM O SERVIDOR:")
    log_fn(f"   • Total de arquivos copiados / atualizados: {total_copiados}")
    log_fn(f"   • Total de arquivos pulados: {total_ignorados}")
    if total_erros > 0:
        log_fn(f"   • Erros encontrados: {total_erros}")
    log_fn(f"   • Arquivos locais preservados em seu computador.")
    log_fn("=" * 65)

    return {
        "sucesso": True,
        "copiados": total_copiados,
        "ignorados": total_ignorados,
        "erros": total_erros
    }

if __name__ == "__main__":
    copiar_arquivos_para_servidor()
