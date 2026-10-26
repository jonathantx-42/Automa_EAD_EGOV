"""
Automação Moodle - EGOV (Escola de Governo do DF)
Avaliação e Formatação Final das Planilhas de Notas do Moodle

O que este script faz:
1. Localiza os arquivos de notas (.xlsx) exportados do Moodle.
2. Solicita o período de realização do curso para o cabeçalho.
3. Identifica a sigla e a turma com base no nome do arquivo e busca o nome oficial no dicionário de cursos.
4. Aplica a regra de avaliação:
   - RPIM: nota mínima 60% (sessenta por cento).
   - Demais cursos: nota mínima 70% (setenta por cento).
   - Define status: 'Certificado', 'Não concluinte' ou 'Desistente'.
5. Gera a planilha oficial formatada em Excel (com cabeçalhos, bordas, tabela estilizada, fórmulas e base legal).
6. Salva as planilhas finalizadas na pasta 'Notas_Avaliadas_Finais/' com o sufixo '_SEDOC.xlsx'.
"""

import os
import sys
import glob
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

# Configuração segura para evitar erros de acentuação no terminal do Windows
if hasattr(sys.stdout, 'encoding') and sys.stdout.encoding and sys.stdout.encoding.lower() != 'utf-8':
    try:
        if hasattr(sys.stdout, 'reconfigure'):
            sys.stdout.reconfigure(encoding='utf-8')
        if hasattr(sys.stderr, 'reconfigure'):
            sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# --- DICIONÁRIO DE CURSOS ---
cursos_dict = {
    "AOLP": "Acordo Ortográfico da Língua Portuguesa",
    "AGLP": "Atualização em Gramática da Língua Portuguesa",
    "AP": "Atendimento ao Público",
    "BPLPR": "Boas Práticas em Língua Portuguesa e Redação",
    "LIBRAS": "Comunicando em LIBRAS",
    "DCG": "Desenvolvimento de competências gerenciais",
    "EAD": "Educação a Distância - conceitos, histórico e panorama",
    "ECS": "Entendendo o Controle Social",
    "ETICA": "Ética - uma questão de escolha",
    "FF": "Formação de Formadores",
    "FLUENCIA": "fluêncIA - Noções básicas de I.A.",
    "GQ": "Gestão da Qualidade",
    "GPROC": "Gestão de Processos",
    "GCON": "Gestão do Conhecimento",
    "GSCRO": "Governança de Serviços - Construindo resultados a partir da Ouvidoria",
    "LEI14133": "Implementação da Lei de Licitações e Contratos – Lei nº 14.133/2021",
    "LAI": "Lei de Acesso à Informação - LAI na prática",
    "LGPD": "Lei Geral de Proteção de Dados (LGPD)",
    "LGPD_LI": "Lei Geral de Proteção de Dados - Legislação e Implementação",
    "MROSC": "Marco Regulatório das Organizações da Sociedade Civil (MROSC) – Módulo: Manual MROSC-DF",
    "MASP": "Metodologia de Análise e Solução de Problemas",
    "NR32": "Norma Regulamentadora 32 (NR 32)",
    "OUVIDORIA": "O 'X' da Questão: Decifrando a Ouvidoria",
    "PMCO": "Produtividade e Melhoria do Clima Organizacional",
    "PAI": "Programa de Ambientação e Integração",
    "RPIM": "Ressignificar: Proteção integral às mulheres",
    "SEI4USAR": "SEI USAR 4.0",
    "ASSEDIO_MSD": "Assédio Moral, Sexual e Discriminação no Trabalho",
    "FLUENCIA_EDUC": "fluêncIA para Educadores",
    "FLUENCIA_JUD": "fluêncIA para Profissionais do Judiciário",
    "GFPCF": "Gestão de Finanças Pessoais e Cidadania Financeira",
    "LC840": "Regime Jurídico dos Servidores Públicos Civis do Distrito Federal – Lei Complementar nº 840/2011",
    "PROAMIS_IPN": "PROAMIS - Importância do pré-natal",
    "PROAMIS_PBLH": "PROAMIS - Papel do BLH e doação do leite materno",
    "PROAMIS_PPN": "PROAMIS - Preparação para o parto e para o nascimento",
    "PROAMIS_PI": "PROAMIS - Proteção à infância",
    "PROAMIS_Puer": "PROAMIS - Puerpério",
    "PROAMIS_SMM": "PROAMIS - Saúde mental materna",
    "PROAMIS_TN": "PROAMIS - Triagem neonatal",
    "PROAMIS_AGL": "PROAMIS - Alimentação da gestante e da lactante",
    "PROAMIS_AFG": "PROAMIS - Alterações fisiológicas na gravidez",
    "PROAMIS_Amam": "PROAMIS - Amamentação",
    "PROAMIS_CRN": "PROAMIS - Cuidados do recém-nascido",
    "PROAMIS_HOB": "PROAMIS - Hora de ouro do bebê",
    "PROAMIS_IACDC": "PROAMIS - Importância do Acompanhamento, do Crescimento e Desenvolvimento da Criança"
}

def definir_status(nota, nota_minima):
    """Classifica a nota do aluno de acordo com a nota mínima"""
    nota_texto = str(nota).strip()
    if nota_texto in ['-', '', 'nan', 'None']:
        return 'Desistente'
    try:
        nota_num = float(nota_texto.replace(',', '.'))
        if nota_num >= nota_minima:
            return 'Certificado'
        else:
            return 'Não concluinte'
    except ValueError:
        return 'Não concluinte'

def formatar_notas_gui(data_curso_global, log_fn=print):
    """Execução do processo de formatação tanto para GUI quanto para terminal"""
    pasta_atual = os.path.dirname(os.path.abspath(__file__))
    pasta_saida = os.path.join(pasta_atual, "Arquivos_Finais_SEDOC")
    os.makedirs(pasta_saida, exist_ok=True)

    # Busca arquivos .xlsx (procura primeiro em 'Arquivos_Originais_Moodle', depois na pasta atual)
    pasta_origem = os.path.join(pasta_atual, "Arquivos_Originais_Moodle")
    if os.path.exists(pasta_origem):
        arquivos_excel = glob.glob(os.path.join(pasta_origem, "*.xlsx"))
    else:
        arquivos_excel = []

    if not arquivos_excel:
        arquivos_excel = glob.glob(os.path.join(pasta_atual, "*.xlsx"))

    ignorar_termos = ["finalizado", "sedoc", "lista", "cursos"]
    arquivos_excel = [
        arq for arq in arquivos_excel 
        if not any(termo in os.path.basename(arq).lower() for termo in ignorar_termos)
    ]

    if len(arquivos_excel) == 0:
        log_fn("\n❌ Nenhum arquivo de Notas (.xlsx) de alunos foi encontrado (verifique a pasta 'Arquivos_Originais_Moodle').")
        log_fn("Certifique-se de que os arquivos de notas baixados do Moodle estão na pasta correta.")
        return

    log_fn("=" * 60)
    log_fn("📊 AVALIAÇÃO E FORMATAÇÃO DE NOTAS - EGOV")
    log_fn("=" * 60)
    log_fn(f"Foram encontrados {len(arquivos_excel)} arquivos de notas para processamento.")
    log_fn(f"Período do curso informado: {data_curso_global}\n")
    log_fn("-" * 60)

    total_processados = 0

    for arquivo in arquivos_excel:
        nome_base = os.path.basename(arquivo)
        log_fn(f"Processando turma: {nome_base}...")

        try:
            # Descobre a sigla e a turma com base no nome do arquivo
            nome_limpo = nome_base.replace(".xlsx", "").replace(" Notas", "").strip()
            partes = nome_limpo.split(" ")
            sigla = partes[0].upper()
            turma = partes[1] if len(partes) > 1 else ""

            # Regra de nota mínima
            if sigla == "RPIM":
                nota_min = 60
                texto_porcentagem = "60% (sessenta porcento)"
            else:
                nota_min = 70
                texto_porcentagem = "70% (setenta porcento)"

            df = pd.read_excel(arquivo)

            # Junta Nome e Sobrenome
            if 'Nome' in df.columns and 'Sobrenome' in df.columns:
                df['Nome Completo'] = df['Nome'].astype(str).str.strip() + " " + df['Sobrenome'].astype(str).str.strip()
            else:
                col_nome, col_sobrenome = df.columns[0], df.columns[1]
                df['Nome Completo'] = df[col_nome].astype(str).str.strip() + " " + df[col_sobrenome].astype(str).str.strip()

            df = df[~df['Nome Completo'].str.contains('nan nan|None', case=False, na=False)]
            df = df[df['Nome Completo'].str.strip() != ""]
            df = df.sort_values(by='Nome Completo')

            # Localiza Notas
            nome_coluna_nota = "Total do curso (Real)"
            if nome_coluna_nota not in df.columns:
                colunas_total = [c for c in df.columns if "total" in str(c).lower()]
                if colunas_total:
                    nome_coluna_nota = colunas_total[0]
                else:
                    log_fn(f"  -> ❌ ERRO: Coluna de notas 'Total do curso (Real)' não encontrada. Pulando arquivo.")
                    continue

            df_final = pd.DataFrame()
            df_final['Nome Completo'] = df['Nome Completo']
            df_final['Total do curso'] = df[nome_coluna_nota]
            df_final['Situação'] = df_final['Total do curso'].apply(lambda n: definir_status(n, nota_min))
            df_final.insert(0, 'N.', range(1, len(df_final) + 1))

            # Cria a planilha formatada no openpyxl
            nome_completo = cursos_dict.get(sigla, sigla)
            texto_final = f"{nome_completo} - Turma {turma}" if turma else nome_completo

            wb = Workbook()
            ws = wb.active
            ws.title = "Certificados"

            font_padrao = Font(name="Verdana", size=12)
            font_negrito = Font(name="Verdana", size=12, bold=True)

            ws["A1"] = "Curso:"
            ws["B1"] = texto_final
            ws["B1"].font = font_negrito

            ws["A2"] = "Realização:"
            ws["B2"] = data_curso_global
            ws["B2"].font = font_negrito

            # Base Legal Dinâmica (Linha 4)
            ws.merge_cells("A4:D4")
            ws["A4"] = f"Conforme o Art. 14, Seção II, Capitulo III, da Ordem de Serviço nº 1, de 12/01/2024, terá direito a certificado o cursista/participante de cursos/atividades que tiver aproveitamento mínimo de {texto_porcentagem} no desenvolvimento das atividades avaliativas disponíveis no AVA."
            ws["A4"].alignment = Alignment(horizontal="left", vertical="bottom", wrap_text=True)
            ws.row_dimensions[4].height = 65

            # Títulos
            titulos = ["N.", "Nome", "Total do curso", "Situação"]
            for col_idx, titulo in enumerate(titulos, 1):
                cell = ws.cell(row=6, column=col_idx, value=titulo)
                cell.font = font_negrito

            # Dados
            for row_idx, row_data in enumerate(df_final.values, 7):
                for col_idx, value in enumerate(row_data, 1):
                    ws.cell(row=row_idx, column=col_idx, value=value)

            last_row = len(df_final) + 6

            tab = Table(displayName="Tabela_Certificados", ref=f"A6:D{last_row}")
            style = TableStyleInfo(name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
            tab.tableStyleInfo = style
            ws.add_table(tab)

            # Largura das colunas (Coluna A com largura 15 para caber 'Realização:')
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 50
            ws.column_dimensions['C'].width = 20
            ws.column_dimensions['D'].width = 20

            # Linha FIM
            linha_fim = last_row + 2
            ws.merge_cells(f"A{linha_fim}:D{linha_fim}")
            celula_fim = ws[f"A{linha_fim}"]
            celula_fim.value = "FIM"
            celula_fim.alignment = Alignment(horizontal="right", vertical="bottom")
            borda_topo = Border(top=Side(style='medium'))
            for col in ['A', 'B', 'C', 'D']:
                ws[f"{col}{linha_fim}"].border = borda_topo

            # Total de Certificados
            ws["G7"] = "TOTAL"
            ws["G8"] = '=COUNTIF(D:D, "Certificado")'
            ws["G7"].font = font_negrito

            for row in ws.iter_rows():
                for cell in row:
                    if not cell.font or not cell.font.bold:
                        cell.font = font_padrao

            # Salva o arquivo com sufixo _SEDOC
            novo_nome = nome_base.replace(".xlsx", "_SEDOC.xlsx")
            caminho_saida_excel = os.path.join(pasta_saida, novo_nome)
            wb.save(caminho_saida_excel)

            log_fn(f"  -> ✅ Avaliado e formatado com sucesso! Salvo em: Arquivos_Finais_SEDOC/{novo_nome}")
            total_processados += 1

        except Exception as e:
            log_fn(f"  -> ❌ ERRO ao processar '{nome_base}': {e}")

    log_fn("\n" + "=" * 60)
    log_fn(f"🎉 Processo finalizado! {total_processados} planilhas formatadas na pasta:")
    log_fn(f"📂 {pasta_saida}")
    log_fn("=" * 60)

if __name__ == "__main__":
    # Abre seletor com calendário drop-down para evitar digitação manual
    periodo = None
    try:
        import tkinter as tk
        from tkcalendar import DateEntry

        root = tk.Tk()
        root.title("Período do Curso - SEDOC")
        root.geometry("400x190")
        root.resizable(False, False)

        tk.Label(root, text="Selecione o Período de Realização do Curso:", font=("Verdana", 10, "bold")).pack(pady=(15, 10))

        frame_datas = tk.Frame(root)
        frame_datas.pack(pady=5)

        tk.Label(frame_datas, text="Início:").grid(row=0, column=0, padx=5)
        cal_ini = DateEntry(frame_datas, width=12, background='#1f538d', foreground='white', date_pattern='dd/mm/yyyy', locale='pt_BR')
        cal_ini.grid(row=0, column=1, padx=(0, 15))

        tk.Label(frame_datas, text="Fim:").grid(row=0, column=2, padx=5)
        cal_fim = DateEntry(frame_datas, width=12, background='#1f538d', foreground='white', date_pattern='dd/mm/yyyy', locale='pt_BR')
        cal_fim.grid(row=0, column=3, padx=5)

        resultado = []

        def confirmar():
            resultado.append(f"{cal_ini.get()} a {cal_fim.get()}")
            root.destroy()

        btn_confirmar = tk.Button(root, text="Confirmar e Processar", command=confirmar, bg="#1f538d", fg="white", font=("Verdana", 9, "bold"), padx=10, pady=5)
        btn_confirmar.pack(pady=15)

        root.mainloop()

        if resultado:
            periodo = resultado[0]

    except Exception:
        periodo = input("📅 Digite o período de realização do curso (Ex: 07/09/2026 a 21/09/2026): ").strip()

    if not periodo:
        print("❌ Processo cancelado. O período de realização do curso é obrigatório.")
        sys.exit()

    formatar_notas_gui(periodo)
