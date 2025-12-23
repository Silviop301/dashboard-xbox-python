import pandas as pd
import xlsxwriter
import glob

def gerar_dashboard():
    print("--- Gerando Dashboard (Modo Claro/Light) ---")

    # 1. Localizar arquivo
    arquivos = [f for f in glob.glob('*.xlsx') 
                if not f.startswith('~$') and 'Dashboard_Xbox_Finalizado' not in f]
    
    if not arquivos:
        print("Nenhum arquivo de dados encontrado.")
        return

    arquivo_origem = arquivos[0]
    df_dados = None

    # 2. Carregar dados
    try:
        xls = pd.ExcelFile(arquivo_origem)
        for aba in xls.sheet_names:
            df_temp = pd.read_excel(xls, sheet_name=aba)
            if df_temp.columns is not None:
                df_temp.columns = df_temp.columns.astype(str).str.strip()
            
            if 'Total Value' in df_temp.columns and 'Plan' in df_temp.columns:
                df_dados = df_temp
                break
    except Exception as e:
        print(f"Erro ao ler arquivo: {e}")
        return

    if df_dados is None:
        print("Dados não encontrados.")
        return

    # 3. Cálculos
    faturamento_total = df_dados['Total Value'].sum()
    faturamento_por_plano = df_dados.groupby('Plan')['Total Value'].sum().reset_index()

    if 'Subscription Type' in df_dados.columns and 'Auto Renewal' in df_dados.columns:
        planos_anuais = df_dados[df_dados['Subscription Type'] == 'Annual']
        dados_renovacao = planos_anuais.groupby('Auto Renewal')['Total Value'].sum().reset_index()
    else:
        dados_renovacao = pd.DataFrame({'Auto Renewal': ['No', 'Yes'], 'Total Value': [0, 0]})

    # 4. Geração do Excel
    nome_saida = 'Dashboard_Xbox_Finalizado.xlsx'
    workbook = xlsxwriter.Workbook(nome_saida)
    worksheet = workbook.add_worksheet('Dashboard')

    # --- DEFINIÇÃO DE CORES (MODO CLARO) ---
    verde_xbox = '#107C10' # Verde oficial Xbox mais escuro para ler no branco
    fundo_claro = '#FFFFFF' # Branco total
    texto_preto = '#000000' 
    cinza_cabecalho = '#E0E0E0'
    
    # Formatos
    fmt_titulo = workbook.add_format({'bold': True, 'font_color': verde_xbox, 'font_size': 20, 'bg_color': fundo_claro})
    fmt_rotulo = workbook.add_format({'font_color': '#555555', 'bg_color': fundo_claro, 'italic': True})
    fmt_valor = workbook.add_format({'num_format': 'R$ #,##0', 'bg_color': fundo_claro, 'font_color': texto_preto, 'bold': True, 'font_size': 14})
    fmt_fundo = workbook.add_format({'bg_color': fundo_claro})
    fmt_header_tab = workbook.add_format({'bold': True, 'bg_color': cinza_cabecalho, 'border': 1, 'font_color': texto_preto})
    fmt_celula_tab = workbook.add_format({'bg_color': fundo_claro, 'border': 1, 'font_color': texto_preto})
    fmt_moeda_tab = workbook.add_format({'num_format': 'R$ #,##0', 'bg_color': fundo_claro, 'border': 1, 'font_color': texto_preto})

    # Layout
    worksheet.set_column('A:Z', 15, fmt_fundo) # Aplica fundo branco em tudo
    worksheet.set_column('B:B', 25) # Alarga coluna B para caber textos

    worksheet.write('B2', 'XBOX GAME PASS - SALES DASHBOARD', fmt_titulo)
    worksheet.write('B4', 'FATURAMENTO TOTAL', fmt_rotulo)
    worksheet.write('B5', faturamento_total, fmt_valor)

    # Tabelas Auxiliares (Visíveis agora para conferência, mas discretas)
    row = 20
    start_plan = row + 1
    worksheet.write(row, 1, 'Plano', fmt_header_tab)
    worksheet.write(row, 2, 'Receita', fmt_header_tab)
    for i, linha in faturamento_por_plano.iterrows():
        row += 1
        worksheet.write(row, 1, linha['Plan'], fmt_celula_tab)
        worksheet.write(row, 2, linha['Total Value'], fmt_moeda_tab)
    end_plan = row

    row = 20
    start_ren = row + 1
    worksheet.write(row, 4, 'Renovação', fmt_header_tab)
    worksheet.write(row, 5, 'Receita', fmt_header_tab)
    for i, linha in dados_renovacao.iterrows():
        row += 1
        worksheet.write(row, 4, linha['Auto Renewal'], fmt_celula_tab)
        worksheet.write(row, 5, linha['Total Value'], fmt_moeda_tab)
    end_ren = row

    # 5. Gráficos (Ajustados para fundo claro)
    
    # Colunas
    chart_col = workbook.add_chart({'type': 'column'})
    chart_col.add_series({
        'name': 'Receita por Plano',
        'categories': ['Dashboard', start_plan, 1, end_plan, 1],
        'values':     ['Dashboard', start_plan, 2, end_plan, 2],
        'fill':       {'color': verde_xbox},
        'data_labels': {'value': True, 'num_format': 'R$ #,##0', 'font': {'color': 'black'}}
    })
    chart_col.set_title({'name': 'Receita por Plano (R$)', 'name_font': {'color': 'black'}})
    # Remove bordas e deixa fundo branco
    chart_col.set_chartarea({'fill': {'color': fundo_claro}, 'border': {'none': True}})
    chart_col.set_plotarea({'fill': {'color': fundo_claro}, 'border': {'none': True}})
    # Eixos pretos
    chart_col.set_x_axis({'major_gridlines': {'visible': False}, 'num_font': {'color': 'black'}})
    chart_col.set_y_axis({'major_gridlines': {'visible': True, 'line': {'color': '#D9D9D9'}}, 'num_font': {'color': 'black'}})
    chart_col.set_legend({'none': True})
    worksheet.insert_chart('B7', chart_col)

    # Rosca
    chart_doughnut = workbook.add_chart({'type': 'doughnut'})
    chart_doughnut.add_series({
        'name': 'Renovação Automática',
        'categories': ['Dashboard', start_ren, 4, end_ren, 4],
        'values':     ['Dashboard', start_ren, 5, end_ren, 5],
        'points': [{'fill': {'color': '#CCCCCC'}}, {'fill': {'color': verde_xbox}}], # Cinza e Verde
        'data_labels': {'percentage': True, 'num_font': {'color': 'black'}}
    })
    chart_doughnut.set_title({'name': 'Renovação Automática (Anual)', 'name_font': {'color': 'black'}})
    chart_doughnut.set_chartarea({'fill': {'color': fundo_claro}, 'border': {'none': True}})
    chart_doughnut.set_legend({'position': 'bottom', 'font': {'color': 'black'}})
    worksheet.insert_chart('G7', chart_doughnut)

    workbook.close()
    print(f"Sucesso! Arquivo '{nome_saida}' atualizado com fundo branco.")

if __name__ == "__main__":
    gerar_dashboard()