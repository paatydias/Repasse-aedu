import param  # Importa as variáveis do primeiro arquivo
import pandas as pd  # Importa pandas para leitura do Excel
import os  # Para manipular arquivos e caminhos
import easygui  # Para exibir mensagens

# Monta o caminho completo do arquivo Excel
arquivo_excel = os.path.join(param.PASTA, f"REP_UNIDERP_COLABORAR-{param.ANO}-{param.MES}.xlsx")

# Verifica se o arquivo existe antes de tentar importar
if os.path.exists(arquivo_excel):
    # Importa a planilha específica "1000_30"
    df = pd.read_excel(arquivo_excel, sheet_name="1000_30")

    # Exibe uma mensagem com a quantidade de linhas e colunas importadas
    easygui.msgbox(f"Arquivo importado com sucesso!\n\nLinhas: {df.shape[0]}\nColunas: {df.shape[1]}", "Importação Concluída")
    
    # (Opcional) Exibir as primeiras 5 linhas no terminal
    print(df.head())

else:
    easygui.msgbox("O arquivo não foi encontrado na pasta especificada.", "Erro")
