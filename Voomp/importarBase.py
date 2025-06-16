import pandas as pd  # Importa pandas para leitura do Excel
import os  # Para manipular arquivos e caminhos
import easygui  # Para exibir mensagens
from parametro_voomp import inicio_codigo  # Importa a função inicio_codigo

# Definição de cores para exibição no terminal.
GREEN = "\033[92m"  # Verde para sucesso.
RED = "\033[91m"  # Vermelho para erro.
BLUE = "\033[34m"  # Azul para informações.
YELLOW = "\033[93m"  # Amarelo para avisos ou alertas.


# Solicita o ANOMES ao usuário
ANOMES = easygui.enterbox("Digite o ANOMES (YYYYMM):", "Entrada de Dados")


variaveis_de_controle = {
    'CRIAR_DF_VOOMP' : True,
    'CRIAR_DF_PLATOS': True,
    'CRIAR_DF_MEPOUPE': True
}


# Verifica se o usuário clicou em "Cancelar"
if ANOMES is None:
    easygui.msgbox("Operação cancelada.", "Saindo")
else:
    # Chama a função inicio_codigo para obter as variáveis
    ANO, MES, ANOMES, PASTA = inicio_codigo(ANOMES)
 
    # Monta o caminho completo do arquivo Excel
    arquivo_excel = os.path.join(PASTA, f"{MES}.{ANO}_Receita Pós Graduação Infoprodutores (Contabil).xlsx")
 

 
    # Verifica se o arquivo existe antes de tentar importar
    if os.path.exists(arquivo_excel):
        try:
            # Importa a planilha específica "Resumo"
            df = pd.read_excel(arquivo_excel, sheet_name="Resumo")
 
            # Exibe uma mensagem com a quantidade de linhas e colunas importadas
            easygui.msgbox(
                f"Arquivo importado com sucesso!\n\nLinhas: {df.shape[0]}\nColunas: {df.shape[1]}",
                "Importação Concluída"
            )
 
            # (Opcional) Exibir as primeiras 5 linhas no terminal
            print(f'{GREEN}Excel importado : \n',df.head())
 
        except Exception as e:
            easygui.msgbox(f"{RED}Erro ao importar o arquivo: {str(e)}", "Erro")
    else:
        easygui.msgbox(f"{YELLOW}O arquivo não foi encontrado na pasta especificada.", "Erro")


if variaveis_de_controle["CRIAR_DF_VOOMP"]:
    # EXIBINDO UMA MENSAGEM
    print(F'{YELLOW} Criando o DataFrame de Voomp: \n')

    # CRIANDO ODATAFRAME PERSONALIZADO PARA VOOMP
    df_voomp = pd.DataFrame({
        'nome': ["Voomp"] * len(df),
        'CNT_DEB': ["1120112"] * len(df),
        'CNT_CRE': ["3110310"] * len(df),
        'CD_IES': [""] * len(df),
        'CD_PROTHEUS': ["R119"] * len(df),
        '4_ENTI_DEB': ["R119"] * len(df),
        '4_ENTI_CRED': ["R119"] * len(df),
        'matricula': df['matricula'],
        'CD_ESPC': [""] * len(df),
        'CD_POLO': [""] * len(df),
        'HISTORICO_CARGA': ["RECEITA INFOPRODUTOR"] * len(df),
        'COMP_ANO': [ANO] * len(df),
        'COMP_MES': [MES] * len(df),
        'C_CUSTO_DEB': ["14989288700"] * len(df),
        'C_CUSTO_CRE': ["14989288700"] * len(df),
        'VALOR': df['VLR_VOOMP'].round(2)
    })

    # EXIBINDO AS PRIMEIRAS 5 LINHAS DE VOOMP
    print(f'{GREEN} 5 PRIMEIRAS LINHAS DO DATAFRAME VOOMP: \n',df_voomp.head(5))


if variaveis_de_controle['CRIAR_DF_PLATOS']:
    # EXIBINDO UMA MENSAGEM
    print(f'{YELLOW} Criando DataFrame de Platos: \n')

    #CRIANDO DATAFRAME PERSONALIZADO PARA PLATOS
    df_platos = pd.DataFrame({
        'nome': ["Platos"] * len(df),
        'CNT_DEB': ["1120112"] * len(df),
        'CNT_CRE': ["3110111"] * len(df),
        'CD_IES': [""] * len(df),
        'CD_PROTHEUS': ["PL02"] * len(df),
        '4_ENTI_DEB': ["PL02"] * len(df),
        '4_ENTI_CRED': ["PL02"] * len(df),
        'matricula': df['matricula'],
        'CD_ESPC': [""] * len(df),
        'CD_POLO': [""] * len(df),
        'HISTORICO_CARGA': ["RECEITA INFOPRODUTOR"] * len(df),
        'COMP_ANO': [ANO] * len(df),
        'COMP_MES': [MES] * len(df),
        'C_CUSTO_DEB': ["14885900258"] * len(df),
        'C_CUSTO_CRE': ["14885900258"] * len(df),
        'VALOR': df['VLR_PLATOS'].round(2)
    })

    #EXIBINDO 5 PIMEIRAS LINHAS DE PLATOS
    print(f'{GREEN} 5 PRIMEIRAS LINHAS DO DATAFRAME DE PLATOS: \n', df_platos.head(5))


if variaveis_de_controle['CRIAR_DF_MEPOUPE']:
    #EXIBINDO UMA MENSAGEM
    print(f'{YELLOW} Criado DataFrame de Me Poupe: \n')

    #CRIANDO DATAFRAME PERSONALIZADO PARA ME POUPE
    df_mepoupe = pd.DataFrame({
        'nome': ["Me Poupe"] * len(df),
        'CNT_DEB': ["1121411"] * len(df),
        'CNT_CRE': ["2110101"] * len(df),
        'CD_IES': [""] * len(df),
        'CD_PROTHEUS': ["R119"] * len(df),
        '4_ENTI_DEB': ["R119"] * len(df),
        '4_ENTI_CRED': ["R119"] * len(df),
        'matricula': df['matricula'],
        'CD_ESPC': [""] * len(df),
        'CD_POLO': [""] * len(df),
        'HISTORICO_CARGA': ["RECEITA INFOPRODUTOR"] * len(df),
        'COMP_ANO': [ANO] * len(df),
        'COMP_MES': [MES] * len(df),
        'C_CUSTO_DEB': ["14989288700"] * len(df),
        'C_CUSTO_CRE': ["14989288700"] * len(df),
        'VALOR': df['VLR_INFOPRODUTOR'].round(2)
    })

    #EXIBINDO 5 PRIMEIRAS LINHAS DE ME POUPE
    print(f'{GREEN} 5 PRIMEIRAS LINHAS DO DATAFRAME DE ME POUPE \n', df_mepoupe.head(5))

#CRIANDO FUNÇÃO QUE CASO O VALOR SEJA NEGATIVO, ELE INVERTE DEBITO E CREDITO E DEIXA O VALOR POSITIVO
def valor_negat(row):
    if row['VALOR'] < 0:
        row['CNT_DEB'], row['CNT_CRE'] = row['CNT_CRE'], row['CNT_DEB']
        row['VALOR'] = abs(row['VALOR'])
    return row

#APLICANDO A REGRA DOS VALORES NEGATIVOS PARA AS TABELAS
df_voomp = df_voomp.apply(valor_negat, axis=1)
df_platos = df_platos.apply(valor_negat, axis=1)
df_mepoupe = df_mepoupe.apply(valor_negat,axis=1)


#EXIBINDO UMA MENSAGEM
print(f'{YELLOW} CRIANDO DATAFRAME FINAL CONSOLIDADA \n')

#CRIANDO DATAFRAME FINAL CONSOLIDADA
df_final = pd.concat([df_voomp,df_platos,df_mepoupe], axis=0)

#REMOVENDO LINHAS COM VALOR <= 0 DO df_final
df_final = df_final[df_final['VALOR'] > 0]


#EXIBINDO 5 PRIMEIRAS LINHAS DO DATAFRAME FINAL CONSOLIDADO
print(f'{GREEN} 5 PRIMEIRAS LINHAS DO DATAFRAME FINAL CONSOLIDADO \n', df_final.head(5))


#EXPORTANDO O ARQUIVO EXCEL
output_file = os.path.join(PASTA,f"{ANOMES}_Carga_Receita_Pos_Graduacao_Infoprodutores_round_teste_negativos_CC.xlsx")
df_final.to_excel(output_file, sheet_name='BASE TM', index=False)

#EXIBINDO MENSAGEM QUE FOI EXPORTADO COM SUCESSO
easygui.msgbox(f"ARQUIVO EXPORTADO COM SUCESSO PARA {output_file}", "Exportação Concluída")
