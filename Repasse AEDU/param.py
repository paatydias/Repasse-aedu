import easygui
import os
import re

while True:
    # Exibir o prompt para entrada do ANOMES
    ANOMES = easygui.enterbox("Digite o ANOMES (YYYYMM):", "Entrada de Dados")

    # Se o usuário clicar em "Cancelar", encerramos o programa
    if ANOMES is None:
        easygui.msgbox("Operação cancelada.", "Saindo")
        break

    # Verifica se a entrada é válida
    if re.fullmatch(r"\d{6}", ANOMES):
        ANO = ANOMES[:4]         # Pega os 4 primeiros dígitos (ano)
        MES = ANOMES[4:6].zfill(2)  # Garante que o mês tenha 2 caracteres (ex: "2" → "02")

        # Construindo o caminho da pasta
        PASTA = os.path.join(r"\\Admspsvpltfs01\Departamentos\Controladoria\Servidor IUNI\Carlos Menolli",
                             f"{ANO}",
                             f"{MES}.{ANO}")

        # Exibir o resultado em um pop-up
        easygui.msgbox(f"ANO: {ANO}\nMÊS: {MES}\nANOMES: {ANOMES}\nPASTA: {PASTA}", "Resultado")
        break  # Sai do loop após entrada válida

    else:
        easygui.msgbox("Erro: Digite um ANOMES válido no formato YYYYMM!", "Erro", "OK")


"""""
# Verifica se a pasta existe
if os.path.exists(PASTA):  
    arquivos = os.listdir(PASTA)
    lista_arquivos = "\n".join(arquivos) if arquivos else "Nenhum arquivo encontrado."

    # Exibe os arquivos no prompt
    easygui.msgbox(f"Pasta encontrada!\n\nPASTA: {PASTA}\n\nArquivos:\n{lista_arquivos}", "Conteúdo da Pasta")

else:
    easygui.msgbox("Erro: A pasta não existe.", "Erro")
"""""