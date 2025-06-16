import os #Manipular arquivos e pastas

def inicio_codigo(ANOMES):
    ANO = ANOMES[:4]         # Pega os 4 primeiros dígitos (ano)
    MES = ANOMES[4:6]  # Garante que o mês tenha 2 caracteres (ex: "2" → "02")
    # Construindo o caminho da pasta
    PASTA = os.path.join(r"\\172.22.0.33\Controladoria\Servidor IUNI\Controladoria\Conciliação Geral - Todas as Contas - IES Grupo IUNI",
                        str(ANO),
                        r"GRUPO KROTON - VOOMP PLATOS",
                        str(ANOMES))
    return ANO,MES,ANOMES,PASTA

def round_to_two_places(x):
     return round(x, 2)