import pandas as pd
import os

tabela_ids = pd.read_excel("C:\Users\t_murillo.sanches\Desktop\Projeto Python\python-excel\Planilha_Teste_1.xlsx")
# o grande gpt falou q devia fazer assim para os filtros
filtro_path = "./Filtro.xlsx"

filtro = "CODES"
tabela_ids_coluna = "UID,C,7"

#codes = tabela_ids['UID,C,7'].tolist()

# remover duplicatas
#codes = set(codes)
#codes = list(codes)

# para cada codigo
for i in range(len(tabela_ids[tabela_ids_coluna])):
    if tabela_ids[tabela_ids_coluna][i] == filtro_path[filtro][i]:
        print("Os valores da linha", i+1, "são iguais nas duas planilhas.")
    else:
        print("Os valores da linha", i+1, "são diferentes nas duas planilhas.")
    # pra cada arquivo na pasta consulta
    # fazer uma lista de arquivos na pasta consulta, acho que da pra fazer com
    # a biblioteca 'os', pede pro gpt

    # for arquivo in consulta: 
        # ver se o 'code' ta no arquivo usando o mesmo codigo que ta ali em cima da linha 3 a 9 
        # pega todos as ocorrencias do id 
        # e faz a magica
