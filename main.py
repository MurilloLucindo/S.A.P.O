import pandas as pd
import os

tabela_ids = pd.read_excel('./Planilha_Teste_1.xlsx')
# o grande gpt falou q divia fazer assim para os filtross
filtro_path = "./Filtro.xlsx"
filtro_files = os.listdir(filtro_path)

codes = tabela_ids['UID,C,7'].tolist()

# remover duplicatas
codes = set(codes)
codes = list(codes)

# para cada codigo
for codes in filtro_files:
    print (codes)
    # pra cada arquivo na pasta consulta
    # fazer uma lista de arquivos na pasta consulta, acho que da pra fazer com
    # a biblioteca 'os', pede pro gpt

    # for arquivo in consulta: 
        # ver se o 'code' ta no arquivo usando o mesmo codigo que ta ali em cima da linha 3 a 9 
        # pega todos as ocorrencias do id 
        # e faz a magica
