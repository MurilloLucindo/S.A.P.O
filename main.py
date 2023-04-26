import pandas as pd

tabela_ids = pd.read_excel('./testes.xlsx')

codes = tabela_ids['ids'].tolist()

# remover duplicatas
codes = set(codes)
codes = list(codes)


# para cada codigo
for code in codes:
    pass
    # pra cada arquivo na pasta consulta
    # fazer uma lista de arquivos na pasta consulta, acho que da pra fazer com
    # a biblioteca 'os', pede pro gpt

    # for arquivo in consulta: 
        # ver se o 'code' ta no arquivo usando o mesmo codigo que ta ali em cima da linha 3 a 9 
        # pega todos as ocorrencias do id 
        # e faz a magica
