import pandas as pd
import os

# a planilha ta na mesma pasta da pra mandar essa braba aqui V V
# tava assim antes se quiser voltar "C:\Users\t_murillo.sanches\Desktop\Projeto Python\python-excel\Planilha_Teste_1.xlsx"

'''
ATENÇÃO

VAI TRAVAR TUDO SE TIVER MUITO LINHA KK
TESTA COM PLANILHA PEQUENA

'''


table_ids_path = "./Planilha_Teste_1.xlsx"
tabela_ids = pd.read_excel("./Planilha_Teste_1.xlsx")

# o grande gpt falou q devia fazer assim para os filtros
filtro_path = "./Filtro.xlsx"

tabela_filtro = pd.read_excel(filtro_path)


filtro = "CODES"
tabela_ids_coluna = "UID,C,7"

#codes = tabela_ids['UID,C,7'].tolist()

# remover duplicatas
#codes = set(codes)
#codes = list(codes)

# para cada codigo
lista_de_codes = list(set(tabela_filtro[filtro].tolist()))

# para cada item na lsita de filtros
for filter_code in lista_de_codes:
    # se o codigo estiver na coluna id da tabela
    if filter_code in tabela_ids[tabela_ids_coluna].tolist():

        # o enumerate é so um for que tambem tem o id, entao nesse caso, 'linha' é a linha
        # que o codigo se encontra dentro da coluna (code)
        for linha, code in enumerate(tabela_ids[tabela_ids_coluna].tolist()):
            if code == filter_code:
                print(f'codigo: {filter_code} da planilha filtros aparece na linha {linha} em {table_ids_path}')
        
    else:
        print("Os valores da linha", "são diferentes nas duas planilhas.")
    # pra cada arquivo na pasta consulta
    # fazer uma lista de arquivos na pasta consulta, acho que da pra fazer com
    # a biblioteca 'os', pede pro gpt

    # for arquivo in consulta: 
        # ver se o 'code' ta no arquivo usando o mesmo codigo que ta ali em cima da linha 3 a 9 
        # pega todos as ocorrencias do id 
        # e faz a magica
