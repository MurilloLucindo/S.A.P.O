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
tabela_ids = pd.read_excel(table_ids_path)

# o grande gpt falou q devia fazer assim para os filtros
filtro_path = "./Filtro.xlsx"
tabela_filtro = pd.read_excel(filtro_path)


filtro = "CODES"
tabela_ids_coluna = "UID,C,7"

#codes = tabela_ids['UID,C,7'].tolist()

# pega e faz o seguinte:
# a função zip junta o CODE e nome em tupla (code, nome) <- fica juntinho
# e esse id: nome for id, nome bla bla 
# faz um loop que cria chave e id para um dicionario do python, ent {code: name} <- pesquisa de dicionario se precisar

id_nome_dict = {code: nome for code, nome in zip(tabela_filtro[filtro], tabela_filtro['UBS'])}
lista_de_codes = list(set(tabela_filtro[filtro].tolist()))


# PRECISA FAZER OS NUMERO SAIR DA NOTAÇÂO DIENTIFICA
# MANUALMENTE É: SELECIOAN TODAS AS CELULAS > DIREITO > FORMATAÇÂO > CUSTOM 0

table_codes = list(set(tabela_ids[tabela_ids_coluna].tolist()))

# cria o resto fora pra poder salvar fora do loop
planilha_resto = pd.DataFrame(columns=tabela_ids.columns)

# para cada item na lsita de filtros
for table_code in table_codes:
    
    if table_code in lista_de_codes:
        print(f'{table_code} está nos filtros: {id_nome_dict[table_code]}')
        conteudo_code = tabela_ids.loc[tabela_ids[tabela_ids_coluna] == table_code]
        
        planilha_resultado = pd.DataFrame(columns=tabela_ids.columns)

        planilha_resultado = planilha_resultado._append(conteudo_code, ignore_index=True)

        planilha_resultado.to_excel(f'teste/{table_code}_{id_nome_dict[table_code]}.xlsx', index=False, float_format="%.0f")

        exit()

    else:
        print(f'{table_code} nao ta nos filtro')

        conteudo_code = tabela_ids.loc[tabela_ids[tabela_ids_coluna] == table_code]

        planilha_resto = planilha_resto._append(conteudo_code, ignore_index=True)


# salva do loop
planilha_resto.to_excel(f'teste/Restantes.xlsx', index=False, float_format="%.0f")

  