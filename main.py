import pandas as pd
import os

# a planilha ta na mesma pasta da pra mandar essa braba aqui V V
# tava assim antes se quiser voltar "C:\Users\t_murillo.sanches\Desktop\Projeto Python\python-excel\Planilha_Teste_1.xlsx"

'''
ATENÇÃO

VAI TRAVAR TUDO SE TIVER MUITO LINHA KK
TESTA COM PLANILHA PEQUENA

'''


table_ids_path = "./Planilha_Teste_2.xlsx"
tabela_ids = pd.read_excel(table_ids_path)

# o grande gpt falou q devia fazer assim para os filtros
filtro_path = "./Filtro.xlsx"
tabela_filtro = pd.read_excel(filtro_path)


filtro = "CODES"
tabela_ids_coluna = "BPI_UID,C,7"

#codes = tabela_ids['UID,C,7'].tolist()

# pega e faz o seguinte:
# a função zip junta o CODE e nome em tupla (code, nome) <- fica juntinho

# e esse id: nome for id, nome bla bla 
# faz um loop que cria chave e id para um dicionario do python, ent {code: name} <- pesquisa de dicionario se precisar

id_nome_dict = {code: nome for code, nome in zip(tabela_filtro[filtro], tabela_filtro['UBS'])}
lista_de_codes = list(set(tabela_filtro[filtro].tolist()))


# PRECISA FAZER OS NUMERO SAIR DA NOTAÇÂO DIENTIFICA
# MANUALMENTE É: SELECIOAN TODAS AS CELULAS > DIREITO > FORMATAÇÂO > CUSTOM 0


# para cada item na lsita de filtros
for filter_code in lista_de_codes:
    
    print(f'Fazendo para {filter_code}: {id_nome_dict[filter_code]}')
    if filter_code in tabela_ids[tabela_ids_coluna].tolist():
        conteudo_code = tabela_ids.loc[tabela_ids[tabela_ids_coluna] == filter_code]
        
        planilha_resultado = pd.DataFrame(columns=tabela_ids.columns)

        planilha_resultado = planilha_resultado._append(conteudo_code, ignore_index=True)

        planilha_resultado.to_excel(f'teste/{filter_code}_{id_nome_dict[filter_code]}.xlsx', index=False, float_format="%.0f")
        
    else:
        print(f"{filter_code} nao encontrado na tabela")
  