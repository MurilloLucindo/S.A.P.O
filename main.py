from XLSXCreator import XLSXCreator

# sistema absolutamente podereoso de organização
sapo = XLSXCreator(tabela_ids_path="./Planilha_Teste_2.xlsx", tabela_ids_col="BPI_UID,C,7", tabela_filtro_path="./Filtro.xlsx", tabela_filtro_col="CODES")

sapo.load()

sapo.start()