import pandas as pd
import os
import threading

class XLSXCreator:

    def __init__(self, tabela_ids_path = None, tabela_ids_col = None, tabela_filtro_path = None, tabela_filtro_col = None, output_folder = None, window = None):
        self.tabela_ids_path = tabela_ids_path #"./Planilha_planilhas_2.xlsx"
        self.tabela_ids_col = tabela_ids_col
        self.tabela_ids = None

        self.tabela_filtro_path = tabela_filtro_path #"./Filtro.xlsx"
        self.tabela_filtro_col = tabela_filtro_col
        self.tabela_filtro = None

        self.id_nome_dict = None
        self.lista_de_codes = None
        self.table_codes = None

        self.planilha_resto = None

        self.output_folder = output_folder if output_folder else './Planilhas'
        
        self.threads = []

        self.loaded = False

        self.semaphore = threading.Semaphore(1)
        #
        self.window = window if window else None

    def load(self):
        self.tabela_ids = pd.read_excel(self.tabela_ids_path)
        self.tabela_filtro = pd.read_excel(self.tabela_filtro_path)

        self.planilha_resto = pd.DataFrame(columns=self.tabela_ids.columns)

        self.id_nome_dict = {code: nome for code, nome in zip(self.tabela_filtro[self.tabela_filtro_col], self.tabela_filtro['UBS'])}
        self.lista_de_codes = list(set(self.tabela_filtro[self.tabela_filtro_col].tolist()))

        self.table_codes = list(set(self.tabela_ids[self.tabela_ids_col].tolist()))

        #

        self.loaded = True

    def __filter_and_create(self, table_code):
        if table_code in self.lista_de_codes:
            if self.window: self.window.log_print(f'Criando planilha para {table_code}: {self.id_nome_dict[table_code]}...') 

            conteudo_code = self.tabela_ids.loc[self.tabela_ids[self.tabela_ids_col] == table_code]
            
            planilha_resultado = pd.DataFrame(columns=self.tabela_ids.columns)

            planilha_resultado = planilha_resultado._append(conteudo_code, ignore_index=True)

            planilha_resultado.to_excel(f'{self.output_folder}/{table_code}_{self.id_nome_dict[table_code]}.xlsx', index=False, float_format="%.0f")

        else:
            if self.window: self.window.log_print(f'{table_code}, {self.id_nome_dict[table_code]} não está nos filtros. Adicionando aos restantes... ') 
            self.semaphore.acquire()
            try:
                # do some work that modifies the shared list
                conteudo_code = self.tabela_ids.loc[self.tabela_ids[self.tabela_ids_col] == table_code]
                self.planilha_resto = self.planilha_resto._append(conteudo_code, ignore_index=True)

            finally:
                self.semaphore.release()
            
    def set_tabela_ids_path(self, path: str):
        self.tabela_ids_path = path

    def set_tabela_filtro_path(self, path: str):
        self.tabela_filtro_path = path

    def set_tabela_ids_col(self, col: str):
        self.tabela_ids_col = col

    def set_tabela_filtro_col(self, col: str):
        self.tabela_filtro_col = col
            
    def start(self):
        if not self.loaded:
            self.load()

        if not os.path.exists(self.output_folder):
            os.mkdir(self.output_folder)

        for table_code in self.table_codes:
            t = threading.Thread(target=self.__filter_and_create, args=(table_code,))
            self.threads.append(t)

        for t in self.threads:
            t.start()
        
        for t in self.threads:
            t.join()

        #salva planilha resto    
        self.planilha_resto.to_excel(f'{self.output_folder}/Restantes.xlsx', index=False, float_format="%.0f")
        