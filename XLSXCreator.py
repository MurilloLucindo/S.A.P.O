import pandas as pd
import os
import threading

class XLSXCreator:

    def __init__(self, tabela_ids_path=None, tabela_ids_col=None, tabela_filtro_path=None, tabela_filtro_col=None, output_folder='./Planilhas', ttkobject=None):
        # CONFIGURACAO DA TABELA
        self.tabela_ids_path = tabela_ids_path
        self.tabela_ids_col = tabela_ids_col

        # CONFIGURACAO DOS FILTROS
        self.tabela_filtro_path = tabela_filtro_path
        self.tabela_filtro_col = tabela_filtro_col

        # Pasta do output
        self.output_folder = output_folder

        # Se tiver a janela ele printa na janela
        self.ttkobject = ttkobject

        # Pra verificar se ja carregou os dados
        self.loaded = False

        # Threads
        self.semaphore = threading.Semaphore(1)
        self.threads = []

    def is_ready(self):
        return all([
            self.tabela_ids_path is not None,
            self.tabela_ids_col is not None,
            self.tabela_filtro_path is not None,
            self.tabela_filtro_col is not None
        ])

    def load(self):

        # carrega tabela 
        self.tabela_ids = pd.read_excel(self.tabela_ids_path)

        # carrega filtros
        self.tabela_filtro = pd.read_excel(self.tabela_filtro_path)

        # pega nomes da tabela filtro
        self.id_nome_dict = dict(zip(self.tabela_filtro[self.tabela_filtro_col], self.tabela_filtro['UBS']))

        # pega os codigos da tabela filtros
        self.lista_de_codes = list(set(self.tabela_filtro[self.tabela_filtro_col]))

        # pega os codigos da tabela
        self.table_codes = list(set(self.tabela_ids[self.tabela_ids_col]))

        # cria planilha resto
        self.planilha_resto = pd.DataFrame(columns=self.tabela_ids.columns)

        # carregado
        self.loaded = True
        
        return True

    def __filter_and_create(self, table_code):
        # Verifica se a tabela está na lista de códigos
        if table_code in self.lista_de_codes:
            # Se estiver, cria a planilha correspondente
            if self.ttkobject:
                # Imprime mensagem de log
                message = f'Criando planilha para {table_code}: {self.id_nome_dict[table_code]}...'
                self.ttkobject.log_print(message)
            
            # Seleciona o conteúdo da tabela específica
            conteudo_code = self.tabela_ids.loc[self.tabela_ids[self.tabela_ids_col] == table_code]
            # Cria uma cópia do conteúdo selecionado
            planilha_resultado = conteudo_code.copy()
            
            # Define o nome e o caminho do arquivo da planilha
            filename = f'{table_code}_{self.id_nome_dict[table_code]}.xlsx'
            path = os.path.join(self.output_folder, filename)
            # Salva a planilha no arquivo especificado
            planilha_resultado.to_excel(path, index=False, float_format="%.0f")
            
        # Se a tabela não estiver na lista de códigos
        else:
            if self.ttkobject:
                # Imprime mensagem de log
                message = f'{table_code} não está nos filtros. Adicionando aos restantes... '
                self.ttkobject.log_print(message)
            
            # Seleciona o conteúdo da tabela específica
            conteudo_code = self.tabela_ids.loc[self.tabela_ids[self.tabela_ids_col] == table_code]
            
            # Aquisição do semáforo antes da operação crítica
            self.semaphore.acquire()
            try:
                self.planilha_resto = self.planilha_resto._append(conteudo_code, ignore_index=True)
            finally:
                # Liberação do semáforo após a operação crítica
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
        if self.is_ready():
            # Checa se esta carregado (as planilhas)
            if not self.loaded:
                self.load()

            # Checa se a pasta de output existe
            if not os.path.exists(self.output_folder):
                os.mkdir(self.output_folder)

            for table_code in self.table_codes:
                t = threading.Thread(target=self.__filter_and_create, args=(table_code,))
                self.threads.append(t)

            for t in self.threads:
                t.start()

            for t in self.threads:
                t.join()

            restantes_filename = 'Restantes.xlsx'
            restantes_path = os.path.join(self.output_folder, restantes_filename)
            self.planilha_resto.to_excel(restantes_path, index=False, float_format="%.0f")

            return True