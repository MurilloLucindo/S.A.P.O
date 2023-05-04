import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter.filedialog import askopenfilename
from tkinter.scrolledtext import ScrolledText
from PIL import Image, ImageTk
import threading

from XLSXCreator import XLSXCreator

class SapoWindow(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.pack(fill=BOTH, expand=YES)

        # ttk variables
        self.filter_filename = ttk.StringVar()
        self.column_filter_entry = None
        self.table_filename = ttk.StringVar()
        self.column_table_entry = None

        # images
        self.forg = ImageTk.PhotoImage(Image.open('assets/forg.png'))

        # styles
        self.style = ttk.Style()
        self.style.configure('success.TButton', font=('Helvetica', 35))

        # upper container
        self.upper_container = ttk.Frame(master=self, padding=2)
        self.upper_container.pack(fill=X, expand=YES)

        # upper container left
        self.upper_container_left = ttk.Frame(master=self.upper_container, padding=2)
        self.upper_container_left.pack(fill=BOTH, expand=YES, side=LEFT)

        # upper container right
        self.upper_container_right = ttk.Frame(master=self.upper_container, padding=2)
        self.upper_container_right.pack(fill=BOTH, expand=YES, side=RIGHT)

        # logs textbox
        self.create_logs()

        # sidebar
        self.create_sidebar()

        # inputs
        self.create_inputs()

        # xlsx stuff
<<<<<<< HEAD
        self.sapo = XLSXCreator(ttkobject=self)
=======
        self.sapo = XLSXCreator(ttkobject=self) #XLSXCreator(tabela_ids_path=None,tabela_ids_col="BPI_UID,C,7", tabela_filtro_path=None, tabela_filtro_col="CODES")
>>>>>>> 3b7c5b82cab30ad17fb8453f88c5091baa5613eb

    def create_logs(self):
        
        # cria a janela de texto de logs
        self.logs_textbox = ScrolledText(
            master=self.upper_container_left,
            highlightcolor=self.style.colors.primary,
            highlightbackground=self.style.colors.border,
            highlightthickness=1,
        )
        self.logs_textbox.pack(fill=X, expand=YES, side=TOP)

        default_txt = "Bem vindo ao SAPO"

        # bota a mensgame la
        self.logs_textbox.insert(END, default_txt)

        # TEM QUE FICA R DISABLED PRA O USUARIO N MUDAR,
        # Se quiser printar algo usa a funcao log_print
        self.logs_textbox.configure(state=DISABLED)
       
    def create_sidebar(self):
        # Create label with frog image
        forg_lbl = ttk.Label(
            master=self.upper_container_right,
            image=self.forg
        )
        forg_lbl.pack(
            side=TOP,
            fill=BOTH,
            expand=YES,
            pady=10,
            padx=(5, 0)
        )

        # Create start button
        start_button = ttk.Button(
            master=self.upper_container_right,
            text="Iniciar",
            command=self.start,
            style='success.TButton'
        )
        start_button.pack(
            side=BOTTOM,
            fill=Y,
            padx=(10, 0)
        )

        # Create separator to separate frog credits
        sep = ttk.Separator(
            master=self.upper_container_right,
            orient=HORIZONTAL
        )
        sep.pack(fill=BOTH, pady=10)

        # Create credits label
        credits_text = '''Criado por:
        
            - Murillo Lucindo 
            - Arthur Fary

            SAPO:

            > Sistema
            > Absolutamente
            > Poderoso de
            > Organização
            
            '''

        credits_text = credits_text.replace('   ', '').strip('\n')

        credits_lbl = ttk.Label(
            master=self.upper_container_right,
            text=credits_text,
            anchor=W,
            style='secondary.TLabel',
        )
        credits_lbl.pack(fill=BOTH, side=TOP)

        # Create separator
        sep = ttk.Separator(
            master=self.upper_container_right,
            orient=HORIZONTAL,
        )
        sep.pack(fill=BOTH, pady=10)

        # Create credits label for MIT license and author
        credits_lbl = ttk.Label(
            master=self.upper_container_right,
            text='MIT License\n\n© 2022 Murillo Lucindo Sanches',
            anchor=W,
            style='secondary.TLabel',
            font=('', 8)
        )
        credits_lbl.pack(fill=BOTH, side=TOP, pady=(10))

    def create_inputs(self):
        # Cria frame para os inputs do filtro
        self.upper_container_left_above = ttk.Frame(master=self.upper_container_left)
        self.upper_container_left_above.pack(fill=X, expand=YES, side=TOP)

        # Botão para selecionar o arquivo de filtro
        browse_filter = ttk.Button(
            master=self.upper_container_left_above, 
            text="Selecionar Filtros", 
            command=self.select_filters_path, 
            width=17
        )
        browse_filter.pack(side=LEFT, padx=(0,10), pady=(10,0))

        # Entrada para o caminho do arquivo de filtro
        file_filter_entry = ttk.Entry(
            master=self.upper_container_left_above, 
            textvariable=self.filter_filename
        )
        file_filter_entry.pack(side=LEFT, fill=BOTH, expand=YES, pady=(10,0))

        # Label para seleção da coluna de filtro
        column_filter_label = ttk.Label(
            master=self.upper_container_left_above, 
            text="Coluna:"
        )
        column_filter_label.pack(side=LEFT, padx=(10), pady=(10,0))

        # Entrada para seleção da coluna de filtro
        self.column_filter_entry = ttk.Entry(
            master=self.upper_container_left_above, 
            width=12
        )
        self.column_filter_entry.pack(side=RIGHT, pady=(10,0))

        # Cria frame para os inputs da tabela
        self.upper_container_left_below = ttk.Frame(master=self.upper_container_left)
        self.upper_container_left_below.pack(fill=X, expand=YES, side=BOTTOM)

        # Botão para selecionar o arquivo de tabela
        browse_table = ttk.Button(
            master=self.upper_container_left_below, 
            text="Selecionar Planilha", 
            command=self.select_table_path, 
            width=17
        )
        browse_table.pack(side=LEFT, padx=(0,10), pady=(10,0))

        # Entrada para o caminho do arquivo de tabela
        file_table_entry = ttk.Entry(
            master=self.upper_container_left_below, 
            textvariable=self.table_filename
        )
        file_table_entry.pack(side=LEFT, fill=X, expand=YES, pady=(10,0))

        # Label para seleção da coluna de tabela
        column_table_label = ttk.Label(
            master=self.upper_container_left_below, 
            text="Coluna:"
        )
        column_table_label.pack(side=LEFT, padx=(10), pady=(10,0))

        # Entrada para seleção da coluna de tabela
        self.column_table_entry = ttk.Entry(
            master=self.upper_container_left_below, 
            width=12
        )
        self.column_table_entry.pack(side=RIGHT, pady=(10,0))

    def select_table_path(self):
        self.table_filename.set(askopenfilename(parent=self, title="Selecione a tabela", filetypes=[("Excel files", "*.xlsx")]))

    def select_filters_path(self):
        self.filter_filename.set(askopenfilename(parent=self, title="Selecione a tabela com os filtros", filetypes=[("Excel files", "*.xlsx")])) 

    def start(self):
        if not self.filter_filename.get():
            self.log_print('Filtros não selecionados.')

        elif not self.column_filter_entry.get():
            self.log_print('Coluna de filtros não selecionada.')

        elif not self.table_filename.get():
            self.log_print('Tabela não selecionada.')

        elif not self.column_table_entry.get():
            self.log_print('Coluna da tabela não selecionada.')

        else:
            self.log_print('Iniciando...')

            self.sapo.set_tabela_filtro_path(self.filter_filename.get())
            self.sapo.set_tabela_filtro_col(self.column_filter_entry.get())
            self.sapo.set_tabela_ids_path(self.table_filename.get())
            self.sapo.set_tabela_ids_col(self.column_table_entry.get())

            
            self.log_print('Carregando e gerando planilhas (pode demorar)...')
            xlsx_start_thread = threading.Thread(target=self.sapo.start)
            xlsx_start_thread.start()


            def keep_alive_start():
                if xlsx_start_thread.is_alive():
                    # schedule another check in 100 ms
                    self.logs_textbox.after(100, keep_alive_start)
                    
                else:
                    # thread has finished, update the status label
                    self.log_print('Planilhas geradas!')

            # check after 100 ms
            self.logs_textbox.after(100, keep_alive_start)
            
    def log_print(self, words: str):
        self.logs_textbox.configure(state=NORMAL)

        words = '\n' + words

        self.logs_textbox.insert(END, words)
        self.logs_textbox.configure(state=DISABLED)

if __name__ == "__main__":
    app = ttk.Window(
        title="SAPO",
        themename="yeti",
        #https://ttkbootstrap.readthedocs.io/en/latest/themes/
        minsize=(1000,300)
    )
    SapoWindow(app)
    app.mainloop()
