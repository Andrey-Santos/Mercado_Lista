import locale
import pandas as pd
import tkinter as tk
from tkinter import ttk
from datetime import datetime

# Configurar a localidade para o formato de moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

root = tk.Tk()

class AutocompleteCombobox(ttk.Combobox):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.completion_list = []
        self.current_text = tk.StringVar()
        self.configure(textvariable=self.current_text)
        self.bind("<KeyRelease>", self.autocomplete)

    def set_completion_list(self, completion_list):
        self.completion_list = completion_list

    def autocomplete(self, event):
        current_text = self.current_text.get()
        if current_text == "":
            self.configure(values=self.completion_list)
        else:
            matching_options = [
                option for option in self.completion_list if option.lower().startswith(current_text.lower())
            ]
            self.configure(values=matching_options)

class Funcs():
    def __init__(self):
        self.conexao()

    def limpa_cliente(self):
        self.codigo_entry.delete(0, tk.END)
        self.cidade_entry.delete(0, tk.END)
        self.fone_entry.delete(0, tk.END)
        self.nome_entry.delete(0, tk.END)

    def variaveis(self):
        self.codigo = self.codigo_entry.get()
        self.nome = self.nome_entry.get()
        self.fone = self.fone_entry.get()
        self.cidade = self.cidade_entry.get()
        
    def select_lista(self):
        self.conexao()

        data_fim = datetime.strptime("2023-06-01", "%Y-%m-%d").date()
        data_inicio = datetime.strptime("2023-05-01", "%Y-%m-%d").date()
        colunas_desejadas = ["Fornecedor", "Data Da Compra", "Descrição", "Qtd", "Valor Unitario", "Valor Total", "Quem Paga"]
        
        self.table["Data Da Compra"] = pd.to_datetime(self.table["Data Da Compra"]).dt.date

        # Filtrar os dados de acordo com as datas e colunas desejadas
        lista = self.table.loc[(self.table["Data Da Compra"] > data_inicio) & (self.table["Data Da Compra"] < data_fim), colunas_desejadas]
        lista = lista.sort_values(by="Data Da Compra", ascending=False)

        # Inserir os dados filtrados no Treeview e formatar os valores
        for index, row in lista.iterrows():
            values = row[colunas_desejadas].tolist()
            values[1] = row["Data Da Compra"].strftime("%d/%m/%Y")  # Formatando a data
            values[3] = int(values[3])  
            values[4] = locale.currency(values[4], grouping=True) 
            values[5] = locale.currency(values[5], grouping=True) 
            self.listaCli.insert("", tk.END, values=values)
       
    def conexao(self):
        self.table = pd.read_excel(r"C:\Users\Andrey\Desktop\Mercado\Mercado.xlsx", sheet_name="Compras")
    
class Application(Funcs):
    def __init__(self):
        super().__init__()
        self.root = root
        self.tela()
        self.load_styles()
        self.frames_da_tela()
        self.widgets_frame1()
        self.lista_frame2()
        self.select_lista()
        self.root.mainloop()

    def tela(self):
        self.root.title("Cadastro de Clientes")
        self.root.configure(background= '#474544')
        self.root.geometry("800x500")
        self.root.resizable(True, True)
        self.root.maxsize(width= 900, height= 700)
        self.root.minsize(width=500, height= 400)

    def frames_da_tela(self):
        self.frame_1 = tk.Frame(self.root, bd = 4, bg= '#363636', highlightbackground= 'black', highlightthickness=3 )
        self.frame_1.place(relx= 0.02, rely=0.02, relwidth= 0.96, relheight= 0.46)

        self.frame_2 = tk.Frame(self.root, bd=4, bg='#363636', highlightbackground='black', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)

    def widgets_frame1(self):
        self.combobox = AutocompleteCombobox(self.frame_1)
        self.combobox.set_completion_list(sorted([str(fornecedor) for fornecedor in set(self.table["Fornecedor"])]))
        self.combobox.set("Brasil Atacadista")
        self.combobox.place(relx=0.1, rely=0.1, relwidth=0.2, relheight=0.2)
        #self.combobox.config(bg="#282c34", fg="#ffffff", activebackground="#282c34", activeforeground="#ffffff", bd=1, highlightthickness=1, width=20, height=10)

    def lista_frame2(self):
        self.listaCli = ttk.Treeview(self.frame_2, style="Treeview", height=3)
        self.listaCli["columns"] = [ "Fornecedor", "Data Da Compra", "Descrição", "Qtd", "Valor Unitário", "Valor Total", "Quem Paga"]
        self.listaCli.heading("Fornecedor"    , text="Fornecedor"         )
        self.listaCli.heading("Data Da Compra", text="Data Da Compra")
        self.listaCli.heading("Descrição"     , text="Descrição"     )
        self.listaCli.heading("Qtd"           , text="Qtd"           )
        self.listaCli.heading("Valor Unitário", text="Valor Unitário")
        self.listaCli.heading("Valor Total"   , text="Valor Total"   )
        self.listaCli.heading("Quem Paga"     , text="Quem Paga"     )

        self.listaCli.column("#0"            , width=0  , stretch=tk.NO   )
        self.listaCli.column("Fornecedor"    , width=75 , anchor=tk.CENTER)
        self.listaCli.column("Data Da Compra", width=70 , anchor=tk.CENTER)
        self.listaCli.column("Descrição"     , width=200, anchor=tk.CENTER)
        self.listaCli.column("Qtd"           , width=30 , anchor=tk.CENTER)
        self.listaCli.column("Valor Unitário", width=50 , anchor=tk.CENTER)
        self.listaCli.column("Valor Total"   , width=50 , anchor=tk.CENTER)
        self.listaCli.column("Quem Paga"     , width=50 , anchor=tk.CENTER)

        self.listaCli.place(relx=0.01, rely=0.03, relwidth=0.95, relheight=0.95)

        self.scroolLista = ttk.Scrollbar(self.frame_2, orient='vertical', style="Custom.Vertical.TScrollbar")
        self.scroolLista.configure(command=self.listaCli.yview)
        self.listaCli.configure(yscrollcommand=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.03, relwidth=0.04, relheight=0.95)

    def load_styles(self):
        # Definir estilo personalizado
        style_treeview = ttk.Style()

        # Configurar o tema escuro
        style_treeview.theme_use("alt")

        # Configurar os estilos personalizados
        style_treeview.configure("Treeview",
                        background="#363636",
                        foreground="#ffffff",
                        fieldbackground="#363636")

        style_treeview.configure("Treeview.Heading",
                        background="#363636",
                        foreground="#ffffff")

        style_treeview.configure("Treeview.Item",
                        background="#363636",
                        foreground="#ffffff")
        
        # Definir estilo para a barra de rolagem
        style_scroolbar = ttk.Style()
        style_scroolbar.theme_use('alt')

        # Configurar as cores para o tema escuro
        style_scroolbar.configure("Custom.Vertical.TScrollbar",
                        background="#363636",
                        troughcolor="#3f4451",
                        gripcount=0,
                        darkcolor="#282c34",
                        lightcolor="#282c34",
                        troughrelief="flat",
                        gripmargin=0)

Application()
