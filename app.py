import locale
import pandas as pd
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta 
from tkcalendar import DateEntry
from ttkthemes import ThemedStyle
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
        self["values"] = self.completion_list

    def autocomplete(self, event):
        current_text = self.current_text.get()
        if current_text =="":
            self.configure(values=self.completion_list)
        else:
            matching_options = [
                option for option in self.completion_list if option.lower().startswith(current_text.lower())
            ]
            self.configure(values=matching_options)

class Funcs():
    def __init__(self):
        self.local = r"C:\Users\Andrey\Desktop\Mercado\Mercado.xlsx"
        self.sheet = "Compras"
        self.conexao()
        locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8') # Configurar a localidade para o formato de moeda brasileira

    def conexao(self):
        self.table = pd.read_excel(self.local, sheet_name=self.sheet)

    def adicionar_registro(self):            
        novo_registro = {
            'Fornecedor': self.fornecedor_select.get(),
            'Data Da Compra': self.data_compra_entry.get(),
            'Descrição': self.entry_descricao.get(),
            'Qtd': self.entry_quantidade.get(),
            'Valor Unitario': self.entry_valor.get(),
            'Valor Total': (float(self.entry_valor.get()) * float(self.entry_quantidade.get())),
            'Quem Paga': self.destinatario_select.get()            
        }
        
        self.table = pd.concat([self.table, pd.DataFrame(novo_registro, index=[0])], ignore_index=True)
        self.table.to_excel(self.local, sheet_name=self.sheet, index=False)
        self.select_lista()

    # Define a função para excluir um registro existente
    def excluir_registro(self):
        selection = self.treview_compras.selection()
        if selection:
            item_id = selection[0]
            item_data = self.treview_compras.item(item_id)
            row_data = item_data['values']
            fornecedor = row_data[0]
            data_compra = datetime.strptime(row_data[1], "%d/%m/%Y").date()
            descricao = row_data[2]
            qtd = int(row_data[3])
            valor_unitario = float(row_data[4].replace('R$', '').replace('.', '').replace(',', '.'))
            valor_total = float(row_data[5].replace('R$', '').replace('.', '').replace(',', '.'))
            quem_paga = row_data[6]

            # Remover o registro do DataFrame
            mask = ((self.table['Fornecedor'    ] == fornecedor    ) & 
                    (self.table['Data Da Compra'] == data_compra   ) &
                    (self.table['Descrição'     ] == descricao     ) &
                    (self.table['Qtd'           ] == qtd           ) &
                    (self.table['Valor Unitario'] == valor_unitario) &
                    (self.table['Valor Total'   ] == valor_total   ) &
                    (self.table['Quem Paga'     ] == quem_paga     ))
            
            self.table = self.table.loc[~mask]
            self.table.to_excel(self.local, sheet_name=self.sheet, index=False)
            self.select_lista()

    def alterar_registro(self):
        selection = self.treview_compras.selection()
        if False:
            item_id = selection[0]
            item_data = self.treview_compras.item(item_id)
            row_data = item_data['values']
            fornecedor = row_data[0]
            data_compra = datetime.strptime(row_data[1], "%d/%m/%Y").date()
            descricao = row_data[2]
            qtd = int(row_data[3])
            valor_unitario = float(row_data[4].replace('R$', '').replace('.', '').replace(',', '.'))
            valor_total = float(row_data[5].replace('R$', '').replace('.', '').replace(',', '.'))
            quem_paga = row_data[6]

            # Realize as alterações necessárias nos valores do registro
            # Depois, salve o DataFrame atualizado no arquivo Excel novamente
            self.table.loc[(self.table['Fornecedor'] == fornecedor) &
                        (self.table['Data Da Compra'] == data_compra) &
                        (self.table['Descrição'] == descricao) &
                        (self.table['Qtd'] == qtd) &
                        (self.table['Valor Unitario'] == valor_unitario) &
                        (self.table['Valor Total'] == valor_total) &
                        (self.table['Quem Paga'] == quem_paga), 'Nova Coluna'] = 'Novo Valor'

            self.table.to_excel(self.local, sheet_name=self.sheet, index=False)
            self.select_lista()


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
        data_inicio       = datetime.strptime(self.data_intervalo_ini.get(), "%d/%m/%Y").date()
        data_fim          = datetime.strptime(self.data_intervalo_fim.get(), "%d/%m/%Y").date()
        colunas_desejadas = ["Fornecedor", "Data Da Compra", "Descrição", "Qtd", "Valor Unitario", "Valor Total", "Quem Paga"]
        
        self.table["Data Da Compra"] = pd.to_datetime(self.table["Data Da Compra"], format="%d/%m/%Y").dt.date

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
            self.treview_compras.insert("", tk.END, values=values)

    def validate_input_numeric(self,new_value):
        if all(char.isdigit() or char == "," for char in new_value):
            return True
        else:
            return False
        
    def get_last_day_of_month(self):
        today = datetime.today()
        last_day = today.replace(day=28) + timedelta(days=4) # Considera o dia 28 para garantir a inclusão do último dia do mês
        return last_day - timedelta(days=last_day.day)
    
class Application(Funcs):
    def __init__(self):
        super().__init__()
        self.root = root
        self.tela()
        self.load_styles()
        self.frames_da_tela()
        self.widgets_frame1()
        self.widgets_frame_filtros()
        self.lista_frame2()
        self.select_lista()
        self.root.mainloop()

    def tela(self):
        self.root.title("Cadastro de Clientes")
        self.root.configure(background='#474544')
        self.root.geometry("800x500")
        self.root.resizable(True, True)
        self.root.maxsize(width= 900, height= 700)
        self.root.minsize(width=500, height= 400)

    def frames_da_tela(self):
        self.frame_1 = tk.Frame(self.root, bd = 4, bg= '#363636', highlightbackground= 'black', highlightthickness=3 )
        self.frame_1.place(relx= 0.02, rely=0.02, relwidth= 0.96, relheight= 0.46)

        self.frame_filtros = tk.Frame(self.frame_1, bd = 4, bg= '#363636', highlightbackground= 'white', highlightthickness=1 )
        self.frame_filtros.place(relx= 0.6, rely=0.05, relwidth= 0.35, relheight= 0.6)

        self.frame_2 = tk.Frame(self.root, bd=4, bg='#363636', highlightbackground='black', highlightthickness=3)
        self.frame_2.place(relx=0.02, rely=0.5, relwidth=0.96, relheight=0.46)

    def widgets_frame1(self):   
        self.label_quantidade = tk.Label(self.frame_1, text="Quantidade:", bg="#363636", fg="white")   
        self.label_quantidade.place(relx=0.03, rely=0.06)
        
        self.entry_quantidade = ttk.Entry(self.frame_1, style="TEntry", validate="key", validatecommand=(root.register(self.validate_input_numeric), "%P"))
        self.entry_quantidade.place(relx=0.03, rely=0.15, relwidth=0.10, relheight=0.12)

        self.label_valor = tk.Label(self.frame_1, text="Valor R$:", bg="#363636", fg="white")   
        self.label_valor.place(relx=0.17, rely=0.06)

        self.entry_valor = ttk.Entry(self.frame_1, style="TEntry", validate="key", validatecommand=(root.register(self.validate_input_numeric), "%P"))
        self.entry_valor.place(relx=0.17, rely=0.15, relwidth=0.10, relheight=0.12)

        self.label_valor = tk.Label(self.frame_1, text="Destinatário:", bg="#363636", fg="white")   
        self.label_valor.place(relx=0.31, rely=0.06)

        self.destinatario_select = ttk.Combobox(self.frame_1)
        self.destinatario_select.set("TODOS")
        self.destinatario_select["values"] = ["TODOS", "ANDREY", "ANDRIELLY"]
        self.destinatario_select.place(relx=0.31, rely=0.15, relwidth=0.12, relheight=0.12)

        self.label_descricao = tk.Label(self.frame_1, text="Descrição:", bg="#363636", fg="white")   
        self.label_descricao.place(relx=0.03, rely=0.30)

        self.entry_descricao = ttk.Entry(self.frame_1, style="TEntry")
        self.entry_descricao.place(relx=0.03, rely=0.39, relwidth=0.44, relheight=0.12)

        self.label_fornecedor = tk.Label(self.frame_1, text="Fornecedor:", bg="#363636", fg="white")   
        self.label_fornecedor.place(relx=0.03, rely=0.56)

        self.fornecedor_select = AutocompleteCombobox(self.frame_1)
        self.fornecedor_select.set("Brasil Atacadista")
        self.fornecedor_select.set_completion_list(sorted([str(fornecedor) for fornecedor in set(self.table["Fornecedor"])]))
        self.fornecedor_select.place(relx=0.03, rely=0.65, relwidth=0.22, relheight=0.12)

        self.label_data_compra = tk.Label(self.frame_1, text="Data da compra:", bg="#363636", fg="white")   
        self.label_data_compra.place(relx=0.3, rely=0.56)

        self.data_compra_entry = DateEntry(self.frame_1, date_pattern="dd/mm/yyyy", width=12, background="white", foreground="black")
        self.data_compra_entry.place(relx=0.3, rely=0.65, relwidth=0.22, relheight=0.12)

        self.button_adicionar = ttk.Button(self.frame_1, style="TButton",text="Adicionar", command=self.adicionar_registro)
        self.button_adicionar.place(relx=0.03, rely=0.85, relwidth=0.10, relheight=0.13)

        self.button_alterar = ttk.Button(self.frame_1, style="TButton",text="Alterar", command=self.alterar_registro)
        self.button_alterar.place(relx=0.14, rely=0.85, relwidth=0.10, relheight=0.13)

        self.button_excluir = ttk.Button(self.frame_1, style="TButton",text="Excluir", command=self.excluir_registro)
        self.button_excluir.place(relx=0.25, rely=0.85, relwidth=0.10, relheight=0.13)

        self.button_atualizar = ttk.Button(self.frame_1, style="TButton",text="Atualizar", command=self.select_lista)
        self.button_atualizar.place(relx=0.36, rely=0.85, relwidth=0.10, relheight=0.13)

    def widgets_frame_filtros(self):
        self.label_quantidade = tk.Label(self.frame_filtros, text="Filtros", bg="#363636", fg="white",font=("Helvetica", 11, "bold"))   
        self.label_quantidade.place(relx=0.35, rely=0.06)

        self.label_data_compra = tk.Label(self.frame_filtros, text="Intervalo Desejado", bg="#363636", fg="white")   
        self.label_data_compra.place(relx=0.3, rely=0.4)

        self.data_intervalo_ini = DateEntry(self.frame_filtros, date_pattern="dd/mm/yyyy", width=12, background="white", foreground="black")
        self.data_intervalo_ini.place(relx=0.05, rely=0.59, relwidth=0.4, relheight=0.22)

        self.data_intervalo_fim = DateEntry(self.frame_filtros, date_pattern="dd/mm/yyyy", width=12, background="white", foreground="black")
        self.data_intervalo_fim.place(relx=0.55, rely=0.59, relwidth=0.4, relheight=0.22)
        self.data_intervalo_fim.set_date(self.get_last_day_of_month())

    def lista_frame2(self):
        self.treview_compras = ttk.Treeview(self.frame_2, style="Treeview", height=3)
        self.treview_compras["columns"] = [ "Fornecedor", "Data Da Compra", "Descrição", "Qtd", "Valor Unitário", "Valor Total", "Quem Paga"]
        self.treview_compras.heading("Fornecedor"    , text="Fornecedor"    )
        self.treview_compras.heading("Data Da Compra", text="Data Da Compra")
        self.treview_compras.heading("Descrição"     , text="Descrição"     )
        self.treview_compras.heading("Qtd"           , text="Qtd"           )
        self.treview_compras.heading("Valor Unitário", text="Valor Unitário")
        self.treview_compras.heading("Valor Total"   , text="Valor Total"   )
        self.treview_compras.heading("Quem Paga"     , text="Quem Paga"     )

        self.treview_compras.column("#0"            , width=0  , stretch=tk.NO   )
        self.treview_compras.column("Fornecedor"    , width=75 , anchor=tk.CENTER)
        self.treview_compras.column("Data Da Compra", width=70 , anchor=tk.CENTER)
        self.treview_compras.column("Descrição"     , width=200, anchor=tk.CENTER)
        self.treview_compras.column("Qtd"           , width=30 , anchor=tk.CENTER)
        self.treview_compras.column("Valor Unitário", width=50 , anchor=tk.CENTER)
        self.treview_compras.column("Valor Total"   , width=50 , anchor=tk.CENTER)
        self.treview_compras.column("Quem Paga"     , width=50 , anchor=tk.CENTER)

        self.treview_compras.place(relx=0.01, rely=0.03, relwidth=0.95, relheight=0.95)

        self.scroolLista = ttk.Scrollbar(self.frame_2, orient='vertical', style="Custom.Vertical.TScrollbar")
        self.scroolLista.configure(command=self.treview_compras.yview)
        self.treview_compras.configure(yscrollcommand=self.scroolLista.set)
        self.scroolLista.place(relx=0.96, rely=0.03, relwidth=0.04, relheight=0.95)

    def load_styles(self):
        entry_style = ThemedStyle(root)
        entry_style.set_theme("alt")
        entry_style.configure("TEntry",
                              fieldbackground="#474544",
                              background="#474544",
                              foreground="white",
                              insertbackground="white")
        
        combo_style = ttk.Style()
        combo_style.theme_use('alt')

        combo_style.configure("TCombobox",
                              background="#474544",
                              fieldbackground="#474544",
                              foreground="white",
                              selectbackground="gray")
        
        button_style = ThemedStyle(root)
        button_style.set_theme("alt")
        button_style.configure("TButton",
                               background="#474544",
                               foreground="white",
                               borderwidth=0,
                               highlightthickness=0,
                               font=("Helvetica", 10))

        style_treeview = ttk.Style()
        style_treeview.theme_use("alt")
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
        
        style_scroolbar = ttk.Style()
        style_scroolbar.theme_use('alt')
        style_scroolbar.configure("Custom.Vertical.TScrollbar",
                                  background="#363636",
                                  troughcolor="#474544",
                                  gripcount=0,
                                  darkcolor="#363636",
                                  lightcolor="#474544",
                                  troughrelief="flat",
                                  gripmargin=0)
               
Application()