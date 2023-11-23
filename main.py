import tkinter as tk
from tkinter import font as tkFont

from tkinter import messagebox
import pandas as pd
from pathlib import Path

EXCEL_FILE = "estoque.xlsx"
# SHEET_NAME = "Estoque"
if not Path(EXCEL_FILE).is_file():
    df = pd.DataFrame(columns=['Numero_Sequencia', 'Nome_Produto', 'Descricao', 'Quantidade', 'Tipo_Peca',
                               'Preco_Unitario', 'Preco_Mercado', 'Imposto_IPI', 'Substituicao_Tributaria'])
    df.to_excel(EXCEL_FILE, index=False)
# comentario p merge

def add_item():
    def submit():
        data = {
            'Numero_Sequencia': entry_numero_seq.get(),
            'Nome_Produto': entry_nome_prod.get(),
            'Descricao': entry_desc.get(),
            'Quantidade': entry_quantidade.get(),
            'Tipo_Peca': entry_tipo_peca.get(),
            'Preco_Unitario': entry_preco_unit.get(),
            'Preco_Mercado': entry_preco_mercado.get(),
            'Imposto_IPI': entry_imposto_ipi.get(),
            'Substituicao_Tributaria': entry_subst_trib.get()
        }

        try:
            df = pd.read_excel(EXCEL_FILE)
            data = pd.DataFrame(data, index=[0])
            df = pd.concat([df, data], ignore_index=True)
            df.to_excel('estoque.xlsx', index=False)

            messagebox.showinfo("Sucesso", "Item adicionado com sucesso")
            add_window.destroy()
        except Exception as e:
            print(e)
            messagebox.showerror("Erro\n", e, "\nDados inválidos")

    add_window = tk.Toplevel()
    center_window(add_window, 420, 380)
    add_window.title("Adicionar Item")

    # Criar campos de entrada
    tk.Label(add_window, text="Número Sequência").grid(row=0, column=0, pady=5)
    entry_numero_seq = tk.Entry(add_window)
    entry_numero_seq.grid(row=0, column=1, pady=5)

    arial_font = tkFont.Font(family="Arial", size=11)

    # Criar campos de entrada
    tk.Label(add_window, text="Número Sequência", font=arial_font).grid(row=0, column=0, pady=5)
    entry_numero_seq = tk.Entry(add_window, font=arial_font)
    entry_numero_seq.grid(row=0, column=1, pady=5)

    tk.Label(add_window, text="Nome do Produto", font=arial_font).grid(row=1, column=0, pady=5)
    entry_nome_prod = tk.Entry(add_window, font=arial_font)
    entry_nome_prod.grid(row=1, column=1, pady=5)

    tk.Label(add_window, text="Descrição", font=arial_font).grid(row=2, column=0, pady=5)
    entry_desc = tk.Entry(add_window, font=arial_font)
    entry_desc.grid(row=2, column=1, pady=5)

    tk.Label(add_window, text="Quantidade", font=arial_font).grid(row=3, column=0, pady=5)
    entry_quantidade = tk.Entry(add_window, font=arial_font)
    entry_quantidade.grid(row=3, column=1, pady=5)

    tk.Label(add_window, text="Tipo da Peça", font=arial_font).grid(row=4, column=0, pady=5)
    entry_tipo_peca = tk.Entry(add_window, font=arial_font)
    entry_tipo_peca.grid(row=4, column=1, pady=5)

    tk.Label(add_window, text="Preço Unitário", font=arial_font).grid(row=5, column=0, pady=5)
    entry_preco_unit = tk.Entry(add_window, font=arial_font)
    entry_preco_unit.grid(row=5, column=1, pady=5)

    tk.Label(add_window, text="Preço de Mercado", font=arial_font).grid(row=6, column=0, pady=5)
    entry_preco_mercado = tk.Entry(add_window, font=arial_font)
    entry_preco_mercado.grid(row=6, column=1, pady=5)

    tk.Label(add_window, text="Imposto IPI", font=arial_font).grid(row=7, column=0, pady=5)
    entry_imposto_ipi = tk.Entry(add_window, font=arial_font)
    entry_imposto_ipi.grid(row=7, column=1, pady=5)

    tk.Label(add_window, text="Substituição Tributária", font=arial_font).grid(row=8, column=0, pady=5)
    entry_subst_trib = tk.Entry(add_window, font=arial_font)
    entry_subst_trib.grid(row=8, column=1, pady=5)

    submit_button = tk.Button(add_window, text="Adicionar", command=submit, font=arial_font)
    submit_button.grid(row=10, column=1, pady=5)


def remove_item():
    arial_font = tkFont.Font(family="Arial", size=11)
    def submit():
        numero_seq = entry_numero_seq.get()
        df = pd.read_excel(EXCEL_FILE)
        df = df[df['Numero_Sequencia'] != int(numero_seq)]
        df.to_excel('estoque.xlsx', index=True)

        messagebox.showinfo("Sucesso", "Item removido com sucesso")
        remove_window.destroy()

    remove_window = tk.Toplevel()
    remove_window.title("Remover Item")

    center_window(remove_window, 300, 85)

    tk.Label(remove_window, text="     Número Sequência", font=arial_font).grid(row=0, column=0)
    entry_numero_seq = tk.Entry(remove_window)
    entry_numero_seq.grid(row=0, column=1, pady=10)

    submit_button = tk.Button(remove_window, text="Remover", command=submit, font=arial_font)
    submit_button.grid(row=1, column=1, pady=10)


def view_items():
    df = pd.read_excel(EXCEL_FILE)
    text = ''
    for row in  range(0, len(df)):
        text += "Num. Seq.: "
        text += f'{df["Numero_Sequencia"][row]}\n'
        text += f'{df["Nome_Produto"][row]}\n'
        text += f'{df["Quantidade"][row]}\n'
        text += f'{df["Preco_Unitario"][row]}\n'
        text += "___________\n"

    messagebox.showinfo(f"Itens no Estoque \n", text)

def center_window(window, width, height):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    window.geometry(f'{width}x{height}+{x}+{y}')

def main():
    window = tk.Tk()
    window.title("Controle de Estoque")

    label = tk.Label(window, text="Bem-vindo ao Sistema de Controle de Estoque")
    label.pack()
    label.config(font=("Arial", 10))

    center_window(window, 300, 400)

    add_button = tk.Button(window, text="Adicionar Item", command=add_item, width=25, height=3)
    add_button.config(font=("Arial", 12))
    add_button.pack(pady=30)

    remove_button = tk.Button(window, text="Remover Item", command=remove_item, width=25, height=3)
    remove_button.config(font=("Arial", 12))
    remove_button.pack(pady=30)

    view_button = tk.Button(window, text="Visualizar Itens", command=view_items, width=25, height=3)
    view_button.config(font=("Arial", 12))
    view_button.pack(pady=30)

    window.mainloop()

if __name__ == "__main__":
    main()
