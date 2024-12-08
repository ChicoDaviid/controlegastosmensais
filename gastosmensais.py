import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

# Função para adicionar um gasto
def adicionar_gasto():
    data = entrada_data.get()
    descricao = entrada_descricao.get()
    categoria = entrada_categoria.get()
    pagamento = entrada_pagamento.get()
    valor = entrada_valor.get()

    if not (data and descricao and categoria and pagamento and valor):
        messagebox.showwarning("Aviso", "Todos os campos são obrigatórios!")
        return

    try:
        valor = float(valor)
    except ValueError:
        messagebox.showerror("Erro", "O valor deve ser um número!")
        return

    tabela.insert("", "end", values=(data, descricao, categoria, pagamento, f"R$ {valor:.2f}"))

    entrada_data.delete(0, tk.END)
    entrada_descricao.delete(0, tk.END)
    entrada_categoria.delete(0, tk.END)
    entrada_pagamento.delete(0, tk.END)
    entrada_valor.delete(0, tk.END)

# Função para apagar um gasto
def apagar_gasto():
    selecionado = tabela.selection()
    if not selecionado:
        messagebox.showwarning("Aviso", "Nenhum item selecionado!")
        return
    for item in selecionado:
        tabela.delete(item)
    messagebox.showinfo("Sucesso", "Gasto removido com sucesso!")

# Função para salvar os dados em Excel
def salvar_excel():
    items = tabela.get_children()
    if not items:
        messagebox.showwarning("Aviso", "Não há dados para salvar!")
        return

    dados = [tabela.item(item, "values") for item in items]
    colunas = ["Data", "Descrição", "Categoria", "Forma de Pagamento", "Valor"]
    df = pd.DataFrame(dados, columns=colunas)

    try:
        file_path = "Controle_Gastos.xlsx"
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Sucesso", f"Dados salvos com sucesso em {file_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")

# Criar janela principal
janela = tk.Tk()
janela.title("Controle de Gastos Mensais")

# Layout
tk.Label(janela, text="Data (dd/mm/aaaa):").grid(row=0, column=0, padx=10, pady=5)
entrada_data = tk.Entry(janela)
entrada_data.grid(row=0, column=1, padx=10, pady=5)

tk.Label(janela, text="Descrição:").grid(row=1, column=0, padx=10, pady=5)
entrada_descricao = tk.Entry(janela)
entrada_descricao.grid(row=1, column=1, padx=10, pady=5)

tk.Label(janela, text="Categoria:").grid(row=2, column=0, padx=10, pady=5)
entrada_categoria = tk.Entry(janela)
entrada_categoria.grid(row=2, column=1, padx=10, pady=5)

tk.Label(janela, text="Forma de Pagamento:").grid(row=3, column=0, padx=10, pady=5)
entrada_pagamento = tk.Entry(janela)
entrada_pagamento.grid(row=3, column=1, padx=10, pady=5)

tk.Label(janela, text="Valor (R$):").grid(row=4, column=0, padx=10, pady=5)
entrada_valor = tk.Entry(janela)
entrada_valor.grid(row=4, column=1, padx=10, pady=5)

btn_adicionar = tk.Button(janela, text="Adicionar Gasto", command=adicionar_gasto)
btn_adicionar.grid(row=5, column=0, columnspan=2, pady=10)

# Tabela para exibir os gastos
colunas = ["Data", "Descrição", "Categoria", "Forma de Pagamento", "Valor"]
tabela = ttk.Treeview(janela, columns=colunas, show="headings")
for col in colunas:
    tabela.heading(col, text=col)
    tabela.column(col, minwidth=0, width=150)
tabela.grid(row=6, column=0, columnspan=2, pady=10)

btn_salvar = tk.Button(janela, text="Salvar em Excel", command=salvar_excel)
btn_salvar.grid(row=7, column=0, columnspan=2, pady=10)

btn_apagar = tk.Button(janela, text="Apagar Gasto", command=apagar_gasto)
btn_apagar.grid(row=8, column=0, columnspan=2, pady=10)

# Iniciar aplicação
janela.mainloop()
