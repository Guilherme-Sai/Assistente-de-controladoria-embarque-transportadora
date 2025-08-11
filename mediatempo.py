import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from statistics import mean
import pandas as pd
import matplotlib.pyplot as plt

registros = []

def adicionar_registro():
    cedente = entry_cedente.get()
    transportadora = entry_transportadora.get()
    data_emissao = entry_emissao.get()
    data_embarque = entry_embarque.get()
    try:
        emissao = datetime.strptime(data_emissao, "%d/%m/%Y")
        embarque = datetime.strptime(data_embarque, "%d/%m/%Y")
        dias = (embarque - emissao).days
        if dias < 0:
            messagebox.showerror("Erro", "Data de embarque não pode ser anterior à emissão.")
            return
        registro = {
            "cedente": cedente,
            "transportadora": transportadora,
            "emissao": emissao,
            "embarque": embarque,
            "dias": dias
        }
        registros.append(registro)
        atualizar_tabela()
        limpar_campos()
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")

def atualizar_tabela(filtrados=None):
    for i in tabela.get_children():
        tabela.delete(i)
    data = filtrados if filtrados is not None else registros
    for idx, r in enumerate(data):
        tabela.insert("", "end", iid=idx, values=(
            r["cedente"],
            r["transportadora"],
            r["emissao"].strftime("%d/%m/%Y"),
            r["embarque"].strftime("%d/%m/%Y"),
            r["dias"]
        ))

def limpar_campos():
    entry_cedente.delete(0, tk.END)
    entry_transportadora.delete(0, tk.END)
    entry_emissao.delete(0, tk.END)
    entry_embarque.delete(0, tk.END)

def calcular_media():
    cedente_filtro = entry_media_cedente.get().strip().lower()
    if cedente_filtro:
        dias_list = [r["dias"] for r in registros if r["cedente"].strip().lower() == cedente_filtro]
    else:
        dias_list = [r["dias"] for r in registros]
    if not dias_list:
        messagebox.showinfo("Média", "Nenhum registro encontrado.")
        return
    media_dias = mean(dias_list)
    messagebox.showinfo("Média", f"Média de dias: {media_dias:.2f}")

def exportar_excel():
    if not registros:
        messagebox.showinfo("Exportar", "Nenhum dado para exportar.")
        return
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        df = pd.DataFrame([{
            "Cedente": r["cedente"],
            "Transportadora": r["transportadora"],
            "Data de Emissão": r["emissao"].strftime("%d/%m/%Y"),
            "Data de Embarque": r["embarque"].strftime("%d/%m/%Y"),
            "Dias": r["dias"]
        } for r in registros])
        df.to_excel(filepath, index=False)
        messagebox.showinfo("Exportar", f"Dados exportados com sucesso para:\n{filepath}")

def editar_registro():
    selected = tabela.selection()
    if not selected:
        messagebox.showwarning("Editar", "Selecione um registro para editar.")
        return
    idx = int(selected[0])
    cedente = entry_cedente.get()
    transportadora = entry_transportadora.get()
    data_emissao = entry_emissao.get()
    data_embarque = entry_embarque.get()
    try:
        emissao = datetime.strptime(data_emissao, "%d/%m/%Y")
        embarque = datetime.strptime(data_embarque, "%d/%m/%Y")
        dias = (embarque - emissao).days
        if dias < 0:
            messagebox.showerror("Erro", "Data de embarque não pode ser anterior à emissão.")
            return
        registros[idx] = {
            "cedente": cedente,
            "transportadora": transportadora,
            "emissao": emissao,
            "embarque": embarque,
            "dias": dias
        }
        atualizar_tabela()
        limpar_campos()
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")

def deletar_registro():
    selected = tabela.selection()
    if not selected:
        messagebox.showwarning("Deletar", "Selecione um registro para deletar.")
        return
    idx = int(selected[0])
    registros.pop(idx)
    atualizar_tabela()

def selecionar_registro(event):
    selected = tabela.selection()
    if not selected:
        return
    idx = int(selected[0])
    r = registros[idx]
    entry_cedente.delete(0, tk.END)
    entry_cedente.insert(0, r["cedente"])
    entry_transportadora.delete(0, tk.END)
    entry_transportadora.insert(0, r["transportadora"])
    entry_emissao.delete(0, tk.END)
    entry_emissao.insert(0, r["emissao"].strftime("%d/%m/%Y"))
    entry_embarque.delete(0, tk.END)
    entry_embarque.insert(0, r["embarque"].strftime("%d/%m/%Y"))

def filtrar_periodo():
    try:
        inicio = datetime.strptime(entry_inicio.get(), "%d/%m/%Y")
        fim = datetime.strptime(entry_fim.get(), "%d/%m/%Y")
        filtrados = [r for r in registros if inicio <= r["emissao"] <= fim]
        atualizar_tabela(filtrados)
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")

def mostrar_grafico():
    if not registros:
        messagebox.showinfo("Gráfico", "Nenhum dado para exibir.")
        return
    dados = {}
    for r in registros:
        dados.setdefault(r["cedente"], []).append(r["dias"])
    medias = {c: mean(dias) for c, dias in dados.items()}
    plt.bar(medias.keys(), medias.values(), color="skyblue")
    plt.ylabel("Média de Dias")
    plt.title("Tempo Médio por Cedente")
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.show()

# Interface Gráfica
janela = tk.Tk()
janela.title("Controle de Emissão e Embarque")
janela.geometry("1050x650")

# Entrada de dados
frame_top = tk.Frame(janela)
frame_top.pack(pady=10)
tk.Label(frame_top, text="Cedente").grid(row=0, column=0)
entry_cedente = tk.Entry(frame_top, width=20)
entry_cedente.grid(row=0, column=1)
tk.Label(frame_top, text="Transportadora").grid(row=0, column=2)
entry_transportadora = tk.Entry(frame_top, width=20)
entry_transportadora.grid(row=0, column=3)
tk.Label(frame_top, text="Emissão (DD/MM/AAAA)").grid(row=0, column=4)
entry_emissao = tk.Entry(frame_top, width=15)
entry_emissao.grid(row=0, column=5)
tk.Label(frame_top, text="Embarque (DD/MM/AAAA)").grid(row=0, column=6)
entry_embarque = tk.Entry(frame_top, width=15)
entry_embarque.grid(row=0, column=7)

# Botões
frame_buttons = tk.Frame(janela)
frame_buttons.pack()
tk.Button(frame_buttons, text="Adicionar", command=adicionar_registro).grid(row=0, column=0, padx=5)
tk.Button(frame_buttons, text="Editar", command=editar_registro).grid(row=0, column=1, padx=5)
tk.Button(frame_buttons, text="Deletar", command=deletar_registro).grid(row=0, column=2, padx=5)
tk.Button(frame_buttons, text="Exportar Excel", command=exportar_excel).grid(row=0, column=3, padx=5)
tk.Button(frame_buttons, text="Gráfico", command=mostrar_grafico).grid(row=0, column=4, padx=5)

# Tabela
colunas = ("Cedente", "Transportadora", "Emissão", "Embarque/Emissão CTE", "Dias")
tabela = ttk.Treeview(janela, columns=colunas, show="headings", height=15)
for col in colunas:
    tabela.heading(col, text=col)
    tabela.column(col, anchor="center", width=150)
tabela.pack(pady=10, fill="both", expand=True)
tabela.bind("<<TreeviewSelect>>", selecionar_registro)

# Média
frame_media = tk.Frame(janela)
frame_media.pack(pady=5)
tk.Label(frame_media, text="Cedente para média (ou vazio para todos):").grid(row=0, column=0)
entry_media_cedente = tk.Entry(frame_media, width=30)
entry_media_cedente.grid(row=0, column=1)
tk.Button(frame_media, text="Calcular Média", command=calcular_media).grid(row=0, column=2)

# Filtro por período
frame_filtro = tk.Frame(janela)
frame_filtro.pack(pady=10)
tk.Label(frame_filtro, text="Filtrar período - Início (DD/MM/AAAA):").grid(row=0, column=0)
entry_inicio = tk.Entry(frame_filtro, width=15)
entry_inicio.grid(row=0, column=1)
tk.Label(frame_filtro, text="Fim (DD/MM/AAAA):").grid(row=0, column=2)
entry_fim = tk.Entry(frame_filtro, width=15)
entry_fim.grid(row=0, column=3)
tk.Button(frame_filtro, text="Filtrar", command=filtrar_periodo).grid(row=0, column=4)

janela.mainloop()
