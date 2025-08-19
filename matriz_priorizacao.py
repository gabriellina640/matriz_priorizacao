import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys

# ===============================
# Configurações iniciais
# ===============================
PONTUACAO = {
    'Muito Alto': 10,
    'Alto': 8,
    'Médio': 6,
    'Baixo': 4,
    'Muito Baixo': 2,
}
CATEGORIAS = list(PONTUACAO.keys())

# Pesos iniciais
PESO_IMPACTO = 3
PESO_URGENCIA = 2
PESO_FACILIDADE = 1
PESO_NECESSIDADE = 2

# ===============================
# Caminho seguro para qualquer PC
# ===============================
if getattr(sys, 'frozen', False):
    # Quando rodando como .exe
    base_path = os.path.join(os.environ['USERPROFILE'], 'matriz_priorizacao')
else:
    # Quando rodando como script Python
    base_path = os.path.dirname(os.path.abspath(__file__))

os.makedirs(base_path, exist_ok=True)
data_file = os.path.join(base_path, 'data.xlsx')

# ===============================
# Criar ou carregar DataFrame
# ===============================
if os.path.exists(data_file):
    df = pd.read_excel(data_file)
else:
    # Cria DataFrame vazio com colunas necessárias
    df = pd.DataFrame(columns=['ID','Item','Impacto','Urgência','Facilidade Técnica','Necessidade'])
    df.to_excel(data_file, index=False)

# ===============================
# Função de cálculo de prioridade
# ===============================
def calcular_prioridade(df, pesos=None):
    if pesos is None:
        pesos = {'Impacto':PESO_IMPACTO, 'Urgência':PESO_URGENCIA,
                 'Facilidade Técnica':PESO_FACILIDADE, 'Necessidade':PESO_NECESSIDADE}
    out = df.copy()
    for col in ['Impacto','Urgência','Facilidade Técnica','Necessidade']:
        out[f'Nota {col}'] = out[col].map(PONTUACAO)
    out['Prioridade'] = (
        out['Nota Impacto']*pesos['Impacto'] +
        out['Nota Urgência']*pesos['Urgência'] +
        out['Nota Facilidade Técnica']*pesos['Facilidade Técnica'] +
        out['Nota Necessidade']*pesos['Necessidade']
    )
    out = out.sort_values('Prioridade', ascending=False).reset_index(drop=True)
    return out

# ===============================
# Interface Tkinter
# ===============================
root = tk.Tk()
root.title("Matriz de Priorização")
root.configure(bg='white')

style = ttk.Style()
style.theme_use('clam')
style.configure("Treeview", background="white", foreground="black", rowheight=30, fieldbackground="white", font=('Helvetica', 11))
style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'), background="#2E7D32", foreground="white")
style.map('Treeview', background=[('selected', '#A5D6A7')])
style.configure('TButton', font=('Helvetica',11,'bold'), background='#2E7D32', foreground='white')
style.map('TButton', background=[('active','#1B5E20')])
style.configure('TLabel', background='white')
style.configure('TEntry', font=('Helvetica',11))
style.configure('TCombobox', font=('Helvetica',11))

cols = ['ID','Item','Impacto','Urgência','Facilidade Técnica','Necessidade','Prioridade']
tree = ttk.Treeview(root, columns=cols, show='headings', height=10)
for c in cols:
    tree.heading(c, text=c)
    tree.column(c, width=140, anchor='center')
tree.pack(pady=20, padx=20)

# ===============================
# Pesos ajustáveis
# ===============================
frame_pesos = tk.Frame(root, bg='white')
frame_pesos.pack(pady=5)

peso_impacto = tk.IntVar(value=PESO_IMPACTO)
peso_urgencia = tk.IntVar(value=PESO_URGENCIA)
peso_facilidade = tk.IntVar(value=PESO_FACILIDADE)
peso_necessidade = tk.IntVar(value=PESO_NECESSIDADE)

def pesos_alterados(*args):
    refresh_table()

peso_impacto.trace_add('write', pesos_alterados)
peso_urgencia.trace_add('write', pesos_alterados)
peso_facilidade.trace_add('write', pesos_alterados)
peso_necessidade.trace_add('write', pesos_alterados)

tk.Label(frame_pesos, text='Peso Impacto', bg='white').grid(row=0, column=0, padx=5)
tk.Entry(frame_pesos, textvariable=peso_impacto, width=5).grid(row=1, column=0, padx=5)

tk.Label(frame_pesos, text='Peso Urgência', bg='white').grid(row=0, column=1, padx=5)
tk.Entry(frame_pesos, textvariable=peso_urgencia, width=5).grid(row=1, column=1, padx=5)

tk.Label(frame_pesos, text='Peso Facilidade', bg='white').grid(row=0, column=2, padx=5)
tk.Entry(frame_pesos, textvariable=peso_facilidade, width=5).grid(row=1, column=2, padx=5)

tk.Label(frame_pesos, text='Peso Necessidade', bg='white').grid(row=0, column=3, padx=5)
tk.Entry(frame_pesos, textvariable=peso_necessidade, width=5).grid(row=1, column=3, padx=5)

# ===============================
# Atualiza tabela
# ===============================
def refresh_table():
    pesos = {
        'Impacto': peso_impacto.get(),
        'Urgência': peso_urgencia.get(),
        'Facilidade Técnica': peso_facilidade.get(),
        'Necessidade': peso_necessidade.get()
    }
    df_ordenado = calcular_prioridade(df, pesos)
    for i in tree.get_children():
        tree.delete(i)
    for idx, row in df_ordenado.iterrows():
        tree.insert('', 'end', values=(row['ID'], row['Item'], row['Impacto'], row['Urgência'], row['Facilidade Técnica'], row['Necessidade'], row['Prioridade']))

refresh_table()

# ===============================
# Funções de salvar/adicionar/excluir
# ===============================
def salvar_dados():
    df.to_excel(data_file, index=False)

def adicionar():
    top = tk.Toplevel(root)
    top.title("Adicionar Projeto")
    top.configure(bg='white')

    tk.Label(top, text="Item:", bg='white').grid(row=0, column=0, padx=5, pady=5)
    entry_item = tk.Entry(top)
    entry_item.grid(row=0, column=1, padx=5, pady=5)

    var_impacto = tk.StringVar(value='Médio')
    tk.Label(top, text="Impacto:", bg='white').grid(row=1, column=0, padx=5, pady=5)
    ttk.Combobox(top, textvariable=var_impacto, values=CATEGORIAS, state='readonly').grid(row=1, column=1, padx=5, pady=5)

    var_urgencia = tk.StringVar(value='Médio')
    tk.Label(top, text="Urgência:", bg='white').grid(row=2, column=0, padx=5, pady=5)
    ttk.Combobox(top, textvariable=var_urgencia, values=CATEGORIAS, state='readonly').grid(row=2, column=1, padx=5, pady=5)

    var_facilidade = tk.StringVar(value='Médio')
    tk.Label(top, text="Facilidade Técnica:", bg='white').grid(row=3, column=0, padx=5, pady=5)
    ttk.Combobox(top, textvariable=var_facilidade, values=CATEGORIAS, state='readonly').grid(row=3, column=1, padx=5, pady=5)

    var_necessidade = tk.StringVar(value='Médio')
    tk.Label(top, text="Necessidade:", bg='white').grid(row=4, column=0, padx=5, pady=5)
    ttk.Combobox(top, textvariable=var_necessidade, values=CATEGORIAS, state='readonly').grid(row=4, column=1, padx=5, pady=5)

    def salvar():
        global df
        item = entry_item.get().strip()
        if not item:
            messagebox.showerror("Erro", "Informe o nome do projeto.")
            return
        novo = pd.DataFrame([{
            'ID': df['ID'].max()+1 if not df.empty else 1,
            'Item': item,
            'Impacto': var_impacto.get(),
            'Urgência': var_urgencia.get(),
            'Facilidade Técnica': var_facilidade.get(),
            'Necessidade': var_necessidade.get()
        }])
        df = pd.concat([df, novo], ignore_index=True)
        salvar_dados()
        refresh_table()
        top.destroy()

    ttk.Button(top, text="Salvar", command=salvar).grid(row=5, column=0, pady=10, padx=5)
    ttk.Button(top, text="Cancelar", command=top.destroy).grid(row=5, column=1, pady=10, padx=5)

def excluir():
    top = tk.Toplevel(root)
    top.title("Excluir Projeto")
    top.configure(bg='white')

    tk.Label(top, text="Selecione o ID:", bg='white').grid(row=0, column=0, padx=5, pady=5)
    var_id = tk.IntVar()
    ids = df['ID'].tolist()
    ttk.Combobox(top, textvariable=var_id, values=ids, state='readonly').grid(row=0, column=1, padx=5, pady=5)

    def confirmar():
        global df
        df = df[df['ID'] != var_id.get()].reset_index(drop=True)
        salvar_dados()
        refresh_table()
        top.destroy()

    ttk.Button(top, text="Excluir", command=confirmar).grid(row=1, column=0, pady=10, padx=5)
    ttk.Button(top, text="Cancelar", command=top.destroy).grid(row=1, column=1, pady=10, padx=5)

# ===============================
# Botões principais
# ===============================
frame_btn = tk.Frame(root, bg='white')
frame_btn.pack(pady=10)
ttk.Button(frame_btn, text="Adicionar Projeto", command=adicionar).grid(row=0, column=0, padx=10)
ttk.Button(frame_btn, text="Excluir Projeto", command=excluir).grid(row=0, column=1, padx=10)

# ===============================
# Botão salvar Excel com diálogo
# ===============================
def salvar_excel():
    global df
    if df.empty:
        messagebox.showwarning("Aviso", "Não há dados para salvar!")
        return

    pesos = {
        'Impacto': peso_impacto.get(),
        'Urgência': peso_urgencia.get(),
        'Facilidade Técnica': peso_facilidade.get(),
        'Necessidade': peso_necessidade.get()
    }
    df_ordenado = calcular_prioridade(df, pesos)

    caminho = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salvar arquivo como"
    )
    if caminho:
        df_ordenado.to_excel(caminho, index=False)
        messagebox.showinfo("Salvo!", f"Arquivo salvo em:\n{caminho}")

ttk.Button(root, text="Salvar Excel", command=salvar_excel).pack(pady=10)

root.mainloop()
