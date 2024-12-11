import pandas as pd
import tkinter as tk
from tkinter import messagebox, ttk

def carregar_dados(arquivo):
    try:
        df = pd.read_excel(arquivo)
        if df.empty:
            print("O arquivo está vazio. Criando um novo arquivo.")
            return pd.DataFrame(columns=['Número de Título', 'Nome', 'Telefone', 'Data de entrega'])
        return df
    except FileNotFoundError:
        print("Arquivo não encontrado. Criando um novo arquivo.")
        return pd.DataFrame(columns=['Número de Título', 'Nome', 'Telefone', 'Data de entrega'])
    except ValueError as ve:
        print(f"Erro ao ler o arquivo: {ve}")
        return pd.DataFrame(columns=['Número de Título', 'Nome', 'Telefone', 'Data de entrega'])

def limpar_campos():
    entry_numero_titulo.delete(0, tk.END)
    entry_nome.delete(0, tk.END)
    entry_telefone.delete(0, tk.END)
    entry_data_entrega.delete(0, tk.END)

def adicionar_informacao():
    numero_titulo = entry_numero_titulo.get()
    nome = entry_nome.get()
    telefone = entry_telefone.get()
    data_entrega = entry_data_entrega.get()
    
    if not numero_titulo or not nome or not telefone or not data_entrega:
        messagebox.showwarning("Entrada Inválida", "Por favor, preencha todos os campos.")
        return
    
    nova_linha = pd.DataFrame({
        'Número de Título': [numero_titulo],
        'Nome': [nome],
        'Telefone': [telefone],
        'Data de Entrega': [data_entrega]
    })
    
    global df
    df = pd.concat([df, nova_linha], ignore_index=True)
    salvar_dados()
    messagebox.showinfo("Sucesso", "Informação adicionada com sucesso!")
    limpar_campos() 

def buscar_informacao():
    numero_titulo = entry_buscar.get()
    resultado = df[df['Número de Título'] == numero_titulo]
    
    if not resultado.empty:
        messagebox.showinfo("Resultado da Busca", resultado.to_string(index=False))
    else:
        messagebox.showinfo("Resultado da Busca", "Nenhuma informação encontrada para o número fornecido.")

def salvar_dados():
    df[['Número de Título', 'Nome', 'Telefone', 'Data de Entrega']].to_excel('', index=False)

df = carregar_dados('')

root = tk.Tk()
root.title("Gerenciador de Informações")

frame_adicionar = ttk.LabelFrame(root, text="Adicionar Informações")
frame_adicionar.grid(row=0, column=0, padx=30, pady=30)

ttk.Label(frame_adicionar, text="Nome:").grid(row=0, column=0)
entry_nome = ttk.Entry(frame_adicionar)
entry_nome.grid(row=0, column=1)

ttk.Label(frame_adicionar, text="Número do Título:").grid(row=1, column=0)
entry_numero_titulo = ttk.Entry(frame_adicionar)
entry_numero_titulo.grid(row=1, column=1)

ttk.Label(frame_adicionar, text="Telefone:").grid(row=2, column=0)
entry_telefone = ttk.Entry(frame_adicionar)
entry_telefone.grid(row=2, column=1)

ttk.Label(frame_adicionar, text="Data de Entrega:").grid(row=3, column=0)
entry_data_entrega = ttk.Entry(frame_adicionar)
entry_data_entrega.grid(row=3, column=1)

btn_adicionar = ttk.Button(frame_adicionar, text="Adicionar", command=adicionar_informacao)
btn_adicionar.grid(row=4, columnspan=2, pady=5)

frame_buscar = ttk.LabelFrame(root, text="Buscar Informações")
frame_buscar.grid(row=1, column=0, padx=10, pady=10)

ttk.Label(frame_buscar, text="Número do Título:").grid(row=0, column=0)
entry_buscar = ttk.Entry(frame_buscar)
entry_buscar.grid(row=0, column=1)

btn_buscar = ttk.Button(frame_buscar, text="Buscar", command=buscar_informacao)
btn_buscar.grid(row=1, columnspan=2, pady=5)

root.mainloop()