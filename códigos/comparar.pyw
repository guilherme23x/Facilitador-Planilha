import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Função para abrir o arquivo de planilha
def adicionar_planilha():
    # Selecionar o arquivo Excel
    file_path = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if file_path:
        # Adiciona o caminho do arquivo no Textbox
        listbox.insert(ctk.END, file_path + "\n")

# Função para realizar a comparação das planilhas
def comparar_planilhas():
    
    # Recuperar a coluna base e verificar se as planilhas foram carregadas
    coluna_base = entry_coluna.get()
    if not coluna_base:
        messagebox.showerror("Erro", "Por favor, insira o nome da coluna base.")
        return
    
    # Recuperar os arquivos da lista
    arquivos = listbox.get("1.0", ctk.END).strip().split("\n")
    if len(arquivos) < 2:
        messagebox.showerror("Erro", "Adicione pelo menos duas planilhas para comparar.")
        return
    
    # Criar uma lista de DataFrames com as planilhas carregadas
    dfs = []
    for file_path in arquivos:
        try:
            df = pd.read_excel(file_path)
            if coluna_base not in df.columns:
                messagebox.showerror("Erro", f"A coluna '{coluna_base}' não foi encontrada em {file_path}.")
                return
            dfs.append(df)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar {file_path}: {str(e)}")
            return
    
    # Realizar a comparação e mesclar as planilhas
    try:
        base_df = dfs[0]  # DataFrame base
        df_iguais = base_df  # Inicializa o df_iguais com o primeiro DataFrame

        # Percorrer os outros DataFrames e fazer merge
        for df in dfs[1:]:
            # Mescla o DataFrame base com os outros DataFrames usando a coluna base como chave
            df_iguais = pd.merge(df_iguais, df, on=coluna_base, how='outer')  # Merge sem sufixos extras

        # Salvar o DataFrame com os dados mesclados
        df_iguais.to_excel("valores_mesclados.xlsx", index=False)

        messagebox.showinfo("Sucesso", "Comparação concluída! Resultados salvos em 'valores_mesclados.xlsx'.")
    
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro durante a comparação: {str(e)}")

# Interface gráfica
app = ctk.CTk()
app.title("Comparador de Planilhas")
app.geometry("500x450")
app.resizable(False, False)

label_titulo = ctk.CTkLabel(app, text="Comparador", font=("Arial", 24))
label_titulo.pack(pady=10)

label_descricao = ctk.CTkLabel(app, text="Comparador de Planilhas\nSelecione as planilhas e a coluna base")
label_descricao.pack(pady=5)

label_planilhas = ctk.CTkLabel(app, text="Arquivos .xlsx")
label_planilhas.pack(pady=5)

listbox = ctk.CTkTextbox(app, height=100, width=450)
listbox.pack(pady=10)

botao_adicionar = ctk.CTkButton(app, text="Adicionar Planilha", command=adicionar_planilha, fg_color="#21A366", hover_color="#185C37")
botao_adicionar.pack(pady=5)

label_coluna = ctk.CTkLabel(app, text="Nome da Coluna de Base")
label_coluna.pack(pady=5)

entry_coluna = ctk.CTkEntry(app)
entry_coluna.pack(pady=5)

botao_comparar = ctk.CTkButton(app, text="Comparar", command=comparar_planilhas, fg_color="#21A366", hover_color="#185C37")
botao_comparar.pack(pady=10)

app.mainloop()
