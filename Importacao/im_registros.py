import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def agrupar_planilha(nome_planilha, local_salvar, progress_bar, status_label):
    # Carregar a planilha
    planilha = pd.read_excel(nome_planilha)
    
    # Verificar duplicatas na primeira coluna e agrupar
    grupos = planilha.groupby(planilha.iloc[:, 0])
    
    # Criar uma pasta para salvar as planilhas geradas
    if not os.path.exists(local_salvar):
        os.makedirs(local_salvar)
    
    total_grupos = len(grupos)
    progress_bar["maximum"] = total_grupos
    
    # Iterar sobre os grupos e salvar cada um em uma nova planilha
    for i, (nome, grupo) in enumerate(grupos, 1):
        nome_sanitizado = str(nome).strip().replace(' ', '_')  # Sanitizar o nome do arquivo
        nome_arquivo = os.path.join(local_salvar, f"{nome_sanitizado}.xlsx")
        # Criar diretório se não existir
        os.makedirs(os.path.dirname(nome_arquivo), exist_ok=True)
        grupo.to_excel(nome_arquivo, index=False)
        
        # Atualizar a barra de progresso
        progress_bar["value"] = i
        status_label.config(text=f"Processando {i} de {total_grupos}...")
        root.update_idletasks()
    
    status_label.config(text="Concluído!")
    messagebox.showinfo("Sucesso", "Planilhas geradas com sucesso!")

def selecionar_arquivo():
    arquivo_path = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if arquivo_path:
        entry_arquivo.delete(0, tk.END)
        entry_arquivo.insert(0, arquivo_path)

def selecionar_pasta():
    pasta_path = filedialog.askdirectory(title="Selecione a pasta para salvar os arquivos")
    if pasta_path:
        entry_pasta.delete(0, tk.END)
        entry_pasta.insert(0, pasta_path)

def executar():
    nome_planilha = entry_arquivo.get()
    local_salvar = entry_pasta.get()
    
    if not nome_planilha or not local_salvar:
        messagebox.showerror("Erro", "Por favor, selecione a planilha e o local para salvar os arquivos.")
        return
    
    progress_bar["value"] = 0
    status_label.config(text="Iniciando...")
    root.update_idletasks()
    agrupar_planilha(nome_planilha, local_salvar, progress_bar, status_label)

def mostrar_tela_principal():
    instrucoes_frame.pack_forget()
    principal_frame.pack()

def encerrar_programa():
    root.destroy()

# Criar a interface gráfica
root = tk.Tk()
root.title("Agrupador de Planilhas Excel")

# Tela inicial com instruções
instrucoes_frame = tk.Frame(root)
tk.Label(instrucoes_frame, text="**** ANALISE DE ARQUIVO EM EXCEL ****", font=('Helvetica', 13, 'bold')).pack(pady=10)
tk.Label(instrucoes_frame, text="Esse programa vai analisar itens que se repetem em determinada coluna, vai agrupar os registros referente a esse item e gerar uma nova planilha em Excel.", wraplength=400, justify="left").pack(pady=5)
tk.Label(instrucoes_frame, text="1 - Selecione a planilha que contem os registros.", wraplength=400, justify="left", font=('Helvetica', 10, 'bold')).pack(pady=5)
tk.Label(instrucoes_frame, text="2 - Escolha a pasta que deseja salvar os arquivos gerados.", wraplength=400, justify="left", font=('Helvetica', 10, 'bold')).pack(pady=5)
tk.Label(instrucoes_frame, text="3 - Fecha a janela da aplicação ao finalizar e confira a extração.", wraplength=400, justify="left", font=('Helvetica', 10, 'bold')).pack(pady=5)
tk.Label(instrucoes_frame, text="Desenvolvido por: Samuel Santos.", wraplength=400, justify="left", font=('Helvetica', 10)).pack(pady=10)
btn_avancar = tk.Button(instrucoes_frame, text="Avançar", command=mostrar_tela_principal)
btn_avancar.pack(side="left", padx=10, pady=20)
btn_encerrar = tk.Button(instrucoes_frame, text="Encerrar", command=encerrar_programa)
btn_encerrar.pack(side="right", padx=10, pady=20)
instrucoes_frame.pack()

# Tela principal para seleção de arquivos e execução
principal_frame = tk.Frame(root)

tk.Label(principal_frame, text="Selecione a planilha Excel:").grid(row=0, column=0, padx=10, pady=10)
entry_arquivo = tk.Entry(principal_frame, width=50)
entry_arquivo.grid(row=0, column=1, padx=10, pady=10)
btn_arquivo = tk.Button(principal_frame, text="Buscar", command=selecionar_arquivo)
btn_arquivo.grid(row=0, column=2, padx=10, pady=10)

tk.Label(principal_frame, text="Selecione o local para salvar os arquivos:").grid(row=1, column=0, padx=10, pady=10)
entry_pasta = tk.Entry(principal_frame, width=50)
entry_pasta.grid(row=1, column=1, padx=10, pady=10)
btn_pasta = tk.Button(principal_frame, text="Buscar", command=selecionar_pasta)
btn_pasta.grid(row=1, column=2, padx=10, pady=10)

btn_executar = tk.Button(principal_frame, text="Executar", command=executar)
btn_executar.grid(row=3, column=1, padx=10, pady=20)

progress_bar = ttk.Progressbar(principal_frame, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

status_label = tk.Label(principal_frame, text="")
status_label.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# Iniciar o loop da interface gráfica
root.mainloop()




