from tkinter import Tk, Button, Label, PhotoImage, filedialog, messagebox
from utils import obter_caminho_relativo  # Certifique-se de que essa função está disponível
from recibos import gerar_recibos
import os
from tkinter import ttk  # Para a barra de progresso
import threading

def selecionar_arquivo():
    """Abre uma janela para selecionar o arquivo Excel."""
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo dados.xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")],
        initialdir=os.getcwd()
    )
    return arquivo

def abrir_pasta_documentos():
    """Abre a pasta Documentos no explorador de arquivos."""
    pasta = os.path.expanduser("~/Documents")
    os.startfile(pasta)

def gerar_recibos_com_feedback(progress_bar, root):
    """Função para gerar recibos com barra de progresso."""
    arquivo = selecionar_arquivo()
    if not arquivo:
        messagebox.showinfo("Atenção", "Nenhum arquivo selecionado.")
        return

    def task():
        progress_bar["value"] = 0
        progress_bar.grid(row=4, column=0, columnspan=2, pady=10)  # Exibe a barra
        root.update_idletasks()  # Atualiza a interface

        try:
            # Chama a função de gerar recibos e passa o callback para atualizar a barra
            gerar_recibos(arquivo, progress_bar)
            messagebox.showinfo("Sucesso", "Recibos gerados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar recibos: {str(e)}")
        finally:
            progress_bar.grid_remove()  # Esconde a barra novamente
            root.update_idletasks()  # Atualiza a interface

    threading.Thread(target=task).start()  # Executa a tarefa em uma thread separada    

def criar_interface():
    """Cria a interface gráfica."""
    root = Tk()
    root.title("Gerador de Vales em PDF")
    root.geometry("400x250")  # Aumentado para acomodar o rodapé
    root.configure(bg="#37373D")  # Fundo escuro

    # Define o ícone da janela
    caminho_icone = obter_caminho_relativo("assets/janela_icon.png")
    icone_janela = PhotoImage(file=caminho_icone)
    root.iconphoto(True, icone_janela)

    # Funções para o efeito de hover
    def ao_passar(event, botao):
        botao.config(bg="#1F1F1F")

    def ao_sair(event, botao):
        botao.config(bg="#37373D")

    # Carregar ícones dos botões
    caminho_icone_excel = obter_caminho_relativo("assets/excel_icon.png")
    caminho_icone_folder = obter_caminho_relativo("assets/folder_icon.png")
    caminho_icone_close = obter_caminho_relativo("assets/close_icon.png")

    icone_excel = PhotoImage(file=caminho_icone_excel)
    icone_folder = PhotoImage(file=caminho_icone_folder)
    icone_close = PhotoImage(file=caminho_icone_close)

    # Título
    Label(root, text="Vale III PDF", font=("Arial", 16), bg="#37373D", fg="white").pack(pady=10)

    # Botão para selecionar Excel
    botao_excel = Button(root, text=" Selecionar Excel", image=icone_excel, compound="left",
                         relief="flat", bg="#37373D", fg="white", activebackground="#171717", activeforeground="white",
                         command=lambda: gerar_recibos(selecionar_arquivo()))
    botao_excel.pack(pady=5)
    botao_excel.bind("<Enter>", lambda event: ao_passar(event, botao_excel))
    botao_excel.bind("<Leave>", lambda event: ao_sair(event, botao_excel))

    # Botão para abrir Documentos
    botao_folder = Button(root, text=" Abrir Documentos", image=icone_folder, compound="left",
                          relief="flat", bg="#37373D", fg="white", activebackground="#171717", activeforeground="white",
                          command=abrir_pasta_documentos)
    botao_folder.pack(pady=5)
    botao_folder.bind("<Enter>", lambda event: ao_passar(event, botao_folder))
    botao_folder.bind("<Leave>", lambda event: ao_sair(event, botao_folder))

    # Botão para fechar o programa
    botao_fechar = Button(root, text=" Fechar", image=icone_close, compound="left",
                          relief="flat", bg="#37373D", fg="white", activebackground="#171717", activeforeground="white",
                          command=root.quit)
    botao_fechar.pack(pady=5)
    botao_fechar.bind("<Enter>", lambda event: ao_passar(event, botao_fechar))
    botao_fechar.bind("<Leave>", lambda event: ao_sair(event, botao_fechar))

    # Rodapé com o nome do desenvolvedor
    Label(root, text="Desenvolvido por Aureo D. Yamanaka Filho", font=("Arial", 8),
          bg="#37373D", fg="white").pack(side="bottom", pady=5)

    # Necessário para exibir os ícones
    root.mainloop()


