import pandas as pd
from jinja2 import Template
from weasyprint import HTML
import os
from tkinter import Tk, Button, Label, filedialog, messagebox
from datetime import datetime
import base64
import locale
import sys
import subprocess

# Configura o locale para formato monetário brasileiro
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')

def formatar_moeda(valor):
    """Formata um número como moeda brasileira ou retorna '-' para valores ausentes."""
    if pd.isna(valor) or valor == 0:  # Substituir NaN ou valores zerados
        return "-"
    return locale.currency(valor, grouping=True, symbol=True).replace(' ', '')

def imagem_para_base64(caminho_imagem):
    """Converte uma imagem em base64 para uso direto no HTML."""
    with open(caminho_imagem, "rb") as img:
        return f"data:image/png;base64,{base64.b64encode(img.read()).decode()}"

def obter_caminho_relativo(caminho_arquivo):
    """Resolve o caminho correto para arquivos ao empacotar com PyInstaller."""
    if getattr(sys, 'frozen', False):  # Executando como executável
        caminho_base = sys._MEIPASS
    else:  # Executando como script Python normal
        caminho_base = os.path.dirname(__file__)
    return os.path.join(caminho_base, caminho_arquivo)

def obter_pasta_documentos():
    """Retorna o caminho da pasta Documentos do usuário."""
    return os.path.join(os.path.expanduser("~"), "Documents")

def abrir_pasta_documentos():
    """Abre a pasta Documentos no explorador de arquivos."""
    pasta_documentos = obter_pasta_documentos()
    subprocess.Popen(f'explorer "{pasta_documentos}"')

def selecionar_arquivo():
    """Abre uma janela para selecionar o arquivo Excel."""
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo dados.xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")],
        initialdir=os.getcwd()  # Pasta inicial é onde está o programa
    )
    return arquivo

def gerar_recibos(caminho_planilha):
    try:
        # Caminho para os arquivos
        caminho_html = obter_caminho_relativo('recibo.html')
        caminho_css = obter_caminho_relativo('bootstrap.min.css')
        caminho_logo = obter_caminho_relativo('logo.png')

        # Verificar se os arquivos necessários existem
        if not os.path.exists(caminho_html):
            messagebox.showerror("Erro", "O arquivo 'recibo.html' não foi encontrado.")
            return
        if not os.path.exists(caminho_css):
            messagebox.showerror("Erro", "O arquivo 'bootstrap.min.css' não foi encontrado.")
            return
        if not os.path.exists(caminho_logo):
            messagebox.showerror("Erro", "O arquivo 'logo.png' não foi encontrado.")
            return

        # Ler o HTML e substituir o caminho do CSS para absoluto
        with open(caminho_html, 'r', encoding='utf-8') as arquivo_html:
            template_html = arquivo_html.read()
        
        # Substituir o caminho relativo pelo absoluto no HTML
        template_html = template_html.replace(
            '<link rel="stylesheet" href="bootstrap.min.css">',
            f'<link rel="stylesheet" href="file://{caminho_css}">'
        )

        # Ler os dados da planilha
        dados = pd.read_excel(caminho_planilha)

        # Filtrar funcionários sem valores
        campos_valores = ['convenio_medico', 'uniodonto', 'refeicao', 'farmacia']
        dados['tem_valores'] = dados[campos_valores].apply(
            lambda linha: any(not pd.isna(valor) and valor != 0 for valor in linha), axis=1
        )
        dados = dados[dados['tem_valores']]

        if dados.empty:
            messagebox.showinfo("Informação", "Nenhum funcionário possui valores para gerar recibos.")
            return

        # Criar a pasta de saída na pasta Documentos
        data_atual = datetime.now().strftime("%d.%m.%Y")
        pasta_saida = os.path.join(obter_pasta_documentos(), f"{data_atual}-recibos")
        if not os.path.exists(pasta_saida):
            os.makedirs(pasta_saida)

        # Gerar os PDFs para cada funcionário
        for _, linha in dados.iterrows():
            template = Template(template_html)
            caminho_logo_base64 = imagem_para_base64(caminho_logo)
            html_preenchido = template.render(
                nome_funcionario=linha['Nome do Funcionário'],
                convenio_medico=formatar_moeda(linha['convenio_medico']),
                uniodonto=formatar_moeda(linha['uniodonto']),
                refeicao=formatar_moeda(linha['refeicao']),
                farmacia=formatar_moeda(linha['farmacia']),
                total=formatar_moeda(linha['total']),
                dia=datetime.now().strftime("%d"),
                mes=datetime.now().strftime("%B").capitalize(),
                ano=datetime.now().strftime("%Y"),
                caminho_logo=caminho_logo_base64
            )

            # Gerar o PDF com o CSS incorporado
            nome_pdf = os.path.join(pasta_saida, f"{linha['Nome do Funcionário'].replace(' ', '_')}.pdf")
            HTML(string=html_preenchido).write_pdf(nome_pdf, stylesheets=[caminho_css], dpi=96)

        messagebox.showinfo("Sucesso", f"Recibos gerados na pasta:\n{pasta_saida}")

    except Exception as e:
        print(f"Erro inesperado: {str(e)}")

def criar_interface():
    print("Iniciando a interface gráfica...")
    root = Tk()
    root.title("Gerador de Recibos")
    root.geometry("300x150")

    Label(root, text="Gerador de Recibos", font=("Arial", 14)).pack(pady=10)
    Button(root, text="Selecionar Excel", command=lambda: gerar_recibos(selecionar_arquivo())).pack(pady=5)
    Button(root, text="Abrir Pasta Documentos", command=abrir_pasta_documentos).pack(pady=5)
    Button(root, text="Fechar", command=root.quit).pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    criar_interface()
