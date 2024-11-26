import pandas as pd
from jinja2 import Template
from weasyprint import HTML
from utils import obter_caminho_relativo, formatar_moeda, imagem_para_base64
import os
from datetime import datetime
from tkinter import messagebox

def gerar_recibos(caminho_planilha, progress_bar=None):
    """Gera os recibos e atualiza a barra de progresso, se fornecida."""
    try:
        # Caminho para os arquivos
        caminho_html = obter_caminho_relativo('templates/recibo.html')
        caminho_css = obter_caminho_relativo('templates/bootstrap.min.css')
        caminho_logo = obter_caminho_relativo('assets/logo.png')

        # Verificar se os arquivos necessários existem
        for caminho, nome in [(caminho_html, "recibo.html"), (caminho_css, "bootstrap.min.css"), (caminho_logo, "logo.png")]:
            if not os.path.exists(caminho):
                messagebox.showerror("Erro", f"O arquivo '{nome}' não foi encontrado.")
                return

        # Ler o HTML e substituir o caminho do CSS para absoluto
        with open(caminho_html, 'r', encoding='utf-8') as arquivo_html:
            template_html = arquivo_html.read()
        
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

        # Criar a pasta de saída
        data_atual = datetime.now().strftime("%d.%m.%Y")
        pasta_saida = os.path.join(os.path.expanduser("~/Documents"), f"{data_atual}-recibos")
        os.makedirs(pasta_saida, exist_ok=True)

        # Total de registros para a barra de progresso
        total = len(dados)

        # Gerar os PDFs
        for idx, (_, linha) in enumerate(dados.iterrows(), start=1):
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
            nome_pdf = os.path.join(pasta_saida, f"{linha['Nome do Funcionário'].replace(' ', '_')}.pdf")
            HTML(string=html_preenchido).write_pdf(nome_pdf, stylesheets=[caminho_css], dpi=96)

            # Atualiza a barra de progresso, se fornecida
            if progress_bar:
                progress_bar["value"] = (idx / total) * 100
                progress_bar.update_idletasks()

        messagebox.showinfo("Sucesso", f"Recibos gerados na pasta:\n{pasta_saida}")

    except Exception as e:
        messagebox.showerror("Erro", f"Erro inesperado: {str(e)}")

