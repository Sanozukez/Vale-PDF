import os
import sys
import base64
import locale
import pandas as pd

# Configurar o locale
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil.1252')

def formatar_moeda(valor):
    """Formata um número como moeda brasileira ou retorna '-' para valores ausentes."""
    if pd.isna(valor) or valor == 0:
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
