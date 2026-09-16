"""
automatiza_caixa.py
--------------------
Copia o valor "DINHEIRO" da tela "Conferência de Caixa" do Gcom
para a planilha "FINANCEIRO REI DO MATE" (Google Sheets), como uma
nova linha de ENTRADA / CAIXA DO DIA.

COMO FUNCIONA
1. Selenium se conecta a uma janela do Chrome que VOCÊ já abriu e já
   está logado (não guarda nem pede sua senha).
2. Procura a linha "DINHEIRO" na tabela de conferência de caixa e lê
   o valor do campo ao lado.
3. Usa a API do Google Sheets (gspread) para escrever uma nova linha
   na aba de movimentação de caixa.

ANTES DE RODAR, siga a seção "CONFIGURAÇÃO" logo abaixo.
"""

import re
import sys
import datetime
import subprocess
import time

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import gspread
from google.oauth2.service_account import Credentials


# =========================================================================
# CONFIGURAÇÃO — AJUSTE ESTES VALORES ANTES DE USAR
# =========================================================================

# 1) Porta de depuração remota do Chrome (ver instruções de uso mais abaixo)
CHROME_DEBUG_PORT = 9222

# 2) Trecho da URL da página do Gcom que tem a Conferência de Caixa
GCOM_URL_TRECHO = "analise-fechamento-caixa"

# 3) Planilha do Google Sheets
GOOGLE_SHEET_ID = "1gs1knbb8zaVgxpXSw3gvSTmwfNMQF-oY72G1zdJ9H24"   # tirado da URL da planilha
NOME_DA_ABA = "MOVIMENTAÇÃO DE CAIXA 26"
CAMINHO_CREDENCIAIS = "credenciais_google.json"          # gerado no passo a passo abaixo

# 4) Quem/o que registrar na linha nova
RESPONSAVEL_PADRAO = "AUTOMACAO"
TIPO_PADRAO = "ENTRADA"
MOTIVO_PADRAO = "CAIXA DO DIA"

# 5) Colunas da planilha (A=1, B=2, C=3, D=4, E=5, F=6)
COL_DATA = 1
COL_RESPONSAVEL = 2
COL_TIPO = 3
COL_MOTIVO = 4
COL_VALOR = 5

# =========================================================================


def conectar_chrome_existente():
    """Anexa o Selenium a um Chrome que já está aberto e logado."""
    options = Options()
    options.debugger_address = f"127.0.0.1:{CHROME_DEBUG_PORT}"
    driver = webdriver.Chrome(options=options)
    return driver


def achar_aba_do_gcom(driver):
    """Troca para a aba do Chrome que tem a página do Gcom aberta."""
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if GCOM_URL_TRECHO in driver.current_url:
            return True
    return False


def _rect(driver, elemento):
    """Pega a posição/tamanho do elemento na tela (coordenadas de tela)."""
    return driver.execute_script(
        "const r = arguments[0].getBoundingClientRect();"
        "return {top: r.top, bottom: r.bottom, left: r.left, right: r.right};",
        elemento,
    )


def ler_valor_dinheiro(driver):
    """
    Procura o texto 'DINHEIRO' na tela e lê o valor do campo que está
    na MESMA ALTURA (mesma linha visual), à direita dele.

    Essa abordagem por posição na tela é usada porque a tela de
    Conferência de Caixa é um grid (estilo planilha), onde cada célula
    fica em um bloco separado, não em uma <tr> tradicional — então não
    dá pra confiar em "subir para o elemento pai".
    """
    wait = WebDriverWait(driver, 15)

    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[normalize-space(text())='DINHEIRO']")
        )
    )

    candidatos = driver.find_elements(
        By.XPATH, "//*[normalize-space(text())='DINHEIRO']"
    )
    if not candidatos:
        raise Exception("Não encontrei o texto 'DINHEIRO' na tela.")

    # Entre os candidatos, pega o mais "de baixo" da árvore (o mais
    # específico), que deve ser o próprio rótulo, não um contêiner grande.
    rotulo = min(candidatos, key=lambda el: len(el.find_elements(By.XPATH, ".//*")))
    rotulo_rect = _rect(driver, rotulo)
    centro_y_rotulo = (rotulo_rect["top"] + rotulo_rect["bottom"]) / 2

    inputs = driver.find_elements(By.TAG_NAME, "input")

    melhor_input = None
    menor_diff = None
    for campo in inputs:
        try:
            rect = _rect(driver, campo)
        except Exception:
            continue
        centro_y = (rect["top"] + rect["bottom"]) / 2
        diff = abs(centro_y - centro_y_rotulo)

        # precisa estar na mesma altura (tolerância de 12px) e à direita do rótulo
        if diff < 12 and rect["left"] > rotulo_rect["left"]:
            if menor_diff is None or diff < menor_diff:
                menor_diff = diff
                melhor_input = campo

    if melhor_input is None:
        raise Exception(
            "Achei o texto 'DINHEIRO' mas não achei o campo de valor ao lado. "
            "Confira se a linha está visível na tela (sem precisar rolar) "
            "e rode de novo."
        )

    valor_texto = melhor_input.get_attribute("value")
    return extrair_valor(valor_texto)


def extrair_valor(texto):
    """Converte '66,42' -> 66.42 (pega o último número no formato BR)."""
    numeros = re.findall(r"[\d\.]+,\d{2}", texto)
    if not numeros:
        raise ValueError(f"Não consegui achar um valor em: {texto!r}")
    valor_str = numeros[-1].replace(".", "").replace(",", ".")
    return float(valor_str)


def escrever_na_planilha(valor):
    """Adiciona uma nova linha na próxima linha vazia da aba escolhida."""
    escopos = [
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds = Credentials.from_service_account_file(CAMINHO_CREDENCIAIS, scopes=escopos)
    cliente = gspread.authorize(creds)

    planilha = cliente.open_by_key(GOOGLE_SHEET_ID)
    aba = planilha.worksheet(NOME_DA_ABA)

    # Acha a primeira linha vazia olhando a coluna DATA
    valores_coluna_data = aba.col_values(COL_DATA)
    proxima_linha = len(valores_coluna_data) + 1

    hoje = datetime.date.today().strftime("%d/%m/%Y")

    # Valor formatado como a planilha espera (célula de moeda "R$ ...")
    linha_nova = [None] * COL_VALOR
    linha_nova[COL_DATA - 1] = hoje
    linha_nova[COL_RESPONSAVEL - 1] = RESPONSAVEL_PADRAO
    linha_nova[COL_TIPO - 1] = TIPO_PADRAO
    linha_nova[COL_MOTIVO - 1] = MOTIVO_PADRAO
    linha_nova[COL_VALOR - 1] = valor

    aba.update(f"A{proxima_linha}:E{proxima_linha}", [linha_nova])
    print(f"Linha {proxima_linha} escrita com sucesso: {linha_nova}")


def main():
    print("Conectando ao Chrome já aberto...")
    driver = conectar_chrome_existente()

    if not achar_aba_do_gcom(driver):
        print("Não encontrei uma aba do Chrome aberta na página do Gcom "
              "(Conferência de Caixa). Abra a página e rode de novo.")
        sys.exit(1)

    print("Lendo valor DINHEIRO na tela...")
    valor = ler_valor_dinheiro(driver)
    print(f"Valor encontrado: R$ {valor:.2f}")

    print("Gravando na planilha...")
    escrever_na_planilha(valor)

    print("Concluído.")


if __name__ == "__main__":
    main()
