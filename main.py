import os
import re
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook


URL = "https://www.noticiasagricolas.com.br/cotacoes/milho/milho-b3-prego-regular"
CONTRATO_ALVO = "Setembro/2026"
ARQUIVO_EXCEL = "cotacoes_milho_b3.xlsx"


def buscar_cotacao():
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    print("Acessando página...")
    response = requests.get(URL, headers=headers, timeout=30)
    response.raise_for_status()

    print("Página acessada com sucesso.")

    soup = BeautifulSoup(response.text, "html.parser")
    texto = soup.get_text("\n", strip=True)

    fechamento_match = re.search(r"Fechamento:\s*(\d{2}/\d{2}/\d{4})", texto)

    if not fechamento_match:
        print("Texto parcial da página:")
        print(texto[:2000])
        raise Exception("Não encontrei a data de fechamento.")

    data_fechamento = fechamento_match.group(1)

    padrao = rf"{CONTRATO_ALVO}\s+([\d,]+)\s+(-?[\d,]+)"
    contrato_match = re.search(padrao, texto)

    if not contrato_match:
        print("Texto parcial da página:")
        print(texto[:3000])
        raise Exception(f"Não encontrei o contrato {CONTRATO_ALVO}.")

    fechamento = contrato_match.group(1)
    variacao = contrato_match.group(2)

    fechamento_num = float(fechamento.replace(",", "."))
    variacao_num = float(variacao.replace(",", "."))

    resultado = {
        "data_coleta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "data_fechamento": data_fechamento,
        "contrato": CONTRATO_ALVO,
        "fechamento_rs_sc_60kg": fechamento_num,
        "variacao_percentual": variacao_num,
        "fonte": "Notícias Agrícolas / B3",
        "url": URL,
        "chave": f"{CONTRATO_ALVO}_{data_fechamento}",
    }

    print("Cotação encontrada:")
    print(resultado)

    return resultado


def criar_ou_abrir_excel():
    if os.path.exists(ARQUIVO_EXCEL):
        print("Arquivo Excel já existe. Abrindo...")
        workbook = load_workbook(ARQUIVO_EXCEL)
        sheet = workbook.active
    else:
        print("Arquivo Excel não existe. Criando...")
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "cotacoes"

        cabecalhos = [
            "data_coleta",
            "data_fechamento",
            "contrato",
            "fechamento_rs_sc_60kg",
            "variacao_percentual",
            "fonte",
            "url",
            "chave",
        ]

        sheet.append(cabecalhos)

    return workbook, sheet


def chave_ja_existe(sheet, chave):
    # A coluna H é a coluna "chave"
    for row in sheet.iter_rows(min_row=2, values_only=True):
        chave_existente = row[7]

        if chave_existente == chave:
            return True

    return False


def salvar_no_excel(cotacao):
    workbook, sheet = criar_ou_abrir_excel()

    if chave_ja_existe(sheet, cotacao["chave"]):
        print(f"A cotação {cotacao['chave']} já existe no Excel. Não vou duplicar.")
        return

    nova_linha = [
        cotacao["data_coleta"],
        cotacao["data_fechamento"],
        cotacao["contrato"],
        cotacao["fechamento_rs_sc_60kg"],
        cotacao["variacao_percentual"],
        cotacao["fonte"],
        cotacao["url"],
        cotacao["chave"],
    ]

    sheet.append(nova_linha)

    workbook.save(ARQUIVO_EXCEL)

    print(f"Cotação adicionada no arquivo {ARQUIVO_EXCEL}.")


def main():
    cotacao = buscar_cotacao()
    salvar_no_excel(cotacao)


if __name__ == "__main__":
    main()
