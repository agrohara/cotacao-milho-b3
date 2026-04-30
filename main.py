import os
import re
from datetime import datetime
from urllib.parse import quote

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


URL_MILHO = "https://www.noticiasagricolas.com.br/cotacoes/milho/milho-b3-prego-regular"
URL_SOJA = "https://www.noticiasagricolas.com.br/cotacoes/soja/soja-bolsa-de-chicago-cme-group"
URL_DOLAR = "https://www.noticiasagricolas.com.br/cotacao-do-dolar/"

ARQUIVO_EXCEL = "cotacoes_graos.xlsx"
NOME_TABELA_EXCEL = "tb_cotacoes_graos"


def numero_br_para_float(valor):
    if valor is None:
        return None

    valor = str(valor).strip()
    valor = valor.replace("+", "")
    valor = valor.replace(".", "")
    valor = valor.replace(",", ".")

    return float(valor)


def baixar_texto_pagina(url):
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    print(f"Acessando: {url}")

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")
    texto = soup.get_text("\n", strip=True)

    return texto


def buscar_dolar():
    texto = baixar_texto_pagina(URL_DOLAR)

    match = re.search(r"R\$\s*(\d+[,.]\d{2,4})", texto)

    if not match:
        print("Não consegui encontrar o dólar. Texto parcial da página:")
        print(texto[:1500])
        raise Exception("Dólar não encontrado.")

    dolar = numero_br_para_float(match.group(1))
    dolar = round(dolar, 2)

    print(f"Dólar encontrado: {dolar:.2f}")

    return dolar


def extrair_bloco_mais_recente(texto):
    match_data = re.search(r"Fechamento:\s*(\d{2}/\d{2}/\d{4})", texto)

    if not match_data:
        print("Texto parcial da página:")
        print(texto[:2000])
        raise Exception("Data de fechamento não encontrada.")

    data_cotacao = match_data.group(1)

    inicio = match_data.end()

    proximo_fechamento = re.search(r"Fechamento:\s*\d{2}/\d{2}/\d{4}", texto[inicio:])

    if proximo_fechamento:
        fim = inicio + proximo_fechamento.start()
        bloco = texto[inicio:fim]
    else:
        bloco = texto[inicio:]

    return data_cotacao, bloco


def buscar_cotacoes_milho(dolar):
    texto = baixar_texto_pagina(URL_MILHO)

    data_cotacao, bloco = extrair_bloco_mais_recente(texto)

    linhas = []

    padrao = re.compile(
        r"([A-Za-zçÇãÃéÉíÍóÓúÚ]+/\d{4})\s+([\d,]+)\s+([-+]?[\d,]+)"
    )

    for match in padrao.finditer(bloco):
        contrato_mes = match.group(1)
        fechamento = numero_br_para_float(match.group(2))

        linha = {
            "data_coleta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data_cotacao": data_cotacao,
            "grao": "MILHO",
            "contrato_mes": contrato_mes,
            "fechamento_rs_sc_60kg": fechamento,
            "fechamento_usd_bushel": None,
            "dolar": dolar,
            "fonte": "Notícias Agrícolas / B3",
            "url": URL_MILHO,
            "chave": f"MILHO_{contrato_mes}_{data_cotacao}",
        }

        linhas.append(linha)

    if not linhas:
        print("Bloco de milho:")
        print(bloco[:2000])
        raise Exception("Nenhuma cotação de milho encontrada.")

    print(f"Cotações de milho encontradas: {len(linhas)}")

    return linhas


def buscar_cotacoes_soja(dolar):
    texto = baixar_texto_pagina(URL_SOJA)

    data_cotacao, bloco = extrair_bloco_mais_recente(texto)

    linhas = []

    padrao = re.compile(
        r"([A-Za-zçÇãÃéÉíÍóÓúÚ]+/\d{2})\s+([\d,]+)\s+([-+]?[\d,]+)\s+([-+]?[\d,]+)"
    )

    for match in padrao.finditer(bloco):
        contrato_mes = match.group(1)
        fechamento_usd_bushel = numero_br_para_float(match.group(2))

        linha = {
            "data_coleta": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "data_cotacao": data_cotacao,
            "grao": "SOJA",
            "contrato_mes": contrato_mes,
            "fechamento_rs_sc_60kg": None,
            "fechamento_usd_bushel": fechamento_usd_bushel,
            "dolar": dolar,
            "fonte": "Notícias Agrícolas / CME Group",
            "url": URL_SOJA,
            "chave": f"SOJA_{contrato_mes}_{data_cotacao}",
        }

        linhas.append(linha)

    if not linhas:
        print("Bloco de soja:")
        print(bloco[:2000])
        raise Exception("Nenhuma cotação de soja encontrada.")

    print(f"Cotações de soja encontradas: {len(linhas)}")

    return linhas


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
            "data_cotacao",
            "grao",
            "contrato_mes",
            "fechamento_rs_sc_60kg",
            "fechamento_usd_bushel",
            "dolar",
            "fonte",
            "url",
            "chave",
        ]

        sheet.append(cabecalhos)

    return workbook, sheet


def carregar_chaves_existentes(sheet):
    chaves = set()

    for row in sheet.iter_rows(min_row=2, values_only=True):
        chave = row[9]

        if chave:
            chaves.add(chave)

    return chaves


def formatar_colunas_excel(sheet):
    for cell in sheet["A"]:
        if cell.row == 1:
            continue
        cell.number_format = "yyyy-mm-dd hh:mm:ss"

    for cell in sheet["E"]:
        if cell.row == 1:
            continue
        cell.number_format = "0.00"

    for cell in sheet["F"]:
        if cell.row == 1:
            continue
        cell.number_format = "0.0000"

    for cell in sheet["G"]:
        if cell.row == 1:
            continue
        cell.number_format = "0.00"

    larguras = {
        "A": 22,
        "B": 15,
        "C": 12,
        "D": 18,
        "E": 24,
        "F": 24,
        "G": 12,
        "H": 32,
        "I": 80,
        "J": 35,
    }

    for coluna, largura in larguras.items():
        sheet.column_dimensions[coluna].width = largura


def criar_ou_atualizar_tabela_excel(sheet):
    ultima_linha = sheet.max_row
    ultima_coluna = sheet.max_column

    if ultima_linha < 2:
        print("Não há linhas suficientes para criar tabela.")
        return

    intervalo_tabela = f"A1:J{ultima_linha}"

    if NOME_TABELA_EXCEL in sheet.tables:
        print(f"Tabela {NOME_TABELA_EXCEL} já existe. Atualizando intervalo...")
        sheet.tables[NOME_TABELA_EXCEL].ref = intervalo_tabela
    else:
        print(f"Criando tabela {NOME_TABELA_EXCEL} no intervalo {intervalo_tabela}...")

        tabela = Table(
            displayName=NOME_TABELA_EXCEL,
            ref=intervalo_tabela
        )

        estilo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )

        tabela.tableStyleInfo = estilo

        sheet.add_table(tabela)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = intervalo_tabela


def salvar_no_excel(linhas):
    workbook, sheet = criar_ou_abrir_excel()
    chaves_existentes = carregar_chaves_existentes(sheet)

    novas_linhas = 0
    duplicadas = 0

    for cotacao in linhas:
        if cotacao["chave"] in chaves_existentes:
            print(f"Já existe, não vou duplicar: {cotacao['chave']}")
            duplicadas += 1
            continue

        nova_linha = [
            cotacao["data_coleta"],
            cotacao["data_cotacao"],
            cotacao["grao"],
            cotacao["contrato_mes"],
            cotacao["fechamento_rs_sc_60kg"],
            cotacao["fechamento_usd_bushel"],
            cotacao["dolar"],
            cotacao["fonte"],
            cotacao["url"],
            cotacao["chave"],
        ]

        sheet.append(nova_linha)
        chaves_existentes.add(cotacao["chave"])
        novas_linhas += 1

    formatar_colunas_excel(sheet)
    criar_ou_atualizar_tabela_excel(sheet)

    workbook.save(ARQUIVO_EXCEL)

    print(f"Arquivo salvo: {ARQUIVO_EXCEL}")
    print(f"Novas linhas adicionadas: {novas_linhas}")
    print(f"Linhas duplicadas ignoradas: {duplicadas}")


def obter_token_graph():
    tenant_id = os.environ["TENANT_ID"]
    client_id = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]

    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    payload = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials",
    }

    response = requests.post(url, data=payload, timeout=30)
    response.raise_for_status()

    return response.json()["access_token"]


def obter_site_id(token):
    hostname = os.environ["SHAREPOINT_HOSTNAME"]
    site_path = os.environ["SHAREPOINT_SITE_PATH"]

    headers = {
        "Authorization": f"Bearer {token}"
    }

    if site_path == "/":
        url = f"https://graph.microsoft.com/v1.0/sites/{hostname}"
    else:
        url = f"https://graph.microsoft.com/v1.0/sites/{hostname}:{site_path}"

    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()

    site_id = response.json()["id"]

    print(f"Site ID encontrado: {site_id}")

    return site_id


def enviar_excel_para_sharepoint():
    print("Enviando Excel para o SharePoint...")

    token = obter_token_graph()
    site_id = obter_site_id(token)

    folder_path = os.environ["SHAREPOINT_FOLDER_PATH"].strip("/")
    arquivo_destino = ARQUIVO_EXCEL

    encoded_path = quote(f"{folder_path}/{arquivo_destino}")

    url_upload = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{encoded_path}:/content"
    )

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }

    with open(ARQUIVO_EXCEL, "rb") as arquivo:
        response = requests.put(url_upload, headers=headers, data=arquivo, timeout=60)

    response.raise_for_status()

    print("Arquivo enviado para o SharePoint com sucesso.")
    print(f"Destino: {folder_path}/{arquivo_destino}")


def main():
    dolar = buscar_dolar()

    linhas_milho = buscar_cotacoes_milho(dolar)
    linhas_soja = buscar_cotacoes_soja(dolar)

    todas_linhas = linhas_milho + linhas_soja

    print(f"Total de linhas capturadas: {len(todas_linhas)}")

    salvar_no_excel(todas_linhas)

    enviar_excel_para_sharepoint()


if __name__ == "__main__":
    main()
