import re
import requests
from bs4 import BeautifulSoup


URL_SIMA = "https://celepar7.pr.gov.br/sima/cotdiap.asp"

PRODUTOS = [
    {
        "codigo": "7",
        "grao": "MILHO",
        "produto": "Milho amarelo tipo 1 sc 60 Kg",
    },
    {
        "codigo": "8",
        "grao": "SOJA",
        "produto": "Soja industrial tipo 1 sc 60 Kg",
    },
]


def numero_br_para_float(valor):
    if valor is None:
        return None

    valor = str(valor).strip()

    if valor.upper() in ["SINF", "AUS", "-", ""]:
        return None

    valor = valor.replace(".", "")
    valor = valor.replace(",", ".")

    return float(valor)


def baixar_pagina_produto(codigo_produto):
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    params = {
        "produto": codigo_produto
    }

    response = requests.get(
        URL_SIMA,
        params=params,
        headers=headers,
        timeout=30
    )

    response.encoding = response.apparent_encoding
    response.raise_for_status()

    return response.text


def extrair_data_cotacao(texto):
    match = re.search(r"em\s+(\d{2}/\d{2}/\d{4})", texto)

    if match:
        return match.group(1)

    return None


def extrair_tabela_m_c(html, grao, produto_nome):
    soup = BeautifulSoup(html, "html.parser")
    texto = soup.get_text("\n", strip=True)

    data_cotacao = extrair_data_cotacao(texto)

    tabelas = soup.find_all("table")

    if not tabelas:
        print(f"Nenhuma tabela encontrada para {produto_nome}.")
        print(texto[:2000])
        return []

    linhas_extraidas = []

    for tabela in tabelas:
        linhas = tabela.find_all("tr")

        for linha in linhas:
            colunas = linha.find_all(["td", "th"])
            valores = [col.get_text(strip=True) for col in colunas]

            if len(valores) < 4:
                continue

            primeira_coluna = valores[0].strip()

            if primeira_coluna.upper() in ["NÚCLEO REGIONAL", "NUCLEO REGIONAL"]:
                continue

            min_valor = valores[1].strip()
            m_c_valor = valores[2].strip()
            max_valor = valores[3].strip()

            # Evita pegar linhas de legenda/rodapé
            if primeira_coluna.upper() in ["MIN", "M_C", "MAX"]:
                continue

            m_c_numero = numero_br_para_float(m_c_valor)

            linha_saida = {
                "data_cotacao": data_cotacao,
                "grao": grao,
                "produto": produto_nome,
                "nucleo_regional": primeira_coluna,
                "m_c": m_c_numero,
                "m_c_original": m_c_valor,
            }

            linhas_extraidas.append(linha_saida)

    return linhas_extraidas


def main():
    todas_linhas = []

    for produto in PRODUTOS:
        print("=" * 80)
        print(f"Coletando: {produto['produto']} | Código: {produto['codigo']}")

        html = baixar_pagina_produto(produto["codigo"])

        linhas = extrair_tabela_m_c(
            html=html,
            grao=produto["grao"],
            produto_nome=produto["produto"],
        )

        print(f"Linhas encontradas: {len(linhas)}")

        for linha in linhas:
            print(
                f"{linha['data_cotacao']} | "
                f"{linha['grao']} | "
                f"{linha['nucleo_regional']} | "
                f"M_c: {linha['m_c_original']}"
            )

        todas_linhas.extend(linhas)

    print("=" * 80)
    print(f"TOTAL GERAL DE LINHAS: {len(todas_linhas)}")


if __name__ == "__main__":
    main()
