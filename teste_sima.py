import requests
from bs4 import BeautifulSoup


URL_SIMA = "https://celepar7.pr.gov.br/sima/cotdiap.asp"


def baixar_pagina():
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    response = requests.get(URL_SIMA, headers=headers, timeout=30)

    # O site pode usar encoding antigo
    response.encoding = response.apparent_encoding

    response.raise_for_status()

    return response.text


def listar_produtos():
    html = baixar_pagina()

    soup = BeautifulSoup(html, "html.parser")

    selects = soup.find_all("select")

    if not selects:
        print("Nenhum campo SELECT encontrado na página.")
        print("Trecho inicial do HTML:")
        print(html[:3000])
        return

    print(f"Quantidade de SELECT encontrados: {len(selects)}")
    print("-" * 80)

    for i, select in enumerate(selects, start=1):
        nome_select = select.get("name")
        id_select = select.get("id")

        print(f"SELECT {i}")
        print(f"name: {nome_select}")
        print(f"id: {id_select}")
        print("-" * 80)

        options = select.find_all("option")

        for option in options:
            texto = option.get_text(strip=True)
            valor = option.get("value")

            if not texto:
                continue

            # Mostra só os produtos que interessam
            texto_minusculo = texto.lower()

            if "soja" in texto_minusculo or "milho" in texto_minusculo:
                print(f"Produto: {texto}")
                print(f"Valor interno: {valor}")
                print("-" * 40)


def main():
    listar_produtos()


if __name__ == "__main__":
    main()
