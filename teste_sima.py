import requests
from bs4 import BeautifulSoup


URL_SIMA = "https://celepar7.pr.gov.br/sima/cotdiap.asp"


def baixar_pagina_inicial():
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    response = requests.get(URL_SIMA, headers=headers, timeout=30)
    response.encoding = response.apparent_encoding
    response.raise_for_status()

    return response.text


def analisar_formularios():
    html = baixar_pagina_inicial()

    soup = BeautifulSoup(html, "html.parser")

    forms = soup.find_all("form")

    print(f"Quantidade de FORM encontrados: {len(forms)}")
    print("=" * 80)

    if not forms:
        print("Nenhum FORM encontrado.")
        print("Trecho inicial do HTML:")
        print(html[:4000])
        return

    for i, form in enumerate(forms, start=1):
        print(f"FORM {i}")
        print(f"method: {form.get('method')}")
        print(f"action: {form.get('action')}")
        print("-" * 80)

        inputs = form.find_all(["input", "select", "textarea"])

        for campo in inputs:
            tag = campo.name
            nome = campo.get("name")
            tipo = campo.get("type")
            valor = campo.get("value")

            print(f"tag: {tag} | name: {nome} | type: {tipo} | value: {valor}")

            if tag == "select":
                options = campo.find_all("option")

                print("Opções do SELECT:")

                for option in options:
                    texto = option.get_text(strip=True)
                    valor_option = option.get("value")

                    if "soja" in texto.lower() or "milho" in texto.lower():
                        print(f"  Produto: {texto} | value: {valor_option}")

        print("=" * 80)


def testar_envios():
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Content-Type": "application/x-www-form-urlencoded",
    }

    testes = [
        {
            "descricao": "GET com produto=7",
            "metodo": "GET",
            "dados": {"produto": "7"},
        },
        {
            "descricao": "GET com produto=8",
            "metodo": "GET",
            "dados": {"produto": "8"},
        },
        {
            "descricao": "POST com produto=7",
            "metodo": "POST",
            "dados": {"produto": "7"},
        },
        {
            "descricao": "POST com produto=8",
            "metodo": "POST",
            "dados": {"produto": "8"},
        },
    ]

    print("\n")
    print("#" * 80)
    print("TESTANDO FORMAS DE ENVIO")
    print("#" * 80)

    for teste in testes:
        descricao = teste["descricao"]
        metodo = teste["metodo"]
        dados = teste["dados"]

        print("\n")
        print(f"Teste: {descricao}")

        if metodo == "GET":
            response = requests.get(
                URL_SIMA,
                params=dados,
                headers={"User-Agent": "Mozilla/5.0"},
                timeout=30,
            )
        else:
            response = requests.post(
                URL_SIMA,
                data=dados,
                headers=headers,
                timeout=30,
            )

        response.encoding = response.apparent_encoding
        response.raise_for_status()

        texto = BeautifulSoup(response.text, "html.parser").get_text("\n", strip=True)

        encontrou_tabela = "NÚCLEO REGIONAL" in texto or "NUCLEO REGIONAL" in texto
        encontrou_soja = "Soja industrial tipo 1 sc 60 Kg" in texto
        encontrou_milho = "Milho amarelo tipo 1 sc 60 Kg" in texto
        encontrou_media = "Média do Dia" in texto or "Media do Dia" in texto

        print(f"Encontrou tabela NÚCLEO REGIONAL? {encontrou_tabela}")
        print(f"Encontrou Média do Dia? {encontrou_media}")
        print(f"Encontrou texto soja? {encontrou_soja}")
        print(f"Encontrou texto milho? {encontrou_milho}")

        print("Trecho do texto retornado:")
        print("-" * 80)
        print(texto[:2500])
        print("-" * 80)


def main():
    analisar_formularios()
    testar_envios()


if __name__ == "__main__":
    main()
