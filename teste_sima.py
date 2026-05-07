import re
import requests
from bs4 import BeautifulSoup


URL_INICIAL = "https://celepar7.pr.gov.br/sima/cotdiap.asp"
URL_RESULTADO = "https://celepar7.pr.gov.br/sima/cotdiap1.asp"

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

    # Só tenta converter se tiver aparência de número
    # Exemplos aceitos: 52,00 / 52.00 / 109 / 109,50
    if not re.search(r"\d", valor):
        return None

    if any(palavra in valor.upper() for palavra in ["PREÇO", "PRECO", "COMUM", "M_C", "MIN", "MAX"]):
        return None

    valor = valor.replace(".", "")
    valor = valor.replace(",", ".")

    try:
        return float(valor)
    except ValueError:
        return None


def baixar_resultado_produto(codigo_produto):
    session = requests.Session()

    headers_get = {
        "User-Agent": "Mozilla/5.0"
    }

    response_inicial = session.get(
        URL_INICIAL,
        headers=headers_get,
        timeout=30
    )

    response_inicial.encoding = response_inicial.apparent_encoding
    response_inicial.raise_for_status()

    headers_post = {
        "User-Agent": "Mozilla/5.0",
        "Referer": URL_INICIAL,
        "Content-Type": "application/x-www-form-urlencoded",
    }

    dados = {
        "produto": codigo_produto,
        "submit1": "Pesquisar",
    }

    response = session.post(
        URL_RESULTADO,
        data=dados,
        headers=headers_post,
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


def extrair_linhas_m_c(html, grao, produto_nome):
    soup = BeautifulSoup(html, "html.parser")
    texto = soup.get_text("\n", strip=True)

    data_cotacao = extrair_data_cotacao(texto)

    print(f"Data cotação encontrada: {data_cotacao}")

    tabelas = soup.find_all("table")

    if not tabelas:
        print(f"Nenhuma tabela encontrada para {produto_nome}.")
        print("Texto retornado:")
        print(texto[:2500])
        return []

    linhas_extraidas = []

    for tabela in tabelas:
        linhas = tabela.find_all("tr")

        for linha in linhas:
            colunas = linha.find_all(["td", "th"])
            valores = [col.get_text(strip=True) for col in colunas]

            if len(valores) < 4:
                continue

            nucleo_regional = valores[0].strip()
            minimo = valores[1].strip()
            m_c = valores[2].strip()
            maximo = valores[3].strip()

            texto_linha = " ".join(valores).upper()

            # Ignora cabeçalho
            if "NÚCLEO REGIONAL" in texto_linha or "NUCLEO REGIONAL" in texto_linha:
                continue

            # Ignora legenda/rodapé
            if "PREÇO" in texto_linha or "PRECO" in texto_linha:
                continue

            if "FONTE" in texto_linha:
                continue

            if "M_C" in texto_linha and "PREÇO" in texto_linha:
                continue

            m_c_numero = numero_br_para_float(m_c)

            # Se M_c não virou número, ignora a linha
            if m_c_numero is None:
                continue

            linha_saida = {
                "data_cotacao": data_cotacao,
                "grao": grao,
                "produto": produto_nome,
                "nucleo_regional": nucleo_regional,
                "minimo": numero_br_para_float(minimo),
                "m_c": m_c_numero,
                "maximo": numero_br_para_float(maximo),
                "m_c_original": m_c,
            }

            linhas_extraidas.append(linha_saida)

    return linhas_extraidas


def main():
    todas_linhas = []

    for produto in PRODUTOS:
        print("=" * 80)
        print(f"Coletando: {produto['produto']}")
        print(f"Código interno: {produto['codigo']}")

        html = baixar_resultado_produto(produto["codigo"])

        linhas = extrair_linhas_m_c(
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
