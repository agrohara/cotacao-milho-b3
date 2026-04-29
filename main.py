import requests
from bs4 import BeautifulSoup
import re

URL = "https://www.noticiasagricolas.com.br/cotacoes/milho/milho-b3-prego-regular"
CONTRATO_ALVO = "Setembro/2026"

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

print("Cotação encontrada:")
print(f"Data fechamento: {data_fechamento}")
print(f"Contrato: {CONTRATO_ALVO}")
print(f"Fechamento R$/sc 60kg: {fechamento_num}")
print(f"Variação %: {variacao_num}")
print(f"Fonte: Notícias Agrícolas / B3")
print(f"URL: {URL}")
