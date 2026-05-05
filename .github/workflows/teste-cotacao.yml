name: Teste Cotacao Milho B3

on:
  workflow_dispatch:
  schedule:
    - cron: "0 12 * * 1-5" # 09:00 no Brasil
    - cron: "0 15 * * 1-5" # 12:00 no Brasil
    - cron: "0 18 * * 1-5" # 15:00 no Brasil
    - cron: "0 21 * * 1-5" # 18:00 no Brasil

permissions:
  contents: write

jobs:
  testar-cotacao:
    runs-on: ubuntu-latest

    steps:
      - name: Baixar arquivos do repositorio
        uses: actions/checkout@v4

      - name: Configurar Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Instalar dependencias
        run: pip install -r requirements.txt

      - name: Rodar coleta e enviar para SharePoint
        env:
          TENANT_ID: ${{ secrets.TENANT_ID }}
          CLIENT_ID: ${{ secrets.CLIENT_ID }}
          CLIENT_SECRET: ${{ secrets.CLIENT_SECRET }}
          SHAREPOINT_HOSTNAME: ${{ secrets.SHAREPOINT_HOSTNAME }}
          SHAREPOINT_SITE_PATH: ${{ secrets.SHAREPOINT_SITE_PATH }}
          SHAREPOINT_FOLDER_PATH: ${{ secrets.SHAREPOINT_FOLDER_PATH }}
        run: python main.py

      - name: Salvar Excel atualizado no repositorio
        run: |
          git config user.name "github-actions"
          git config user.email "github-actions@github.com"
          git add cotacoes_graos.xlsx
          git commit -m "Atualiza cotacoes graos" || echo "Sem alteracoes para salvar"
          git push
