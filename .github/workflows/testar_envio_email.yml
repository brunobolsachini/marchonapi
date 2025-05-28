name: Testar Envio E-mail

on:
  workflow_dispatch:

jobs:
  test-email:
    runs-on: ubuntu-latest

    env:
      SENHA_API: ${{ secrets.SENHA_API }}

    steps:
      - name: Checkout do repositório
        uses: actions/checkout@v2

      - name: Configurar Python
        uses: actions/setup-python@v2
        with:
          python-version: '3.x'

      - name: Testar envio de e-mail
        run: python .github/scripts/testar_envio_email.py
