"""
Sincroniza a TabelaLancamentos (planilha "Gestão Financeira Pessoal.xlsx",
SharePoint) para um arquivo estático "lancamentos.json" no MESMO repositório
do GitHub Pages do CAP.

Motivo: contas Microsoft PESSOAIS (Gmail/Hotmail) convidadas como "guest" têm
suporte limitado da própria Microsoft pra acessar SharePoint via API do Graph
com o próprio login (funciona só pelo navegador com sessão web, não por
token) — isso é uma limitação de plataforma, não do código deste projeto.
Esse arquivo estático serve de "fallback" pro dashboard: quando o login
delegado do visitante falha ao carregar os dados ao vivo, ele cai pra esse
JSON (modo só leitura). Contas de trabalho/escola (Entra ID) continuam
carregando os dados ao vivo normalmente, sem precisar desse fallback.

Roda a cada 15 minutos via GitHub Actions
(.github/workflows/sincronizar-dados.yml). Também roda manualmente:
    $env:CAP_CLIENT_SECRET = "..."
    python scripts/sincronizar_lancamentos.py
"""

import json
import os
import sys

import msal
import requests

TENANT_ID = "0ea62e2f-152a-4380-bf4b-497083aa0326"
CLIENT_ID = "9291254e-8c79-4641-8d7d-c5771d82ccde"
CLIENT_SECRET = os.environ.get("CAP_CLIENT_SECRET")
if not CLIENT_SECRET:
    sys.exit("ERRO: defina a variável de ambiente CAP_CLIENT_SECRET antes de rodar este script.")

DRIVE_ID = "b!Xyf9Y4l0qUWv1kO5T5rKOn3FJmBOsbRFnpjvpPbvNvrn80E1xObESKagjmnZsZfm"
ITEM_ID = "01ZF6ARXTIQSWRHIKDQVHIHZRFMV5YVCAD"
GRAPH = "https://graph.microsoft.com/v1.0"

SAIDA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "lancamentos.json")


def get_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao obter token: {result.get('error_description')}")
    return result["access_token"]


def ler_tabela_lancamentos(headers):
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook/tables/TabelaLancamentos/rows"
    linhas = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        dados = resp.json()
        linhas.extend(item["values"][0] for item in dados.get("value", []))
        url = dados.get("@odata.nextLink")
    return linhas


def main():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    linhas = ler_tabela_lancamentos(headers)

    # Mesmo formato que a API do Graph devolve ({"value":[{"index":N,"values":[[...]]}]})
    # pra reaproveitar o mesmo parser (graphRowToRow) do dashboard nos dois casos.
    saida = {"value": [{"index": idx, "values": [v]} for idx, v in enumerate(linhas)]}

    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(saida, f, ensure_ascii=False, indent=0, default=str)

    print(f"OK! {len(linhas)} lançamento(s) sincronizado(s) em {SAIDA}.")


if __name__ == "__main__":
    main()
