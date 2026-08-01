"""
Sincroniza a TabelaUsuarios (SharePoint, arquivo separado CAP_Usuarios.xlsx)
para um arquivo estático "usuarios.json" no MESMO repositório do GitHub Pages
do CAP — assim, qualquer conta Microsoft (inclusive pessoal/Gmail, sem nenhum
acesso prévio ao SharePoint) consegue checar aprovação de acesso via um fetch
simples de mesma origem, sem esbarrar em CORS nem em limitações de permissão
cross-tenant do Graph.

Roda a cada 15 minutos via GitHub Actions
(.github/workflows/sincronizar-usuarios.yml). Também roda manualmente:
    $env:CAP_CLIENT_SECRET = "..."
    python scripts/sincronizar_usuarios.py

Este arquivo é publicado no repositório PÚBLICO do CAP — o Client Secret
nunca fica escrito aqui, só vem da variável de ambiente CAP_CLIENT_SECRET.
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
USERS_ITEM_ID = "01ZF6ARXVGGGKKJ2IAPVCJR6FXIOB5XP35"
GRAPH = "https://graph.microsoft.com/v1.0"

SAIDA = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "usuarios.json")


def get_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET)
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao obter token: {result.get('error_description')}")
    return result["access_token"]


def ler_tabela_usuarios(headers):
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{USERS_ITEM_ID}/workbook/tables/TabelaUsuarios/rows"
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
    linhas = ler_tabela_usuarios(headers)

    usuarios = []
    for v in linhas:
        usuarios.append({
            "email": str(v[0] or "").strip(),
            "papel": v[1] or "",
            "status": v[2] or "",
            "data": v[3] or "",
            "filtroDespesaDe": str(v[4] or "").strip() if len(v) > 4 else "",
            "nome": str(v[5] or "").strip() if len(v) > 5 else "",
        })

    with open(SAIDA, "w", encoding="utf-8") as f:
        json.dump(usuarios, f, ensure_ascii=False, indent=2)

    print(f"OK! {len(usuarios)} usuário(s) sincronizado(s) em {SAIDA}.")


if __name__ == "__main__":
    main()
