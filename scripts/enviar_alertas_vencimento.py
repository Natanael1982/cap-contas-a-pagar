"""Versão TEMPORÁRIA de teste — manda só uma mensagem simples via CallMeBot
pro dono, pra isolar se o problema é na entrega do CallMeBot ou em algo do
script original. Será revertida para o script real logo em seguida."""

import os
import sys
from datetime import date, datetime, timedelta

import msal
import requests

try:
    from zoneinfo import ZoneInfo
    HOJE_BR = datetime.now(ZoneInfo("America/Sao_Paulo"))
except Exception:
    HOJE_BR = datetime.now()

TENANT_ID = "0ea62e2f-152a-4380-bf4b-497083aa0326"
CLIENT_ID = "9291254e-8c79-4641-8d7d-c5771d82ccde"
CLIENT_SECRET = os.environ.get("CAP_CLIENT_SECRET")
if not CLIENT_SECRET:
    sys.exit("ERRO: defina a variável de ambiente CAP_CLIENT_SECRET antes de rodar este script.")

DRIVE_ID = "b!Xyf9Y4l0qUWv1kO5T5rKOn3FJmBOsbRFnpjvpPbvNvrn80E1xObESKagjmnZsZfm"
ITEM_ID = "01ZF6ARXTIQSWRHIKDQVHIHZRFMV5YVCAD"
DONO_NOME = "Natanael Silva"
GRAPH = "https://graph.microsoft.com/v1.0"


def get_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=authority, client_credential=CLIENT_SECRET
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Falha ao obter token: {result.get('error_description')}")
    return result["access_token"]


def ler_tabela(headers, nome_tabela):
    url = f"{GRAPH}/drives/{DRIVE_ID}/items/{ITEM_ID}/workbook/tables/{nome_tabela}/rows"
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
    linhas = ler_tabela(headers, "TabelaContatos")
    contato = None
    for v in linhas:
        if str(v[0] or "").strip() == DONO_NOME:
            contato = {"numero": str(v[2] or "").strip(), "apikey": str(v[3] or "").strip()}
            break
    if not contato or not contato["numero"] or not contato["apikey"]:
        sys.exit("ERRO: contato do dono sem WhatsApp/apikey cadastrado.")

    texto = f"Teste CallMeBot direto - {HOJE_BR.strftime('%d/%m/%Y %H:%M:%S')} - se voce recebeu isso, a API esta ok."
    url = "https://api.callmebot.com/whatsapp.php"
    params = {"phone": contato["numero"], "text": texto, "apikey": contato["apikey"]}
    resp = requests.get(url, params=params, timeout=30)
    print("HTTP status:", resp.status_code)
    print("Resposta CallMeBot:", resp.text)


if __name__ == "__main__":
    main()
