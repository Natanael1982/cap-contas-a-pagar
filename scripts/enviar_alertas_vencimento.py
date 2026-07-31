"""
Alerta diário de vencimento do dashboard CAP - Contas a Pagar.

Lê a TabelaLancamentos direto da planilha "Gestão Financeira Pessoal.xlsx"
(SharePoint, via Microsoft Graph) e manda UM aviso consolidado (por WhatsApp
via CallMeBot e por e-mail via Graph/Outlook) só para o dono (ver DONO_NOME),
listando — agrupado por pessoa — todo lançamento NÃO PAGO vencendo HOJE ou
AMANHÃ. Só o número/e-mail do dono precisa estar cadastrado em TabelaContatos
(evita ter que pedir confirmação de WhatsApp de cada pessoa da planilha).

Roda 1x por dia via GitHub Actions (.github/workflows/alertas-vencimento.yml).

Este arquivo é publicado no repositório PÚBLICO do CAP no GitHub (roda via
GitHub Actions) — por isso o Client Secret NUNCA fica escrito aqui, só vem
da variável de ambiente CAP_CLIENT_SECRET (Secret do repositório no GitHub).
"""

import os
import sys
from datetime import date, datetime, timedelta

import msal
import requests

try:
    from zoneinfo import ZoneInfo
    HOJE_BR = datetime.now(ZoneInfo("America/Sao_Paulo")).date()
except Exception:
    HOJE_BR = date.today()

TENANT_ID = "0ea62e2f-152a-4380-bf4b-497083aa0326"
CLIENT_ID = "9291254e-8c79-4641-8d7d-c5771d82ccde"
CLIENT_SECRET = os.environ.get("CAP_CLIENT_SECRET")
if not CLIENT_SECRET:
    sys.exit("ERRO: defina a variável de ambiente CAP_CLIENT_SECRET antes de rodar este script.")

DRIVE_ID = "b!Xyf9Y4l0qUWv1kO5T5rKOn3FJmBOsbRFnpjvpPbvNvrn80E1xObESKagjmnZsZfm"
ITEM_ID = "01ZF6ARXTIQSWRHIKDQVHIHZRFMV5YVCAD"

REMETENTE_EMAIL = "natanael.silva@sinerggia.com.br"
DONO_NOME = "Natanael Silva"  # precisa bater com o Nome cadastrado em TabelaContatos
LINK_DASHBOARD = "https://sinerggia-dev.github.io/cap-contas-a-pagar/"

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


def parse_data_br(txt):
    txt = str(txt or "").strip()
    if not txt:
        return None
    try:
        return datetime.strptime(txt, "%d/%m/%Y").date()
    except ValueError:
        return None


def brl(v):
    try:
        return "R$ " + f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "R$ 0,00"


def carregar_lancamentos(headers):
    linhas = ler_tabela(headers, "TabelaLancamentos")
    itens = []
    for v in linhas:
        empresa = str(v[3] or "")
        nome = empresa.split(" - ")[0] if " - " in empresa else empresa
        venc = parse_data_br(v[5])
        pag = parse_data_br(v[6]) if len(v) > 6 else None
        if not venc or pag:
            continue  # sem data de vencimento ou já pago -> não entra no alerta
        desc = str(v[7] or "")
        valor = v[11] if len(v) > 11 else 0
        itens.append({"nome": nome, "venc": venc, "desc": desc, "valor": valor})
    return itens


def carregar_contatos(headers):
    linhas = ler_tabela(headers, "TabelaContatos")
    contatos = {}
    for v in linhas:
        nome = str(v[0] or "").strip()
        if not nome:
            continue
        contatos[nome] = {
            "email": str(v[1] or "").strip(),
            "whatsapp_numero": str(v[2] or "").strip(),
            "whatsapp_apikey": str(v[3] or "").strip(),
        }
    return contatos


def montar_mensagens(itens):
    """Agrupa por nome -> {'hoje': [...], 'amanha': [...]}."""
    amanha = HOJE_BR + timedelta(days=1)
    agrupado = {}
    for item in itens:
        if item["venc"] == HOJE_BR:
            chave = "hoje"
        elif item["venc"] == amanha:
            chave = "amanha"
        else:
            continue
        agrupado.setdefault(item["nome"], {"hoje": [], "amanha": []})[chave].append(item)
    return agrupado


def texto_alerta_consolidado(agrupado):
    """Uma mensagem só, com os vencimentos de todo mundo agrupados por pessoa."""
    linhas = ["📋 CAP - Contas a vencer"]

    def bloco(chave, titulo):
        pessoas_com_itens = {n: g[chave] for n, g in agrupado.items() if g[chave]}
        if not pessoas_com_itens:
            return
        linhas.append(f"\n{titulo}")
        for nome in sorted(pessoas_com_itens):
            linhas.append(f"\n{nome}:")
            for it in pessoas_com_itens[nome]:
                linhas.append(f"- {it['desc']}: {brl(it['valor'])}")

    bloco("hoje", "📅 VENCE HOJE:")
    bloco("amanha", "⏰ VENCE AMANHÃ:")
    linhas.append(f"\nAcesse: {LINK_DASHBOARD}")
    return "\n".join(linhas)


def enviar_email(headers, destino, assunto, corpo):
    url = f"{GRAPH}/users/{REMETENTE_EMAIL}/sendMail"
    body = {
        "message": {
            "subject": assunto,
            "body": {"contentType": "Text", "content": corpo},
            "toRecipients": [{"emailAddress": {"address": destino}}],
        },
        "saveToSentItems": "true",
    }
    resp = requests.post(url, headers={**headers, "Content-Type": "application/json"}, json=body)
    resp.raise_for_status()


def enviar_whatsapp(numero, apikey, texto):
    url = "https://api.callmebot.com/whatsapp.php"
    params = {"phone": numero, "text": texto, "apikey": apikey}
    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()


def main():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    itens = carregar_lancamentos(headers)
    contatos = carregar_contatos(headers)
    agrupado = montar_mensagens(itens)
    agrupado = {n: g for n, g in agrupado.items() if g["hoje"] or g["amanha"]}

    if not agrupado:
        print("Nenhum lançamento vencendo hoje ou amanhã. Nada a avisar.")
        return

    contato = contatos.get(DONO_NOME)
    if not contato:
        print(f"ERRO: '{DONO_NOME}' não está cadastrado em TabelaContatos — nada enviado.")
        return

    texto = texto_alerta_consolidado(agrupado)

    if contato["email"]:
        try:
            enviar_email(headers, contato["email"], "CAP - Contas a vencer", texto)
            print(f"E-mail enviado para {contato['email']}.")
        except Exception as exc:
            print(f"ERRO ao enviar e-mail: {exc}")
    else:
        print("Dono sem e-mail cadastrado — pulando e-mail.")

    if contato["whatsapp_numero"] and contato["whatsapp_apikey"]:
        try:
            enviar_whatsapp(contato["whatsapp_numero"], contato["whatsapp_apikey"], texto)
            print(f"WhatsApp enviado para {contato['whatsapp_numero']}.")
        except Exception as exc:
            print(f"ERRO ao enviar WhatsApp: {exc}")
    else:
        print("Dono sem WhatsApp/apikey cadastrado — pulando WhatsApp.")


if __name__ == "__main__":
    main()
