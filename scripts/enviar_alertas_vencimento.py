"""
Alerta diário de vencimento do dashboard CAP - Contas a Pagar.

Lê a TabelaLancamentos direto da planilha "Gestão Financeira Pessoal.xlsx"
(SharePoint, via Microsoft Graph) e manda um aviso só para o dono (ver
DONO_NOME): por e-mail (Graph/Outlook) é UMA mensagem consolidada; por
WhatsApp (CallMeBot) é uma mensagem separada por bloco de pessoa (mensagem
única grande demais estava sendo cortada no meio pelo CallMeBot), sempre
listando — agrupado por pessoa — todo lançamento NÃO PAGO vencendo HOJE,
AMANHÃ ou em 2 DIAS. Um lançamento é considerado pago (e fica fora do alerta)
se a Data do Pagamento estiver preenchida OU se o campo Status estiver como
"Pago" (mesmo sem data de pagamento preenchida). Só o número/e-mail do dono
precisa estar cadastrado em TabelaContatos
(evita ter que pedir confirmação de WhatsApp de cada pessoa da planilha).

Pensado para rodar 1x por dia via GitHub Actions (.github/workflows/
alertas-vencimento.yml), mas também roda manualmente:
    python "Enviar Alertas Vencimento.py"

Este arquivo é publicado no repositório PÚBLICO do CAP no GitHub (roda via
GitHub Actions) — por isso o Client Secret NUNCA fica escrito aqui, só vem
da variável de ambiente CAP_CLIENT_SECRET (no Actions, de um Secret do repo;
localmente, defina antes de rodar: $env:CAP_CLIENT_SECRET = "...").

Variável de ambiente obrigatória:
    CAP_CLIENT_SECRET  -> Client Secret do App Registration "Gestão Financeira Claude"
"""

import base64
import os
import sys
import time
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


def parse_data_planilha(v):
    """Data pode vir como número serial do Excel (ex.: 46233) ou como texto dd/mm/aaaa."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return date(1899, 12, 30) + timedelta(days=int(v))
    txt = str(v).strip()
    if not txt:
        return None
    if txt.replace(".", "", 1).isdigit():
        return date(1899, 12, 30) + timedelta(days=int(float(txt)))
    try:
        return datetime.strptime(txt, "%d/%m/%Y").date()
    except ValueError:
        return None


def parse_valor(v):
    """Valor pode vir como número puro ou como texto em formato BR ('1.038,82'), conforme
    a célula da planilha estiver formatada — mesmo tratamento que o dashboard faz em parseBRL()."""
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or v == "":
        return 0.0
    txt = str(v).strip().replace("R$", "").strip()
    txt = txt.replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return 0.0


def brl(v):
    return "R$ " + f"{parse_valor(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def carregar_lancamentos(headers):
    linhas = ler_tabela(headers, "TabelaLancamentos")
    itens = []
    for v in linhas:
        empresa = str(v[3] or "")
        nome = empresa.split(" - ")[0] if " - " in empresa else empresa
        venc = parse_data_planilha(v[5])
        pag = parse_data_planilha(v[6]) if len(v) > 6 else None
        status = str(v[9] or "").strip().lower() if len(v) > 9 else ""
        if not venc or pag or status == "pago":
            continue  # sem data de vencimento ou já pago (data ou status) -> não entra no alerta
        desc = str(v[7] or "")
        valor = v[11] if len(v) > 11 else 0
        anexo = str(v[13] or "") if len(v) > 13 else ""
        itens.append({"nome": nome, "venc": venc, "desc": desc, "valor": valor, "anexo": anexo})
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
    """Agrupa por nome -> {'hoje': [...], 'amanha': [...], 'em_2_dias': [...]}."""
    amanha = HOJE_BR + timedelta(days=1)
    em_2_dias = HOJE_BR + timedelta(days=2)
    agrupado = {}
    for item in itens:
        if item["venc"] == HOJE_BR:
            chave = "hoje"
        elif item["venc"] == amanha:
            chave = "amanha"
        elif item["venc"] == em_2_dias:
            chave = "em_2_dias"
        else:
            continue
        agrupado.setdefault(item["nome"], {"hoje": [], "amanha": [], "em_2_dias": []})[chave].append(item)
    return agrupado


SECAO_EMOJI = {"hoje": "🔴 ", "amanha": "🟠 ", "em_2_dias": "🟡 "}
DIAS_SEMANA = ["Segunda-Feira", "Terça-Feira", "Quarta-Feira", "Quinta-Feira",
               "Sexta-Feira", "Sábado", "Domingo"]
SEPARADOR = "- " * 26


def montar_blocos_alerta(agrupado):
    """Monta a lista de blocos do alerta: [título, um bloco por pessoa (repetindo
    o cabeçalho VENCE HOJE/AMANHÃ/EM 2 DIAS + data + dia da semana), link].

    Cada bloco de pessoa tem nome em *negrito* (não itálico), itens com bullet
    e "Total:" — sem "Subtotal" por seção nem "Total geral". Vira lista de
    blocos (não uma string só) porque o WhatsApp/CallMeBot manda cada bloco
    como mensagem separada (ver enviar_whatsapp_blocos) — mensagem única
    grande demais estava sendo cortada no meio pelo CallMeBot.
    """
    blocos = ["📋 *CAP - Contas a Vencer*"]

    def bloco_dia(chave, titulo, data_ref):
        pessoas_com_itens = {n: g[chave] for n, g in agrupado.items() if g[chave]}
        if not pessoas_com_itens:
            return
        cabecalho = f"{titulo} ({data_ref.strftime('%d/%m')}) - {DIAS_SEMANA[data_ref.weekday()]}"
        primeiro_da_secao = True
        for nome in sorted(pessoas_com_itens):
            prefixo = SECAO_EMOJI[chave] if primeiro_da_secao else ""
            primeiro_da_secao = False
            linhas = [f"{prefixo}{cabecalho}", f"👤 *{nome}*"]
            total_pessoa = 0.0
            for it in pessoas_com_itens[nome]:
                total_pessoa += parse_valor(it["valor"])
                linhas.append(f"   • {it['desc']}: {brl(it['valor'])}")
            linhas.append(f"Total: {brl(total_pessoa)}")
            blocos.append("\n".join(linhas))

    bloco_dia("hoje", "VENCE HOJE", HOJE_BR)
    bloco_dia("amanha", "VENCE AMANHÃ", HOJE_BR + timedelta(days=1))
    bloco_dia("em_2_dias", "VENCE EM 2 DIAS", HOJE_BR + timedelta(days=2))

    blocos.append(f"🔗 Acesse: {LINK_DASHBOARD}")
    return blocos


def texto_alerta_consolidado(agrupado):
    """Texto único pro e-mail: os mesmos blocos, concatenados com separador tracejado."""
    return f"\n{SEPARADOR}\n".join(montar_blocos_alerta(agrupado))


def primeiro_anexo(agrupado):
    """Primeiro lançamento (hoje antes de amanhã antes de 2 dias, ordem alfabética por pessoa) com boleto anexado."""
    for chave in ("hoje", "amanha", "em_2_dias"):
        for nome in sorted(agrupado):
            for it in agrupado[nome][chave]:
                if it.get("anexo"):
                    return it
    return None


def baixar_anexo(headers, url):
    """Baixa um arquivo do SharePoint a partir do link salvo na planilha, via API de Shares do Graph."""
    share_id = "u!" + base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")

    meta_resp = requests.get(f"{GRAPH}/shares/{share_id}/driveItem", headers=headers)
    meta_resp.raise_for_status()
    nome_arquivo = meta_resp.json().get("name") or "boleto.pdf"

    conteudo_resp = requests.get(f"{GRAPH}/shares/{share_id}/driveItem/content", headers=headers)
    conteudo_resp.raise_for_status()
    return nome_arquivo, conteudo_resp.content


def enviar_email(headers, destino, assunto, corpo, anexo=None):
    message = {
        "subject": assunto,
        "body": {"contentType": "Text", "content": corpo},
        "toRecipients": [{"emailAddress": {"address": destino}}],
    }
    if anexo:
        nome_arquivo, conteudo_bytes = anexo
        message["attachments"] = [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": nome_arquivo,
            "contentBytes": base64.b64encode(conteudo_bytes).decode("ascii"),
        }]
    url = f"{GRAPH}/users/{REMETENTE_EMAIL}/sendMail"
    body = {"message": message, "saveToSentItems": "true"}
    resp = requests.post(url, headers={**headers, "Content-Type": "application/json"}, json=body)
    resp.raise_for_status()


def enviar_whatsapp(numero, apikey, texto):
    url = "https://api.callmebot.com/whatsapp.php"
    params = {"phone": numero, "text": texto, "apikey": apikey}
    resp = requests.get(url, params=params, timeout=30)
    resp.raise_for_status()


def enviar_whatsapp_blocos(numero, apikey, blocos):
    """Manda cada bloco como uma mensagem separada (uma pausa curta entre elas
    pra não esbarrar no limite anti-spam do CallMeBot). Uma mensagem única
    grande demais (~1.300+ caracteres) estava sendo cortada no meio pelo
    CallMeBot — mensagens menores e separadas resolvem isso de vez, e cada
    bloco já fica visualmente separado por ser uma mensagem própria."""
    falhas = 0
    for i, bloco in enumerate(blocos):
        try:
            enviar_whatsapp(numero, apikey, bloco)
        except Exception as exc:
            falhas += 1
            print(f"ERRO ao enviar bloco {i + 1}/{len(blocos)} do WhatsApp: {exc}")
        if i < len(blocos) - 1:
            time.sleep(3)
    return falhas


def main():
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    itens = carregar_lancamentos(headers)
    contatos = carregar_contatos(headers)
    agrupado = montar_mensagens(itens)
    agrupado = {n: g for n, g in agrupado.items() if g["hoje"] or g["amanha"] or g["em_2_dias"]}

    if not agrupado:
        print("Nenhum lançamento vencendo hoje, amanhã ou em 2 dias. Nada a avisar.")
        return

    contato = contatos.get(DONO_NOME)
    if not contato:
        print(f"ERRO: '{DONO_NOME}' não está cadastrado em TabelaContatos — nada enviado.")
        return

    texto = texto_alerta_consolidado(agrupado)

    anexo = None
    item_com_anexo = primeiro_anexo(agrupado)
    if item_com_anexo:
        try:
            anexo = baixar_anexo(headers, item_com_anexo["anexo"])
            print(f"Anexo de '{item_com_anexo['desc']}' baixado ({anexo[0]}).")
        except Exception as exc:
            print(f"Aviso: falha ao baixar anexo de '{item_com_anexo['desc']}': {exc}")

    if contato["email"]:
        try:
            enviar_email(headers, contato["email"], "CAP - Contas a vencer", texto, anexo=anexo)
            print(f"E-mail enviado para {contato['email']}.")
        except Exception as exc:
            print(f"ERRO ao enviar e-mail: {exc}")
    else:
        print("Dono sem e-mail cadastrado — pulando e-mail.")

    if contato["whatsapp_numero"] and contato["whatsapp_apikey"]:
        blocos = montar_blocos_alerta(agrupado)
        falhas = enviar_whatsapp_blocos(contato["whatsapp_numero"], contato["whatsapp_apikey"], blocos)
        if falhas:
            print(f"WhatsApp: {len(blocos) - falhas}/{len(blocos)} mensagens enviadas para {contato['whatsapp_numero']} ({falhas} falharam).")
        else:
            print(f"WhatsApp: {len(blocos)} mensagens enviadas para {contato['whatsapp_numero']}.")
    else:
        print("Dono sem WhatsApp/apikey cadastrado — pulando WhatsApp.")


if __name__ == "__main__":
    main()
