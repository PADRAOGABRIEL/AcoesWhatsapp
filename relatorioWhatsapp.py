from datetime import datetime, timedelta, timezone
import requests
import time
import pyautogui as pyag
import pyperclip
import subprocess
import sys
import config


sys.stdout.reconfigure(encoding='utf-8')


acoes_desejadas = ['VALE3', 'BBAS3', 'TAEE11', 'CMIG4', 'CSMG3', 'GGBR4', 'KLBN4', 'ISAE4']

token = config.TOKEN

num_wpp = config.NUM_WPP

def obter_dados_acoes_br(acoes):
    """ Faz uma requisição individual para cada ação e retorna os dados do último mês """
    relatorio = []

    for ticker in acoes:
        url = f"https://brapi.dev/api/quote/{ticker}"
        params = {
            'range': '1mo',  # Último mês
            'interval': '1d',  # Dados diários
            'token': token,
        }

        resposta = requests.get(url, params=params)

        if resposta.status_code != 200:
            relatorio.append({"ticker": ticker, "erro": f"Erro {resposta.status_code}"})
            continue

        dados_json = resposta.json()
        acao_info = dados_json.get("results", [])[0]  # Pega o primeiro resultado

        try:
            data_um_mes_atras = datetime.now() - timedelta(days=30)
            timestamp_um_mes_atras = int(data_um_mes_atras.timestamp())

            historico = acao_info.get("historicalDataPrice", [])
            dados_mes_anterior = next((d for d in historico if d["date"] <= timestamp_um_mes_atras), None)

            if not dados_mes_anterior:
                relatorio.append({"ticker": ticker, "erro": "Sem dados históricos suficientes"})
                continue

            data_historica = datetime.fromtimestamp(dados_mes_anterior["date"], tz=timezone.utc).strftime('%Y-%m-%d')

            print(data_historica)

            relatorio.append({
                "ticker": acao_info["symbol"],
                "nome": acao_info["shortName"],
                "preco_atual": round(acao_info["regularMarketPrice"], 2),
                "fechamento_anterior": round(dados_mes_anterior["close"], 2),
                "variacao": round(((acao_info["regularMarketPrice"] / dados_mes_anterior["close"]) - 1) * 100, 2),
                "data": acao_info["regularMarketTime"][:10],  # Data formatada
                "data_mes_anterior": data_historica,  # Data do fechamento anterior
            })
        except KeyError:
            relatorio.append({"ticker": ticker, "erro": "Dados indisponíveis"})

        time.sleep(1)

    return relatorio


def gerar_relatorio(dados_acoes):
    if not dados_acoes:
        return "Não foi possível obter os dados das ações."

    mes_atual = datetime.now().strftime('%m/%Y')
    relatorio = f"📊 *Relatório Mensal de Ações - {mes_atual}*\n\n"

    for dados in dados_acoes:
        if "erro" in dados:
            relatorio += f"⚠️ {dados['ticker']}: {dados['erro']}\n\n"
        else:
            # Determinar emoji com base na variação
            if dados["variacao"] > 0:
                emoji_variacao = "🟢"
            elif dados["variacao"] < 0:
                emoji_variacao = "🔴"
            else:
                emoji_variacao = "⚪"

            relatorio += (
                f"📈 *{dados['ticker']}*\n"
                f"💰 Preço Atual: *R$ {dados['preco_atual']:.2f}*\n"
                f"📉 Fechamento Mês Anterior: *R$ {dados['fechamento_anterior']:.2f}*\n"
                f"🔄 Variação: *{dados['variacao']:.2f}%* {emoji_variacao}\n\n"
            )

    return relatorio.strip()


dados_acoes = obter_dados_acoes_br(acoes_desejadas)
mensagem = gerar_relatorio(dados_acoes)
contato = config.CONTATO_PADRAO


def abrir_whatsapp():
    """ Abre o WhatsApp Desktop instalado pela Microsoft Store """
    try:
        # Primeira tentativa: Comando mais direto para abrir WhatsApp Desktop
        subprocess.run(["cmd", "/c", "start whatsapp://"])

        time.sleep(5)  # Aguarda o aplicativo abrir
    except Exception as e:
        print(f"❌ Erro ao abrir WhatsApp: {e}")


def enviar_whatsapp_desktop(contato, mensagem):
    """ Pesquisa o contato e envia a mensagem automaticamente no WhatsApp Desktop """

    abrir_whatsapp()
    
    time.sleep(2)
    #  Clicar na barra de pesquisa (ajuste se necessário)
    pyag.click(x=1113, y=126)
    time.sleep(2)

    #  Digitar o nome do contato/grupo e pressionar ENTER
    pyperclip.copy(contato)
    pyag.hotkey("ctrl", "v")
    time.sleep(1)
    pyag.click(x=1160, y=180)
    time.sleep(2)
    pyag.click(x=1490, y=1000)

    #  Digitar a mensagem e enviar
    pyperclip.copy(mensagem)
    pyag.hotkey("ctrl", "v")
    time.sleep(1)
    pyag.press("enter")

    print("✅ Mensagem enviada no WhatsApp Desktop!")

# Executar o envio
enviar_whatsapp_desktop(contato, mensagem)