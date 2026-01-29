# Imports
# ==========================================
import pandas as pd
import os
import yagmail
from dotenv import load_dotenv

# BotCity
from botcity.core import DesktopBot
from botcity.web import WebBot
from botcity.maestro import BotMaestroSDK

# Selenium
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    NoAlertPresentException,
    TimeoutException,
    UnexpectedAlertPresentException,
)

# Serviços e Utils
from utils.planilha_validator import PlanilhaValidator
from utils.logger import LoggerExecucao
from utils.planilha_final import PlanilhaFinal
from services.brasil_api import BrasilAPI
from services.RPAChallenge import RPAChallenge
from CotacaoCorreios import CotacaoCorreios
from formatar_planilha import formatar_planilha
from jadlog import executar_simulacao



# Configuração Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

# Função de envio de e-mail
def enviar_email_app_password():
    # Carrega variáveis do arquivo .env
    load_dotenv()
    
    # Obtém as credenciais do ambiente
    app_password = os.getenv("EMAIL_APP_PASSWORD")
    remetente = os.getenv("EMAIL_REMETENTE")
    destinatario = "mellaniemel4@gmail.com"
    
    if not app_password or not remetente:
        raise ValueError("Configure EMAIL_APP_PASSWORD e EMAIL_REMETENTE no arquivo .env")
    
    try:
        # Usa o App Password em vez da senha normal
        yag = yagmail.SMTP(user=remetente, password=app_password)
        
        assunto = "Bot Concluído com Sucesso"
        corpo = """
        Olá,

        O bot foi executado e concluído com sucesso!
        
        Segue em anexo a planilha de saída com os resultados processados.
        """
        
        yag.send(
            to=destinatario,
            subject=assunto,
            contents=corpo,
            attachments="saida_final.xlsx"
        )
        print("E-mail enviado com sucesso usando App Password!")
        
    except Exception as e:
        print(f"Falha ao enviar e-mail: {e}")

def ajustar_status(row):
        # Se já tiver erro anterior, mantém
        if str(row["STATUS"]).lower().startswith("erro"):
            return row["STATUS"]

        # Se Correios falhou
        if str(row.get("VALOR COTAÇÃO CORREIOS", "")).upper() == "N/A":
            return "Erro ao realizar a cotação Correios"

        # Se Jadlog falhou
        if str(row.get("VALOR COTAÇÃO JADLOG", "")).upper() == "N/A":
            return "Erro ao realizar a cotação Jadlog"

        # Caso contrário, está tudo certo
        return "Sucesso"


def main():
    # Conexão Maestro
    maestro = BotMaestroSDK.from_sys_args()
    try:
        execution = maestro.get_execution()
        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")
    except Exception:
        print("Rodando localmente (sem Maestro).")

    # Bots
    desktop_bot = DesktopBot()
    webbot = WebBot()
    webbot.headless = False
    webbot.driver_path = r"C:\ProjetoFinalCompass\resources\chromedriver-win64\chromedriver.exe"

    # Logger
    logger = LoggerExecucao(nome_processo="RPA_Logger")

    # Validação da planilha
    validator = PlanilhaValidator(
        r"C:\ProjetoFinalCompass\Cadastro de Clientes no Sistema Challenge e Cotação de Novos Pedidos\Processar\Grupo 5.xlsx",
        logger,
    )
    dados_validos = validator.ler_e_validar()

    # Consulta API
    api = BrasilAPI(logger)
    saida = [registro | api.consultar(registro["CNPJ"]) for registro in dados_validos]

    # Planilha Final
    planilha = PlanilhaFinal(saida, logger)
    df = planilha.gerar("saida_final.xlsx")

    rpaBot = RPAChallenge(webbot)
    df = rpaBot.processarDados(planilha.to_dataframe())

    try:
        webbot.stop_browser()
    except:
        pass

    webbot.start_browser()
    webbot.browse("https://www2.correios.com.br/sistemas/precosPrazos/")


    # Cotação Correios
    df = pd.read_excel("saida_final.xlsx", dtype={"CNPJ": str})

    if "VALOR COTAÇÃO CORREIOS" not in df.columns:
        df["VALOR COTAÇÃO CORREIOS"] = ""

    cotador = CotacaoCorreios(webbot)

    for idx, row in df.iterrows():
        cep_destino = str(row["CEP"])
        servico = str(row["TIPO DE SERVIÇO CORREIOS"])
        peso = str(row["PESO DO PRODUTO"])
        dimensoes = str(row["DIMENSÕES DA CAIXA"]).split("x")

        if "vazio" in cep_destino.lower() or "vazio" in peso.lower():
            df.at[idx, "VALOR COTAÇÃO CORREIOS"] = "N/A"
            df.at[idx, "STATUS"] = "Erro: dados ausentes"
            continue

        try:
            altura, largura, comprimento = [d.strip() for d in dimensoes]
        except Exception:
            df.at[idx, "VALOR COTAÇÃO CORREIOS"] = "N/A"
            df.at[idx, "STATUS"] = "Erro: dimensões inválidas"
            continue

        valor = cotador.realizar_cotacao(servico, altura, largura, comprimento, peso, cep_destino)

        if valor and not valor.startswith("Erro"):
            df.at[idx, "VALOR COTAÇÃO CORREIOS"] = valor
            df.at[idx, "STATUS"] = "Sucesso"
        else:
            df.at[idx, "VALOR COTAÇÃO CORREIOS"] = "N/A"
            df.at[idx, "STATUS"] = "Erro ao realizar a cotação Correios"

    df.to_excel("saida_final.xlsx", index=False)
    print("Planilha atualizada com as cotações dos Correios.")

    # Cotação Jadlog
    executar_simulacao(
        caminho_excel="saida_final.xlsx",
        saida_excel="saida_final.xlsx",
        driver_path=r"C:\ProjetoFinalCompass\resources\chromedriver-win64\chromedriver.exe",
        headless=False,
    )

    print("Cotação Jadlog finalizada e planilha atualizada.")

    df = pd.read_excel("saida_final.xlsx", dtype={"CNPJ": str})

    # Garantir N/A em todas células vazias
    df = df.fillna("N/A")
    for col in df.columns:
        df[col] = df[col].apply(lambda x: x if str(x).strip() not in ["", "nan", "NaN"] else "N/A")

    df["STATUS"] = df.apply(ajustar_status, axis=1)

    # Salvar antes da formatação
    df.to_excel("saida_final.xlsx", index=False)

    # Aplicar formatação final
    formatar_planilha("saida_final.xlsx")

    try:
        webbot.stop_browser()
    except:
        print("Navegador já estava fechado, ignorando...")

    webbot.wait(3000)
    
    # Enviar e-mail com anexo
    # ==========================================
    enviar_email_app_password()


    def not_found(label):
        print(f"Element not found: {label}")


if __name__ == "__main__":
    main()