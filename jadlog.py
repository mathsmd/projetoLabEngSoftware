import pandas as pd
from botcity.web import WebBot, Browser, By

def executar_simulacao(caminho_excel, saida_excel="saida_final.xlsx", driver_path=None, headless=False):
    bot = WebBot()
    bot.headless = headless
    bot.browser = Browser.CHROME
    if driver_path:
        bot.driver_path = driver_path

    # Abre o site da Jadlog
    bot.browse("https://www.jadlog.com.br/siteInstitucional/simulacao.jad")

    df = pd.read_excel(caminho_excel)

    # Valores Fixos
    cep_origem = "38182428"
    valor_coleta = "50,00"

    if "VALOR COTAÇÃO JADLOG" not in df.columns:
        df["VALOR COTAÇÃO JADLOG"] = ""

    # Pega os dados da planilha e preenche no site da JadLog
    for i, row in df.iterrows():
        try:
            cep_destino = str(row.get("CEP", "")).strip()
            peso = str(row.get("PESO DO PRODUTO", "0")).replace(',', '.').strip()
            if peso in ["", "N/A"]:
                peso = "0"

            modalidade = str(row.get("TIPO DE SERVIÇO JADLOG", "")).strip()

            try:
                dimensoes = str(row.get("DIMENSÕES DA CAIXA", "")).strip()
                altura, largura, comprimento = dimensoes.split(" x ")
            except:
                altura, largura, comprimento = "0", "0", "0"

            valor_mercadoria = str(row.get("VALOR DO PEDIDO", "0")).strip()
            if valor_mercadoria in ["", "N/A"]:
                valor_mercadoria = "0"
            valor_mercadoria = valor_mercadoria.replace(",", ".")
            try:
                valor_mercadoria = str(float(valor_mercadoria)).replace('.', ',')
            except:
                valor_mercadoria = "0"

            bot.find_element("origem", By.ID).clear()
            bot.find_element("origem", By.ID).send_keys(cep_origem)

            bot.find_element("destino", By.ID).clear()
            bot.find_element("destino", By.ID).send_keys(cep_destino)

            bot.find_element("valor_coleta", By.ID).clear()
            bot.find_element("valor_coleta", By.ID).send_keys(valor_coleta)

            bot.find_element("peso", By.ID).clear()
            bot.find_element("peso", By.ID).send_keys(peso)

            bot.find_element("valAltura", By.ID).clear()
            bot.find_element("valAltura", By.ID).send_keys(altura)

            bot.find_element("valLargura", By.ID).clear()
            bot.find_element("valLargura", By.ID).send_keys(largura)

            bot.find_element("valComprimento", By.ID).clear()
            bot.find_element("valComprimento", By.ID).send_keys(comprimento)

            bot.find_element("valor_mercadoria", By.ID).clear()
            bot.find_element("valor_mercadoria", By.ID).send_keys(valor_mercadoria)

            select = bot.find_element("modalidade", By.ID)
            for option in select.find_elements(By.TAG_NAME, "option"):
                if modalidade.lower() in option.text.lower():
                    option.click()
                    break

            bot.find_element("//input[@value='Simular']", By.XPATH).click()
            bot.wait(3000)

        # Tratamento de erros
            try:
                valor_frete = bot.find_element("//div[@id='j_idt45_content']/span", By.XPATH).text
                status = "Sucesso"

                texto = valor_frete.lower().replace("não", "nao").strip()

                if "frete principal nao cadastrado" in texto:
                    valor_frete = "N/A"
                    status = "Erro ao realizar a cotação Jadlog"
                else:
                    status = "Sucesso"
            except:
                valor_frete = "N/A"
                status = "Erro ao realizar a cotação Jadlog"

            df.loc[i, "VALOR COTAÇÃO JADLOG"] = valor_frete

            # Só sobrescreve STATUS se não for sucesso anterior ou se estiver vazio
            if status != "Sucesso" or not str(df.loc[i, "STATUS"]).strip():
                df.loc[i, "STATUS"] = status
            
            print(f"[{i}] {modalidade} -> {valor_frete} | {status}")

        except Exception as e:
            df.loc[i, "VALOR COTAÇÃO JADLOG"] = "N/A"
            df.loc[i, "STATUS"] = f"Erro: {e}"
            print(f"[{i}] Erro: {e}")

    df["CNPJ"] = df["CNPJ"].astype(str)
    df.to_excel(saida_excel, index=False)
    bot.wait(2000)
    bot.stop_browser()
    print(f"\nSimulação concluída. Planilha gerada: {saida_excel}")
    