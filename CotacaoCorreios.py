# services/cotacao_correios.py

from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    NoSuchElementException,
    TimeoutException,
    UnexpectedAlertPresentException
)
from botcity.web import By

class CotacaoCorreios:
    def __init__(self, webbot):
        """
        Classe responsável por realizar cotações no site dos Correios.
        :param webbot: Instância de WebBot já inicializada.
        """
        self.webbot = webbot
        self.driver = webbot.driver

    def _has_alert(self):
        try:
            self.driver.switch_to.alert
            return True
        except Exception:
            return False

    def realizar_cotacao(self, service, altura, largura, comprimento, peso, cep_destino):
        """
        Realiza a cotação no site dos Correios.
        :param service: Tipo de serviço (PAC, SEDEX)
        :param altura: Altura da embalagem
        :param largura: Largura da embalagem
        :param comprimento: Comprimento da embalagem
        :param peso: Peso em kg
        :param cep_destino: CEP de destino
        :return: Valor do frete ou mensagem de erro
        """
        driver = self.driver
        try:
            # CEP fixo de origem
            self.webbot.find_element("cepOrigem", By.NAME).clear()
            self.webbot.find_element("cepOrigem", By.NAME).send_keys("38182-428")

            # CEP destino
            self.webbot.find_element("cepDestino", By.NAME).clear()
            self.webbot.find_element("cepDestino", By.NAME).send_keys(str(cep_destino))

            # Serviço
            servico_map = {"PAC": "04510", "SEDEX": "04014"}
            Select(self.webbot.find_element("servico", By.NAME)).select_by_value(
                servico_map.get(service.upper(), "04510")
            )

            # Embalagem -> outra
            Select(self.webbot.find_element("embalagem1", By.NAME)).select_by_value("outraEmbalagem1")

            # Dimensões
            for nm, val in [("Altura", altura), ("Largura", largura), ("Comprimento", comprimento)]:
                self.webbot.find_element(nm, By.NAME).clear()
                self.webbot.find_element(nm, By.NAME).send_keys(str(val))

            # Peso (1..30)
            Select(self.webbot.find_element("peso", By.NAME)).select_by_value(str(int(float(peso))))

            # Calcular
            old_handles = driver.window_handles[:]
            self.webbot.find_element("Calcular", By.NAME).click()

            # Espera abrir nova aba OU aparecer alerta
            WebDriverWait(driver, 10).until(
                lambda d: len(d.window_handles) > len(old_handles) or self._has_alert()
            )

            # Se abriu alerta
            if self._has_alert():
                try:
                    msg = driver.switch_to.alert.text
                    driver.switch_to.alert.accept()
                except Exception:
                    msg = "Alerta sem texto"
                if len(driver.window_handles) > len(old_handles):
                    driver.switch_to.window(driver.window_handles[-1])
                    driver.close()
                    driver.switch_to.window(old_handles[0])
                return f"Erro: {msg}"

            # Se abriu aba de resultado
            if len(driver.window_handles) > len(old_handles):
                driver.switch_to.window(driver.window_handles[-1])

                try:
                    WebDriverWait(driver, 8).until(
                        EC.presence_of_element_located((By.XPATH, "//table[contains(@class,'comparaResult')]"))
                    )
                except TimeoutException:
                    driver.close()
                    driver.switch_to.window(old_handles[0])
                    return "Erro: preço não encontrado"

                valor = None
                try:
                    td = driver.find_elements(By.CSS_SELECTOR, "table.comparaResult tfoot.destaque td")
                    if td:
                        valor = td[-1].text.strip()
                except Exception:
                    pass

                if not valor:
                    try:
                        valor = driver.find_element(
                            By.XPATH,
                            "//table[contains(@class,'comparaResult')]//th[normalize-space(.)='Preço do serviço:']/following-sibling::td[1]"
                        ).text.strip()
                    except NoSuchElementException:
                        pass

                if not valor:
                    try:
                        tds_rs = driver.find_elements(
                            By.XPATH, "//table[contains(@class,'comparaResult')]//td[contains(normalize-space(.),'R$')]"
                        )
                        if tds_rs:
                            valor = tds_rs[-1].text.strip()
                    except Exception:
                        pass

                driver.close()
                driver.switch_to.window(old_handles[0])

                return valor if valor else "Erro: preço não encontrado"

            return "Erro: nenhuma aba de resultado"

        except UnexpectedAlertPresentException:
            try:
                msg = driver.switch_to.alert.text
                driver.switch_to.alert.accept()
            except Exception:
                msg = "Alerta inesperado"
            try:
                handles = driver.window_handles
                if len(handles) > 1:
                    driver.switch_to.window(handles[-1])
                    driver.close()
                    driver.switch_to.window(handles[0])
            except Exception:
                pass
            return f"Erro: {msg}"

        except Exception as e:
            try:
                handles = driver.window_handles
                if len(handles) > 1:
                    driver.switch_to.window(handles[-1])
                    driver.close()
                    driver.switch_to.window(handles[0])
            except Exception:
                pass
            return f"Erro: {e}"
