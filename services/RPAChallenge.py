from botcity.web import WebBot, By 
from utils.logger import LoggerExecucao
import pandas as pd
from typing import List, Dict
import time

class RPAChallenge:
    def __init__(self, webbot=None, logger=None):
        self.logger = logger or LoggerExecucao("RPA_Challenge")
        self.webbot = webbot or WebBot()
        self.webbot.delay_between_actions = 1000
        self.webbot.headless = False
        self.url = "https://www.rpachallenge.com/"

    def iniciarDesafio(self) -> bool:
        try:
            self.logger.inicio_tarefa("Acessando o site do RPA Challenge")
            self.webbot.browse(self.url)
            self.webbot.wait(3000)

            self.logger.inicio_tarefa("Clicando no Botao Start")
            buttonStart = self.webbot.find_element("//button[text()='Start']", By.XPATH)
            buttonStart.click()
            self.webbot.wait(2000)

            self.logger.sucesso_tarefa("Desafio iniciado")
            return True
        except Exception as e:
            self.logger.falha_tarefa("Iniciar desafio", str(e))
            return False

    def preencherFormulario(self, dados: dict) -> bool:
        try:
            campos = [
                ("First Name", "labelFirstName", dados['RAZAO SOCIAL']),
                ("Last Name", "labelLastName", dados.get('SITUACAO CADASTRAL', 'N/A')),
                ("Company Name", "labelCompanyName", dados['NOME FANTASIA']),
                ("Role", "labelRole", dados['DESCRICAO MATRIZ FILIAL']),
                ("Address", "labelAddress", dados['ENDERECO']),
                ("Email", "labelEmail", dados['E-MAIL']),
                ("Phone", "labelPhone", dados['TELEFONE +DDD'])
            ]

            for nomeCampo, refNome, valor in campos:
                try:
                    self.logger.passo(f"Preenchendo {nomeCampo}: {valor}")
                    xpath = f"//input[@ng-reflect-name='{refNome}']"
                    campo = self.webbot.find_element(xpath, By.XPATH)
                except Exception:
                    xpathAlt = f"//input[contains(@ng-reflect-name, '{refNome.replace('label', '')}')]"
                    self.logger.passo(f"Tentando XPath alternativo para {nomeCampo}")
                    try:
                        campo = self.webbot.find_element(xpathAlt, By.XPATH)
                    except Exception as e:
                        self.logger.falha_tarefa(f"Campo {nomeCampo}", f"Nao encontrado ({e})")
                        return False

                try:
                    campo.clear()
                    campo.send_keys(str(valor))
                    time.sleep(0.5)
                except Exception as e:
                    self.logger.falha_tarefa(f"Preencher {nomeCampo}", str(e))
                    return False

            # Clicar no botão Submit
            try:
                self.logger.inicio_tarefa("Clicando no botao Submit")
                buttonSubmit = self.webbot.find_element("//input[@value='Submit']", By.XPATH)
                buttonSubmit.click()
                self.logger.sucesso_tarefa("Formulario enviado")
                return True
            except Exception as e:
                self.logger.falha_tarefa("Clicar no botao Submit", str(e))
                return False

        except Exception as e:
            self.logger.falha_tarefa("Preencher formulario", str(e))
            return False

    def processarDados(self, df_dados: pd.DataFrame) -> pd.DataFrame:
        self.logger.inicio_tarefa(f"Processando {len(df_dados)} registros")

        if 'STATUS' not in df_dados.columns:
            df_dados['STATUS'] = ''

        if not self.iniciarDesafio():
            df_dados['STATUS'] = 'Nao conseguiu iniciar'
            return df_dados

        contador = 0
        sucessos = 0
        falhas = 0

        for index, row in df_dados.iterrows():
            self.logger.inicio_tarefa(f"Preenchendo formulario {index + 1}/{len(df_dados)}")
            if self.preencherFormulario(row.to_dict()):
                df_dados.at[index, 'STATUS'] = 'Preenchido com Sucesso'
                self.logger.sucesso_tarefa(f"Formulario {index + 1}")
                sucessos += 1
            else:
                df_dados.at[index, 'STATUS'] = 'Falha ao Preencher'
                self.logger.falha_tarefa(f"Formulario {index + 1}", "Erro no preenchimento")
                falhas += 1
            contador += 1

            if contador == 10 and index + 1 < len(df_dados):
                self.logger.inicio_tarefa("Reiniciando desafio apos 10 formulários")
                self.reiniciarDesafio()
                contador = 0

        self.logger.sucesso_tarefa(f"Processamento concluido: Sucessos={sucessos}, Falhas={falhas}")
        return df_dados

    def reiniciarDesafio(self):
        try:
            self.webbot.browse(self.url)
            time.sleep(3)
            buttonStart = self.webbot.find_element("//button[text()='Start']", By.XPATH)
            buttonStart.click()
            time.sleep(1.5)
            self.logger.sucesso_tarefa("Desafio reiniciado")
        except Exception as e:
            self.logger.falha_tarefa("Reiniciar desafio", str(e))

    def finalizar(self):
        self.logger.inicio_tarefa("Finalizando execucao")
        self.webbot.wait(3000)
        self.webbot.stop_browser()
        self.logger.sucesso_tarefa("Execucao finalizada")
        self.logger.fechar()
            
            
            
            
            
            
            
            
            
            
            
            
            
                
            
        
