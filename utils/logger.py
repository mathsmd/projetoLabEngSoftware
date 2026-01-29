import os
from datetime import datetime
from pathlib import Path

class LoggerExecucao:
    def __init__(self, nome_processo="RPA_Processo"):
        """
        Cria uma pasta de logs com base no dia da execução
        e inicia o arquivo de log do processo.
        """
        self.nome_processo = nome_processo
        data_execucao = datetime.now().strftime("%d%m%Y")

        # Pasta Logs/DDMMAAAA
        self.base_path = Path("Logs") / data_execucao
        self.base_path.mkdir(parents=True, exist_ok=True)

        # Nome do arquivo de log
        log_filename = f"{data_execucao}_{self.nome_processo}.log"
        self.log_path = self.base_path / log_filename

        # Abre arquivo de log
        self.log_file = open(self.log_path, "a", encoding="utf-8")

    def _log(self, mensagem):
        """Escreve uma linha no log com timestamp"""
        timestamp = datetime.now().strftime("(%d/%m/%Y %H:%M:%S)")
        linha = f"{timestamp} {mensagem}"
        print(linha)  # também mostra no console
        self.log_file.write(linha + "\n")
        self.log_file.flush()

    # ========= Métodos principais =========

    def inicio_tarefa(self, nome_tarefa):
        self._log(f"Início - {nome_tarefa}")

    def sucesso_tarefa(self, nome_tarefa):
        self._log(f"Sucesso - {nome_tarefa}")

    def falha_tarefa(self, nome_tarefa, motivo=""):
        self._log(f"Falha - {nome_tarefa}. Motivo: {motivo}")

    def passo(self, descricao):
        """Log de um passo intermediário (ex: preencher campo, capturar valor)."""
        self._log(descricao)

    def captura_tela(self, webbot=None):
        """
        Captura uma screenshot com timestamp.
        Pode receber o webbot (se estiver usando bot de navegador).
        """
        ts = datetime.now().strftime("%d%m%Y_%H%M%S")
        filename = f"{ts} - {self.nome_processo}.jpg"
        caminho = self.base_path / filename

        if webbot:
            try:
                webbot.screenshot(str(caminho))
                self._log(f"Screenshot salva: {caminho}")
            except Exception as e:
                self._log(f"Erro ao capturar tela: {e}")
        else:
            self._log("Captura de tela ignorada (webbot não informado).")

    def fechar(self):
        """Fecha o arquivo de log"""
        self.log_file.close()
