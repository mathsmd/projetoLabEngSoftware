import requests
import re

class BrasilAPI:
    CAMPOS_PADRAO = {
        "razao_social": "Razão Social",
        "nome_fantasia": "Nome Fantasia",
        "situacao_cadastral": "Situação Cadastral",
        "logradouro": "Logradouro",
        "numero": "Número",
        "municipio": "Município",
        "cep": "CEP",
        "descricao_matriz_filial": "Descrição Matriz/Filial",
        "ddd_telefone_1": "Telefone",
        "email": "E-mail"
    }

    def __init__(self, logger):
        self.logger = logger

    def consultar(self, cnpj):
        self.logger.inicio_tarefa("Consulta Brasil API")
        cnpj = re.sub(r'\D', '', str(cnpj))

        if len(cnpj) != 14:
            self.logger.falha_tarefa("Consulta Brasil API", "CNPJ inválido")
            return {k.upper(): "N/A" for k in self.CAMPOS_PADRAO.keys()} | {"STATUS": "CNPJ inválido"}

        url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj}"
        try:
            r = requests.get(url, timeout=10)
            if r.status_code == 200:
                data = r.json()
                registro = {}

                for chave_api, nome_legivel in self.CAMPOS_PADRAO.items():
                    valor = data.get(chave_api)
                    registro[chave_api.upper()] = valor if valor else "N/A"

                situacao = str(data.get("situacao_cadastral", "")).lower()
                registro["STATUS"] = "Empresa Inativa" if situacao == "inativo" else "Sucesso"

                self.logger.sucesso_tarefa(f"Consulta Brasil API concluída para CNPJ: {cnpj}")
                return registro
            else:
                msg = "CNPJ não encontrado" if r.status_code == 404 else f"Erro na API: {r.status_code}"
                self.logger.falha_tarefa("Consulta Brasil API", msg)
                return {k.upper(): "N/A" for k in self.CAMPOS_PADRAO.keys()} | {"STATUS": msg}

        except Exception as e:
            self.logger.falha_tarefa("Consulta Brasil API", f"Erro na consulta: {e}")
            return {k.upper(): "N/A" for k in self.CAMPOS_PADRAO.keys()} | {"STATUS": f"Erro na consulta: {e}"}
