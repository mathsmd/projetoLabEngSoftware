import pandas as pd

class PlanilhaValidator:
    def __init__(self, path, logger):
        self.path = path
        self.logger = logger
        self.campos_obrigatorios = [
            "CNPJ",
            "VALOR DO PEDIDO",
            "DIMENSÕES CAIXA (altura x largura x comprimento cm)",
            "PESO DO PRODUTO",
            "TIPO DE SERVIÇO JADLOG",
            "TIPO DE SERVIÇO CORREIOS"
        ]

    def ler_e_validar(self):
        self.logger.inicio_tarefa("Validação da Planilha")
        try:
            data = pd.read_excel(self.path, dtype=str)
            saida = []

            for _, row in data.iterrows():
                registro = {}
                for campo in self.campos_obrigatorios:
                    valor = row.get(campo, "")

                    if not pd.notna(valor) or str(valor).strip() == "":
                        registro[campo] = "N/A"
                    else:
                        if campo in ["VALOR DO PEDIDO", "PESO DO PRODUTO"]:
                            try:
                                registro[campo] = float(valor)
                            except ValueError:
                                registro[campo] = "N/A"
                        else:
                            registro[campo] = str(valor).strip()

                registro["STATUS"] = "Sucesso"
                saida.append(registro)

            self.logger.sucesso_tarefa("Validação da Planilha")
            return saida
        except Exception as e:
            self.logger.falha_tarefa("Validação da Planilha", str(e))
            return []
