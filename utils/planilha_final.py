import pandas as pd

class PlanilhaFinal:
    def __init__(self, dados, logger):
        self.dados = dados
        self.logger = logger
    
    def to_dataframe(self):
        """Converte os dados internos para um DataFrame do pandas.

        Returns:
            pd.DataFrame: Dados armazenados convertidos em DataFrame.
        """
        return pd.DataFrame(self.dados)

    def gerar(self, caminho_saida):
        """Gera a planilha final consolidada com validações e colunas organizadas.

        Args:
            caminho_saida (str): Caminho onde a planilha final será salva.

        Returns:
            pd.DataFrame | None: DataFrame gerado caso não ocorra erro, caso contrário None.
        """
        self.logger.inicio_tarefa("Gerar Planilha Final")
        try:
            df_saida = pd.DataFrame(self.dados)

            df_saida["ENDERECO"] = (
                df_saida["LOGRADOURO"].replace("O campo Logradouro está vazio", "").astype(str).str.strip() + " " +
                df_saida["NUMERO"].replace("O campo Número está vazio", "").astype(str).str.strip()
            ).str.strip()

            df_saida.rename(columns={
                "RAZAO_SOCIAL": "RAZÃO SOCIAL",
                "NOME_FANTASIA": "NOME FANTASIA",
                "ENDERECO": "ENDEREÇO",
                "CEP": "CEP",
                "DESCRICAO_MATRIZ_FILIAL": "DESCRIÇÃO MATRIZ FILIAL",
                "DDD_TELEFONE_1": "TELEFONE +DDD",
                "EMAIL": "E-MAIL",
                "VALOR DO PEDIDO": "VALOR DO PEDIDO",
                "DIMENSÕES CAIXA (altura x largura x comprimento cm)": "DIMENSÕES DA CAIXA",
                "PESO DO PRODUTO": "PESO DO PRODUTO",
                "TIPO DE SERVIÇO JADLOG": "TIPO DE SERVIÇO JADLOG",
                "TIPO DE SERVIÇO CORREIOS": "TIPO DE SERVIÇO CORREIOS",
                "STATUS": "STATUS"
            }, inplace=True)

            colunas_ordenadas = [
                "CNPJ", "RAZÃO SOCIAL", "NOME FANTASIA", "ENDEREÇO", "CEP",
                "DESCRIÇÃO MATRIZ FILIAL", "TELEFONE +DDD", "E-MAIL", "VALOR DO PEDIDO",
                "DIMENSÕES DA CAIXA", "PESO DO PRODUTO", "TIPO DE SERVIÇO JADLOG",
                "TIPO DE SERVIÇO CORREIOS", "STATUS"
            ]
            df_saida = df_saida[colunas_ordenadas]
            df_saida["CNPJ"] = df_saida["CNPJ"].astype(str).str.zfill(14)

            df_saida.to_excel(caminho_saida, index=False)
            self.logger.sucesso_tarefa("Gerar Planilha Final")
            self.logger.passo(f"Planilha final salva em {caminho_saida} com {len(df_saida)} registros.")
            return df_saida

        except Exception as e:
            self.logger.falha_tarefa("Gerar Planilha Final", str(e))
            return None
