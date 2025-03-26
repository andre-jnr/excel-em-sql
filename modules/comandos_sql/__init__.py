def ler_base_de_dados(caminho_arquivo):
  """Ler a base de dados em excel e converte para uma lista Python

  Args:
      caminho_arquivo (str): caminho do arquivo para transformar em lista Python

  Returns:
      list: retorna uma lista Python com todos os dados da tabela
  """
  import pandas as pd

  df = pd.read_excel(caminho_arquivo)

  dados_lista = df.values.tolist()

  return dados_lista


def converter_string(lista):
  """Converte todos os dados da lista em String

  Args:
      lista (list): lista, onde os dados serão convertidos

  Returns:
      list: retorna a lista com todos os dados convertidos para String
  """

  for i, dado in enumerate(lista):
    for index, informacao in enumerate(dado):
      lista[i][index] = str(informacao)

  return lista


def criar_comando_sql(lista_de_strings):
  comando = """
  INSERT INTO temp_sped_estoque (codigo, descricao, grupo, estoque, unidade, custo)  
  VALUES  
  """

  for dado in lista_de_strings:
    comando += f"('{dado[0]}', '{dado[1].strip()}', '{dado[2].strip()}', {dado[3]}, '{dado[4].strip()}', {dado[5]}),\n"
  comando = comando[:-2] + ';'

  return comando