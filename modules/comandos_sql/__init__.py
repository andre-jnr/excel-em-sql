from modules.menu import limpar_terminal

def ler_base_de_dados(caminho_arquivo):
  """Ler a base de dados em excel e converte para uma lista Python

  Args:
      caminho_arquivo (str): caminho do arquivo para transformar em lista Python

  Returns:
      list: retorna uma lista Python com todos os dados da tabela
  """
  import pandas as pd
  aba = listar_abas(caminho_arquivo)
  df = pd.read_excel(caminho_arquivo, sheet_name=aba)

  dados_lista = df.values.tolist()

  return dados_lista


def listar_abas(nome_arquivo, msg_erro=''):
  import pandas as pd
  limpar_terminal()
  arquivo = pd.ExcelFile(nome_arquivo)
  abas = arquivo.sheet_names

  print('Escolha uma aba da planilha: ')

  for index, aba in enumerate(abas):
    print(f'[{index + 1}] - {aba}')

  try:
    resposta = int(input('-> '))

    if int(resposta) > len(abas):
      listar_abas(nome_arquivo, msg_erro='Você escolheu uma opção que ultrapassa nossas opções!')

    return abas[int(resposta) - 1]

  except ValueError:
    listar_abas(nome_arquivo, msg_erro='Escolha uma opção válida, por gentileza!')



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