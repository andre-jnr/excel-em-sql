def titulo():
  """Criar o titulo do programa no terminal
  """
  titulo = 'CRIADOR DE TABELAS TEMPORÁRIAS'
  print(titulo)
  print('CONVERTA DE EXCEL PARA SQL')

  print('~' * len(titulo))


def limpar_terminal():
  """Limpa o terminal (windows)
  """
  from os import system
  system('cls')

def localizar_arquivos(msg_erro=''):
  """Para localizar o caminho dos arquivos excel

  Args:
      msg_erro (str, optional): mensagem de erro para o usuário conseguir identificar seu erro. Defaults to ''.

  Returns:
      function: função para listar os arquivos no caminho especificado
      or
      str: mensagem de erro
  """
  limpar_terminal()
  titulo()
  print('O arquivo excel está localizado aqui?')
  print('[mesmo local (pasta) do script Python]')
  print('')
  print('[1] - SIM')
  print('[2] - NÃO')

  if len(msg_erro) > 0:
     print(f'{msg_erro}')

  resposta = input(str('-> '))
  print('')

  if not resposta in ['1', '2']:
    localizar_arquivos(msg_erro='Por favor, digite uma opção válida!')

  match resposta:
     case '1':
        arquivos_excel = encontrar_arquivos_excel()
        return listar_arquivos(arquivos_excel)
     case '2':
        return input('Digite o caminho do arquivo Excel: ')
     

def encontrar_arquivos_excel():
  """Encontra todos os arquivos com as extensões .xlsx e .xls (arquivos excel)

  Returns:
      lista: lista com o nome de todos os arquivos excel
  """
  from os import getcwd, listdir

  pasta_atual = getcwd()
  arquivos_excel = [f for f in listdir(pasta_atual) if f.endswith(('.xlsx', '.xls'))]
    
  if arquivos_excel:
    return arquivos_excel
  else:
    print("Nenhum arquivo Excel encontrado na pasta atual.")


def listar_arquivos(lista, msg_erro=''):
   """Lista todos os arquivos excel em forma de menu, onde o usuário escolhe um arquivo.

   Args:
       lista (lista): lista com o nome de todos os arquivos para serem escolhidos.
       msg_erro (str, optional): mensagem de erro para auxiliar o usuário a saber o que ele está digitando de errado. Defaults to ''.
   """
   limpar_terminal()
   print('Escolha um arquivo Excel:')
   for i, arquivo in enumerate(lista, start=1):
      print(f'[{i}] - {arquivo}')

   try:
      if len(msg_erro) > 0:
         print(f'{msg_erro}')
      resposta = int(input('-> '))
    
   except ValueError:
      listar_arquivos(lista, msg_erro='Por favor, digite um número inteiro.')

   if resposta > len(lista):
      listar_arquivos(lista, msg_erro='Você digitou um número que ultrapassa o número de escolhas!')

   return lista[resposta - 1]


def copiar_para_area_trasnferencia(txt):
   """Para copiar qualquer texto para área de transferência

   Args:
       txt (str): texto para ser copiado para área de transferência
   """
   import pyperclip
   limpar_terminal()
   pyperclip.copy(txt)
   print('Comando copiado para área de transferencia')
   print()
   tela_final(txt)

def tela_final(txt, msg_erro=''):

   print('O que deseja fazer agora?')
   print('[1] - Voltar para tela inicial')
   print('[2] - Sair do programa')
   print('[3] - Visualizar comando copiado')

   if len(msg_erro) > 0:
      print(msg_erro)

   respota = str(input('-> '))

   match respota:
      case '1':
         import subprocess
         subprocess.run(["python", "main.py"])
      case '2':
         limpar_terminal()
         exit()
      case '3':
         limpar_terminal()
         print(txt)
         print()
         tela_final(txt)
      case _:
         limpar_terminal()
         tela_final(txt, msg_erro='Comando inválido, digite alguma opção válida!')


def criar_tabela():
   """Cria uma nova tabela temporária no banco de dados
   """
   comando_sql ="""
CREATE TABLE temp_sped_estoque (
    codigo INT,
    descricao VARCHAR(255) NOT NULL,
    grupo VARCHAR(100),
    estoque DECIMAL(10,2) DEFAULT 0,
    unidade VARCHAR(10) NOT NULL,
    custo DECIMAL(10,2) NOT NULL,
    custo_total DECIMAL(12,2) GENERATED ALWAYS AS (estoque * custo) STORED
);   
"""
   copiar_para_area_trasnferencia(comando_sql)
   tela_final(comando_sql)


def tela_inicial():
   """Tela inicial do programa
   """
   from modules import processo
   limpar_terminal()
   titulo()
   print('Escolha uma opção:')
   print('[1] - Comando para criar tabela')
   print('Obs: comando para criar tabela é baseada nas outras planilhas utilizadas para criar outras tabelas')
   print('[2] - Comando para inserir dados na tabela')
   print('Obs: utilizamos uma planilha para alimentar os dados da tabela')
   resposta = str(input('-> '))

   match resposta:
      case '1':
         criar_tabela()
      case '2':
         processo.gerar_comando_sql()
      case _:
         limpar_terminal()
         tela_inicial()
