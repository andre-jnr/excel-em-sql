from modules import menu
from modules import comandos_sql as sql

def gerar_comando_sql(msg_erro=''):
  try:
    nome_arquivo = menu.localizar_arquivos(msg_erro=msg_erro)
    lst_dados = sql.ler_base_de_dados(nome_arquivo)
    lst_dados = sql.converter_string(lst_dados)

    # comando para criar a tabela temporária  com os dados da planilha
    comando_sql = sql.criar_comando_sql(lst_dados)

  except FileNotFoundError:
    gerar_comando_sql(msg_erro='Arquivo não encontrado')

  menu.copiar_para_area_trasnferencia(comando_sql)