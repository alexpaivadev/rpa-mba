import requests
import datetime
import json
import os
import pandas as pd # Importa a biblioteca pandas
import openpyxl # Importa a biblioteca openpyxl para salvar arquivos Excel
# --- Configurações da API ---
API_URL = "https://api.powercrm.com.br/api/report/db"
# ATENÇÃO: Substitua pelo seu token real se necessário, ou considere usar variáveis de ambiente.
BEARER_TOKEN = "NBV3XdGoMDr4uWDlpP7safKsO6U2rQCFyCxweJVtMc2kTt39K5lwZO8Ec2u1EvYQ2gcKmZxT1AeVYtX1s3iVxgZjAgUHqRIv4FEyswMfELwZTE1GXkTLmWQQPgd2x3LKEjAVUdCKAxWxndwC5Q5p9JfNXIC3qRGTcFZYLZaP23LtsoIcr32jgsZSJEKG0SIbeVcZXeMS0wjCRbCpUue8SM76"

# Headers da requisição
HEADERS = {
    'accept': 'application/json', # A API ainda deve retornar JSON
    'authorization': f'Bearer {BEARER_TOKEN}',
    'content-type': 'application/json'
}

# --- Dados (payload) para a requisição POST ---
# ATENÇÃO: As datas "from" e "to" estão fixas conforme o seu exemplo cURL original.
# Se você precisa que essas datas sejam dinâmicas (ex: "últimas 2 horas"),
# elas precisarão ser calculadas antes de montar o payload.
PAYLOAD = {
  "from": "2025-04-01", # Exemplo de data, ajuste conforme necessário
  "stringFilterTypeDate": 1,
  "to": "2025-05-31"   # Exemplo de data, ajuste conforme necessário
}

# --- Configurações do Arquivo de Saída ---
NOME_BASE_ARQUIVO = "CRM" # Nome base do arquivo
FORMATO_ARQUIVO = "xlsx"  # Formato do arquivo
# Altere para a pasta onde você quer salvar os relatórios.
# Exemplo Windows: "C:/Usuarios/SeuUsuario/Documentos/RelatoriosPowerCRM"
# Exemplo Linux/macOS: "/home/seuusuario/relatorios_powercrm"
PASTA_SALVAR = "./relatorios_powercrm"  # Salva em uma subpasta 'relatorios_powercrm' no diretório atual

def extrair_e_salvar_relatorio_powercrm():
    """
    Extrai o relatório da API PowerCRM usando POST e salva em um arquivo Excel
    com o nome fixo CRM.xlsx.
    """
    print(f"Iniciando extração do relatório da PowerCRM para o arquivo {NOME_BASE_ARQUIVO}.{FORMATO_ARQUIVO}: {API_URL}")

    # Cria a pasta para salvar os relatórios, se não existir
    if not os.path.exists(PASTA_SALVAR):
        try:
            os.makedirs(PASTA_SALVAR)
            print(f"Pasta '{PASTA_SALVAR}' criada com sucesso.")
        except OSError as e:
            print(f"Erro ao criar pasta '{PASTA_SALVAR}': {e}")
            return

    # Payload (mantendo a lógica de datas fixas ou dinâmicas como no script anterior)
    # Se precisar de datas dinâmicas (ex: últimas 2 horas):
    # Para usar a data atual e duas horas atrás, você pode fazer:
    # agora = datetime.datetime.now()
    # duas_horas_atras = agora - datetime.timedelta(hours=2)
    # payload_dinamico = {
    #     "from": duas_horas_atras.strftime("%Y-%m-%d"), # Ajuste o formato se a API esperar hora também
    #     "stringFilterTypeDate": 1, # Verifique se este campo é relevante para filtros dinâmicos
    #     "to": agora.strftime("%Y-%m-%d") # Ajuste o formato se a API esperar hora também
    # }
    # dados_para_enviar = payload_dinamico
    # print(f"Payload dinâmico: {dados_para_enviar}")
    
    # Usando o PAYLOAD fixo por padrão (ajuste as datas "from" e "to" em PAYLOAD conforme sua necessidade)
    dados_para_enviar = PAYLOAD
    print(f"Payload: {dados_para_enviar}")

    try:
        # Fazendo a requisição POST para a API
        response = requests.post(API_URL, headers=HEADERS, json=dados_para_enviar, timeout=60)
        response.raise_for_status()  # Lança uma exceção para respostas HTTP de erro

        try:
            dados_json = response.json()
            
            if isinstance(dados_json, list):
                 df = pd.DataFrame(dados_json)
            elif isinstance(dados_json, dict):
                 df = pd.json_normalize(dados_json)
            else:
                print("Formato JSON não esperado para conversão direta para tabela. Verifique a resposta da API.")
                timestamp_erro = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                nome_arquivo_erro_json = os.path.join(PASTA_SALVAR, f"erro_json_inesperado_{timestamp_erro}.json")
                with open(nome_arquivo_erro_json, 'w', encoding='utf-8') as f_err:
                    json.dump(dados_json, f_err, ensure_ascii=False, indent=4)
                print(f"JSON bruto salvo em: {nome_arquivo_erro_json}")
                return

            if df.empty and dados_json: # Se o DataFrame está vazio mas havia dados JSON
                print("O DataFrame resultante está vazio, mas a API retornou dados. Verifique a estrutura do JSON.")
                print("Pode ser necessário ajustar como o JSON é passado para o DataFrame (ex: df = pd.DataFrame(dados_json['chave_dos_dados'])).")
                # Salva o JSON original para depuração
                timestamp_vazio = datetime.datetime.now().strftime("%Y%m%d_%H%M%S") # Timestamp para o arquivo de log/erro
                nome_arquivo_vazio_json = os.path.join(PASTA_SALVAR, f"{NOME_BASE_ARQUIVO}_dados_json_df_vazio_{timestamp_vazio}.json")
                with open(nome_arquivo_vazio_json, 'w', encoding='utf-8') as f_vazio:
                    json.dump(dados_json, f_vazio, ensure_ascii=False, indent=4)
                print(f"JSON original (que resultou em DataFrame vazio) salvo em: {nome_arquivo_vazio_json}")
                # Continua para tentar salvar um Excel vazio, ou você pode optar por `return` aqui.
            elif df.empty: # Se o DataFrame está vazio e não havia dados JSON (ou dados_json era None/vazio)
                print("A API não retornou dados ou os dados resultaram em um DataFrame vazio. O arquivo Excel não será gerado ou estará vazio.")
                # Não há necessidade de salvar um JSON se ele já estava vazio/nulo
                # Se mesmo assim quiser um arquivo Excel vazio:
                # df.to_excel(nome_arquivo_completo, index=False, engine='openpyxl')
                # print(f"Arquivo Excel vazio '{nome_arquivo_completo}' gerado.")
                return # Decide se quer sair ou gerar um Excel vazio


            # Gerando o nome do arquivo Excel fixo (sem timestamp)
            nome_arquivo_completo = os.path.join(PASTA_SALVAR, f"{NOME_BASE_ARQUIVO}.{FORMATO_ARQUIVO}")

            # Salvando o DataFrame para um arquivo Excel (sobrescreverá se já existir)
            df.to_excel(nome_arquivo_completo, index=False, engine='openpyxl')
            print(f"Relatório salvo com sucesso em: {nome_arquivo_completo}")
            print("ATENÇÃO: Este arquivo será sobrescrito na próxima execução do script.")

        except json.JSONDecodeError:
            print("A resposta da API não é um JSON válido. Não é possível converter para Excel.")
            print(f"Resposta recebida (texto): {response.text[:500]}...")
            timestamp_erro = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo_erro_txt = os.path.join(PASTA_SALVAR, f"{NOME_BASE_ARQUIVO}_erro_nao_json_{timestamp_erro}.txt")
            with open(nome_arquivo_erro_txt, 'w', encoding='utf-8') as f_txt_err:
                f_txt_err.write(response.text)
            print(f"Resposta bruta salva em: {nome_arquivo_erro_txt}")
            
        except AttributeError as e:
            print(f"Erro ao converter JSON para DataFrame pandas: {e}")
            print("Isso pode ocorrer se a estrutura do JSON não for uma lista de dicionários ou um dicionário simples, ou se 'dados_json' for None.")
            if 'dados_json' in locals():
                 print(f"Dados JSON recebidos: {dados_json}")
                 timestamp_erro = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                 nome_arquivo_erro_json = os.path.join(PASTA_SALVAR, f"erro_conversao_pandas_{timestamp_erro}.json")
                 with open(nome_arquivo_erro_json, 'w', encoding='utf-8') as f_err:
                    json.dump(dados_json, f_err, ensure_ascii=False, indent=4)
                 print(f"JSON problemático salvo em: {nome_arquivo_erro_json}")


    except requests.exceptions.HTTPError as errh:
        print(f"Erro HTTP: {errh}")
        print(f"Detalhes da resposta: {response.text if 'response' in locals() else 'N/A'}")
    except requests.exceptions.ConnectionError as errc:
        print(f"Erro de Conexão: {errc}")
    except requests.exceptions.Timeout as errt:
        print(f"Erro de Timeout: {errt}")
    except requests.exceptions.RequestException as err:
        print(f"Erro na Requisição: {err}")
    except Exception as e:
        print(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    print("--- ATENÇÃO ---")
    print("Este script usa um Bearer Token diretamente no código. Para maior segurança, considere usar variáveis de ambiente.")
    print("As datas no payload estão fixas. Se precisar de datas dinâmicas para extrações periódicas, ajuste o PAYLOAD.")
    print(f"O script tentará salvar o relatório como '{os.path.join(PASTA_SALVAR, NOME_BASE_ARQUIVO + '.' + FORMATO_ARQUIVO)}'.")
    print("Este arquivo será SOBRESCRITO a cada execução.")
    print("Certifique-se de ter as bibliotecas 'pandas' e 'openpyxl' instaladas (`pip install pandas openpyxl`).")
    print("-----------------")
    
    extrair_e_salvar_relatorio_powercrm()