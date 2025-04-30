# Importar bibliotecas necessárias
import requests
import paramiko
import pandas as pd
import os
import json
import psutil
import time
import datetime
import pytz
import smtplib
import logging
import subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Definir o caminho para o repositório "marchon"
MARCHON_FOLDER = os.path.join(os.getcwd(), 'marchon')  # Obtém o diretório atual

# Cria o diretório do log, se não existir
if not os.path.exists(MARCHON_FOLDER):
    os.makedirs(MARCHON_FOLDER)

# Configuração da API
LOG_FILE = os.path.join("log_envio_api.log")  # Caminho do log

# Configuração do log
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(message)s")

# Configurações do SFTP
SFTP_HOST = 'sftp.marchon.com.br'
SFTP_PORT = 2221
SFTP_USERNAME = 'CompreOculos'
SFTP_PASSWORD = '@CMPCLS$2023'
REMOTE_DIR = 'COMPREOCULOS/ESTOQUE'
FILE_TO_CHECK = 'estoque_disponivel_10.csv'

# Configuração da API
API_URL = 'https://api.bling.com.br/Api/v3/estoques'
LOG_FILE = os.path.join("log_envio_api.log")  # Caminho do log
TOKEN_FILE = os.path.join("token_novo.json")  # Caminho do token
BLING_AUTH_URL = "https://api.bling.com.br/Api/v3/oauth/token"
BASIC_AUTH = ("19f357c5eccab671fe86c94834befff9b30c3cea", "0cf843f8d474ebcb3f398df79077b161edbc6138bcd88ade942e1722303a")

# Configuração do log
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(message)s")

# Definição do ID do depósito
DEPOSITO_ID = 14888253052  # Substitua pelo ID do depósito desejado

# Ativa ou desativa o corte de estoque
ATIVAR_CORTE_ESTOQUE = False

TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token_novo.json")

def registrar_log(mensagem):
    """Registra mensagens no arquivo de log e imprime na saída."""
    logging.info(mensagem)
    print(mensagem)

def conectar_sftp():
    """Conecta ao servidor SFTP e retorna uma sessão."""
    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    try:
        print("Conectando ao servidor SFTP...")
        client.connect(SFTP_HOST, SFTP_PORT, SFTP_USERNAME, SFTP_PASSWORD)
        return client.open_sftp()
    except Exception as e:
        print(f"Erro ao conectar ao servidor SFTP: {e}")
        return None

def baixar_arquivo_sftp(sftp, remote_file_path, local_file_path):
    """Baixa um arquivo do SFTP para o diretório 'marchon'."""
    try:
        print(f"Baixando o arquivo {remote_file_path}...")
        start_time = time.time()
        sftp.get(remote_file_path, local_file_path)
        end_time = time.time()
        download_time = end_time - start_time
        print(f"Arquivo baixado para {local_file_path} em {download_time:.2f} segundos.")
    except Exception as e:
        print(f"Erro ao baixar o arquivo: {e}")

def ler_planilha_sftp(caminho_arquivo):
    """Lê e processa o arquivo CSV baixado do SFTP."""
    try:
        sftp_df = pd.read_csv(caminho_arquivo)
        print(f"Arquivo do SFTP carregado com {sftp_df.shape[0]} linhas.")
        sftp_df[['codigo_produto', 'balanco']] = sftp_df.iloc[:, 0].str.split(';', expand=True)
        sftp_df['balanco'] = sftp_df['balanco'].astype(float)
        return sftp_df[['codigo_produto', 'balanco']]
    except Exception as e:
        print(f"Erro ao ler a planilha do SFTP: {e}")
        return None

def ler_planilha_usuario():
    """Lê os dados da planilha estoque.xlsx da pasta do repositório 'marchon'."""
    caminho_planilha = os.path.join('Estoque_10.xlsx')  # Altere para o nome correto do arquivo

    if not os.path.exists(caminho_planilha):
        print("⚠ Erro: A planilha não pôde ser encontrada.")
        return None

    try:
        df = pd.read_excel(caminho_planilha)
        if df.shape[1] < 3:
            raise ValueError("A planilha deve conter pelo menos 3 colunas.")

        return pd.DataFrame({
            "id_usuario": df.iloc[:, 1].astype(str).str.strip(),
            "codigo_produto": df.iloc[:, 2].astype(str).str.strip()
        })
    except Exception as e:
        print(f"❌ Erro ao ler a planilha {caminho_planilha}: {e}")
        return None

def buscar_correspondencias(sftp_df, usuario_df):
    """Faz a correspondência entre os produtos do usuário e os do SFTP."""
    if sftp_df is None or usuario_df is None:
        print("Erro: Arquivos de entrada não carregados corretamente.")
        return pd.DataFrame()

    # Merge dos DataFrames
    resultado = usuario_df.merge(sftp_df, on="codigo_produto", how="left")

    # Aplicar corte de estoque se ativado
    if ATIVAR_CORTE_ESTOQUE:
        print("🔧 Corte de estoque ativado: Subtraindo 10 unidades de balanços acima de 10.")
        resultado['balanco'] = resultado['balanco'].apply(
            lambda x: max(x - 10, 0) if pd.notna(x) and x > 10 else x
        )
    else:
        print("🚫 Corte de estoque desativado.")

    # Filtrar apenas os produtos com balanço maior que zero
    resultado = resultado[resultado['balanco'] > 0]

    # Ordenar os resultados pelo 'balanco' em ordem decrescente
    resultado = resultado.sort_values(by='balanco', ascending=False)

    return resultado

def commit_e_push_resultados():
    """Faz commit e push do arquivo resultado_correspondencias_10.xlsx para o repositório"""
    try:
        # Configurar identidade do Git
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)

        # Adiciona o arquivo e faz commit
        subprocess.run(["git", "add", "resultado_correspondencias_10.xlsx"], check=True)
        subprocess.run(["git", "commit", "-m", "Atualizando resultado_correspondencias_10.xlsx"], check=True)
        subprocess.run(["git", "push"], check=True)
        
        print("✅ Resultados commitados e enviados para o repositório!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao tentar fazer commit e push: {e}")
    

def log_envio(mensagem):
    """Registra mensagens de envio no log."""
    registrar_log(mensagem)

def enviar_dados_api(resultado_df, deposito_id):
    """Envia os dados processados para a API do Bling."""
    if resultado_df.empty:
        print("Nenhum dado para enviar à API.")
        return

    token = obter_access_token()  # 🔥 Agora o token é gerado automaticamente!
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    session = requests.Session()
    session.headers.update(headers)

    log_envio("\n🔍 Iniciando envio de dados para a API...\n")
    contador_envios = 0
    total_bytes_enviados = 0
    start_time = time.time()

    for _, row in resultado_df.iterrows():
        if pd.notna(row["balanco"]) and pd.notna(row["id_usuario"]) and row["balanco"] > 0:
            payload = {
                "produto": {
                    "id": int(row["id_usuario"]),
                    "codigo": row["codigo_produto"]
                },
                "deposito": {
                    "id": deposito_id
                },
                "operacao": "B",
                "preco": 100,
                "custo": 10,
                "quantidade": row["balanco"],
                "observacoes": "Atualização de estoque via script"
            }
            try:
                send_start_time = time.time()
                response = session.post(API_URL, json=payload)
                send_end_time = time.time()
                total_bytes_enviados += len(json.dumps(payload).encode('utf-8'))

                log_msg = f"\n📦 Enviado para API:\n{json.dumps(payload, indent=2)}"

                if response.status_code in [200, 201]:
                    log_envio(f"✔ Sucesso [{response.status_code}]: Produto {row['codigo_produto']} atualizado na API.{log_msg}")
                    contador_envios += 1
                else:
                    log_envio(f"❌ Erro [{response.status_code}]: {response.text}{log_msg}")

                response_time = send_end_time - send_start_time
                log_envio(f"⏱ Tempo de resposta do servidor para {row['codigo_produto']}: {response_time:.2f} segundos")
                time.sleep(0.4)  # 💤 Aguarda para não exceder o limite da API

            except Exception as e:
                log_envio(f"❌ Erro ao enviar {row['codigo_produto']}: {e}")

        else:
            motivo = []
            if pd.isna(row["balanco"]):
                motivo.append("balanço vazio")
            elif row["balanco"] <= 0:
                motivo.append("balanço zero ou negativo")
            if pd.isna(row["id_usuario"]):
                motivo.append("id_usuario vazio")
            
            log_envio(f"⚠ Produto {row['codigo_produto']} ignorado. Motivo(s): {', '.join(motivo)}")

    end_time = time.time()
    total_time = end_time - start_time
    upload_speed = total_bytes_enviados / total_time if total_time > 0 else 0
    cpu_usage = psutil.cpu_percent(interval=1)

    return contador_envios

TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token_novo.json")

def baixar_token():
    """Lê o token_novo.json armazenado no diretório 'marchon'."""
    if not os.path.exists(TOKEN_FILE):
        print("⚠ Arquivo de token não encontrado.")
        return None

    try:
        with open(TOKEN_FILE, "r") as file:
            return json.load(file)
    except Exception as e:
        print(f"❌ Erro ao ler token: {e}")
        return None

def salvar_token_novo(token_data):
    """Salva o token atualizado no arquivo token_novo.json"""
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(token_data, f, indent=4)
    
    print(f"✅ Token atualizado e salvo em: {TOKEN_FILE}")

def commit_e_push_token():
    """Faz commit e push do token atualizado para o repositório"""
    try:
        subprocess.run(["git", "add", "token_novo.json"], check=True)
        subprocess.run(["git", "commit", "-m", "🔄 Atualizando token_novo.json"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("✅ Token atualizado e enviado para o repositório!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao tentar fazer commit e push: {e}")

def salvar_resultados(resultados):
    """Salva os resultados em um arquivo e faz commit."""
    caminho_arquivo = os.path.join(os.path.dirname(__file__), "resultado_correspondencias_10.xlsx")
    try:
        resultados.to_excel(caminho_arquivo, index=False)
        print(f"✅ Resultados salvos em: {caminho_arquivo}")

        # Adiciona o arquivo ao índice
        subprocess.run(["git", "add", caminho_arquivo], check=True)

        # Verifica se há alterações para comitar
        if subprocess.run(["git", "diff-index", "--quiet", "HEAD"], check=False).returncode != 0:
            subprocess.run(["git", "commit", "-m", "Atualizando resultado_correspondencias_10.xlsx"], check=True)
            subprocess.run(["git", "push"], check=True)
            print("✅ Resultados commitados e enviados para o repositório!")
        else:
            print("⚠️ Nenhuma alteração para comitar.")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao salvar ou enviar os resultados: {e}")

def obter_refresh_token():
    """Obtém o refresh_token do arquivo JSON baixado."""
    data = baixar_token()
    return data.get("refresh_token") if data else None

def gerar_novo_token():
    """Gera um novo access_token e salva no diretório 'marchon'."""
    refresh_token = obter_refresh_token()
    if not refresh_token:
        raise ValueError("⚠ Refresh token não encontrado.")

    payload = {
        "grant_type": "refresh_token",
        "refresh_token": refresh_token
    }

    response = requests.post(BLING_AUTH_URL, data=payload, auth=BASIC_AUTH)

    if response.status_code in [200, 201]:
        novo_token = response.json()
        salvar_token_novo(novo_token)  # Corrigido nome da função
        commit_e_push_token()  # Agora faz commit e push do novo token automaticamente
        print("✅ Novo access_token gerado com sucesso!")
        return novo_token["access_token"]
    else:
        raise Exception(f"❌ Erro ao gerar novo token: {response.status_code} - {response.text}")

def obter_access_token():
    """Sempre gera um novo access_token antes de cada execução."""
    return gerar_novo_token()


def main():
    sftp = conectar_sftp()
    if not sftp:
        print("Conexão com o SFTP falhou. Finalizando o script.")
        return

    # Caminho local para salvar o arquivo baixado
    local_file_path = os.path.join(MARCHON_FOLDER, FILE_TO_CHECK)
    remote_file_path = f"{REMOTE_DIR}/{FILE_TO_CHECK}"

    # Baixar o arquivo do SFTP
    baixar_arquivo_sftp(sftp, remote_file_path, local_file_path)
    sftp.close()

    # Ler o arquivo baixado do SFTP
    sftp_df = ler_planilha_sftp(local_file_path)
    usuario_df = ler_planilha_usuario()

    if sftp_df is None or usuario_df is None:
        return

    # Buscar correspondências entre os dados do SFTP e do usuário
    resultados = buscar_correspondencias(sftp_df, usuario_df)

    # Salvar resultados no repositório
    salvar_resultados(resultados)
    # Fazer commit e push dos resultados
    commit_e_push_resultados()
    # Enviar dados para a API do Bling
    enviar_dados_api(resultados, DEPOSITO_ID)

    # Calcular soma do estoque e contagem de IDs com estoque maior ou igual a 1
    soma_estoque = resultados['balanco'].sum()
    contagem_ids_diferente_zero = resultados[resultados['balanco'] != 0].shape[0]

    # Verificar se o corte de estoque está ativado
    status_corte_estoque = "ativado" if ATIVAR_CORTE_ESTOQUE else "desativado"

    # Mensagem do e-mail com resumo do estoque
    mensagem_email = (
    f"📦 Produtos enviados para a API (balanço ≠ 0): {contagem_ids_diferente_zero}\n"
    f"🧮 Soma total do estoque (balanço): {soma_estoque}\n"
    f"🔒 Corte de Estoque: {status_corte_estoque}\n\n"
    "📎 Segue em anexo o relatório atualizado da Marchon."
    )

     # Enviar dados para a API do Bling
    sucesso = enviar_dados_api(resultados, DEPOSITO_ID)
    print(f"\n✅ {sucesso} produtos foram enviados para a API com sucesso.")

    # Enviar o e-mail com o relatório após o envio dos dados
    enviar_email_com_anexo(
        "victor@compreoculos.com.br",
        "Relatório de Estoque Marchon 10",
        "Segue em anexo o relatório atualizado da Marchon10.",
        os.path.join("resultado_correspondencias_10.xlsx")  # O arquivo que você gerou anteriormente
    )

def enviar_email_com_anexo(destinatario, assunto, mensagem, anexo_path):
    """Envia um e-mail com um arquivo anexo."""
    remetente = "victor@compreoculos.com.br"  # Altere para seu e-mail
    senha = "hhhlzwcyucmjdqqb"  # Use um App Password ou método seguro para armazenar credenciais

    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = assunto

    msg.attach(MIMEText(mensagem, "plain"))

    # Anexar arquivo
    if os.path.exists(anexo_path):
        with open(anexo_path, "rb") as anexo:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(anexo.read())
            encoders.encode_base64(parte)
            parte.add_header("Content-Disposition", f"attachment; filename={os.path.basename(anexo_path)}")
            msg.attach(parte)
    else:
        print(f"⚠ Arquivo {anexo_path} não encontrado para anexo.")

    try:
        servidor = smtplib.SMTP("smtp.gmail.com", 587)
        servidor.starttls()
        servidor.login(remetente, senha)
        servidor.sendmail(remetente, destinatario, msg.as_string())
        servidor.quit()
        print(f"📧 E-mail enviado com sucesso para {destinatario}")
    except Exception as e:
        print(f"❌ Erro ao enviar e-mail: {e}")


if __name__ == "__main__":
    main()
