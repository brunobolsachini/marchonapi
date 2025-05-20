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

# Ativa ou desativa o corte de estoque
ATIVAR_CORTE_ESTOQUE = False

# Valor mínimo de estoque. Se o estoque for menor que esse valor, será considerado 0.
CORTE_ESTOQUE_MINIMO = 0  # Modifique conforme necessário

# Definição do ID do depósito
DEPOSITO_ID = 14886798357

# Definir o caminho para o repositório "marchon"
MARCHON_FOLDER = os.path.join(os.getcwd(), 'marchon')

# Cria o diretório do log, se não existir
if not os.path.exists(MARCHON_FOLDER):
    os.makedirs(MARCHON_FOLDER)

# Caminho do log
LOG_FILE = os.path.join("log_envio_api.log")
logging.basicConfig(filename=LOG_FILE, level=logging.INFO, format="%(asctime)s - %(message)s")

# Configurações do SFTP
SFTP_HOST = 'sftp.marchon.com.br'
SFTP_PORT = 2221
SFTP_USERNAME = 'CompreOculos'
SFTP_PASSWORD = '@CMPCLS$2025'
REMOTE_DIR = 'COMPREOCULOS/ESTOQUE'
FILE_TO_CHECK = 'estoque_disponivel_10.csv'

# Configuração da API
API_URL = 'https://api.bling.com.br/Api/v3/estoques'
TOKEN_FILE = os.path.join(os.path.dirname(__file__), "token_novo.json")
BLING_AUTH_URL = "https://api.bling.com.br/Api/v3/oauth/token"
BASIC_AUTH = (
    "19f357c5eccab671fe86c94834befff9b30c3cea",
    "0cf843f8d474ebcb3f398df79077b161edbc6138bcd88ade942e1722303a"
)

def registrar_log(mensagem):
    logging.info(mensagem)
    print(mensagem)

def conectar_sftp():
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
    try:
        print(f"Baixando o arquivo {remote_file_path}...")
        start_time = time.time()
        sftp.get(remote_file_path, local_file_path)
        end_time = time.time()
        print(f"Arquivo baixado para {local_file_path} em {end_time - start_time:.2f} segundos.")
    except Exception as e:
        print(f"Erro ao baixar o arquivo: {e}")

def ler_planilha_sftp(caminho_arquivo):
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
    caminho_planilha = os.path.join('Estoque.xlsx')
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
    if sftp_df is None or usuario_df is None:
        print("Erro: Arquivos de entrada não carregados corretamente.")
        return pd.DataFrame()

    resultado = usuario_df.merge(sftp_df, on="codigo_produto", how="left")

    if ATIVAR_CORTE_ESTOQUE:
        print(f"🔧 Corte de estoque ativado: Estoques abaixo de {CORTE_ESTOQUE_MINIMO} serão zerados.")
        resultado['balanco'] = resultado['balanco'].apply(
            lambda x: 0 if pd.notna(x) and x < CORTE_ESTOQUE_MINIMO else x
        )
    else:
        print("🚫 Corte de estoque desativado.")

    resultado = resultado.sort_values(by='balanco', ascending=False)
    return resultado

def commit_e_push_resultados():
    try:
        subprocess.run(["git", "config", "--global", "user.name", "github-actions[bot]"], check=True)
        subprocess.run(["git", "config", "--global", "user.email", "github-actions[bot]@users.noreply.github.com"], check=True)
        subprocess.run(["git", "add", "resultado_correspondencias.xlsx"], check=True)
        subprocess.run(["git", "commit", "-m", "Atualizando resultado_correspondencias.xlsx"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("✅ Resultados commitados e enviados para o repositório!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao tentar fazer commit e push: {e}")

def log_envio(mensagem):
    registrar_log(mensagem)

def enviar_dados_api(resultado_df, deposito_id):
    if resultado_df.empty:
        print("Nenhum dado para enviar à API.")
        return

    token = obter_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    session = requests.Session()
    session.headers.update(headers)

    log_envio("\n🔍 Iniciando envio de dados para a API...\n")
    total_bytes_enviados = 0
    start_time = time.time()

    for _, row in resultado_df.iterrows():
        if pd.notna(row["balanco"]) and pd.notna(row["id_usuario"]):
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
                else:
                    log_envio(f"❌ Erro [{response.status_code}]: {response.text}{log_msg}")

                response_time = send_end_time - send_start_time
                log_envio(f"⏱ Tempo de resposta: {response_time:.2f} segundos")
                time.sleep(0.4)
            except Exception as e:
                log_envio(f"❌ Erro ao enviar {row['codigo_produto']}: {e}")
        else:
            motivos = []
            if pd.isna(row["balanco"]): motivos.append("balanço vazio")
            if pd.isna(row["id_usuario"]): motivos.append("id_usuario vazio")
            log_envio(f"⚠ Produto {row['codigo_produto']} ignorado. Motivo(s): {', '.join(motivos)}")

    print(f"⏱ Envio concluído em {time.time() - start_time:.2f} segundos.")

def baixar_token():
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
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        json.dump(token_data, f, indent=4)
    print(f"✅ Token atualizado e salvo em: {TOKEN_FILE}")

def commit_e_push_token():
    try:
        subprocess.run(["git", "add", "token_novo.json"], check=True)
        subprocess.run(["git", "commit", "-m", "🔄 Atualizando token_novo.json"], check=True)
        subprocess.run(["git", "push"], check=True)
        print("✅ Token atualizado e enviado para o repositório!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao tentar fazer commit e push: {e}")

def salvar_resultados(resultados):
    caminho_resultados = os.path.join(os.path.dirname(__file__), "resultado_correspondencias.xlsx")
    resultados.to_excel(caminho_resultados, index=False)
    print(f"✅ Resultados salvos em: {caminho_resultados}")
    subprocess.run(["git", "add", caminho_resultados])
    subprocess.run(["git", "commit", "-m", "Atualizando resultado_correspondencias.xlsx"])
    subprocess.run(["git", "push"])

def obter_refresh_token():
    data = baixar_token()
    return data.get("refresh_token") if data else None

def gerar_novo_token():
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
        salvar_token_novo(novo_token)
        commit_e_push_token()
        print("✅ Novo access_token gerado com sucesso!")
        return novo_token["access_token"]
    else:
        raise Exception(f"❌ Erro ao gerar novo token: {response.status_code} - {response.text}")

def obter_access_token():
    return gerar_novo_token()

def enviar_email_com_anexo(destinatario, assunto, mensagem, anexo_path):
    remetente = "bruno@compreoculos.com.br"
    senha = "diwihenjpuuoxnwc"
    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = destinatario
    msg["Subject"] = assunto
    msg.attach(MIMEText(mensagem, "plain"))

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

def main():
    sftp = conectar_sftp()
    if not sftp:
        print("Conexão com o SFTP falhou. Finalizando o script.")
        return

    local_file_path = os.path.join(MARCHON_FOLDER, FILE_TO_CHECK)
    remote_file_path = f"{REMOTE_DIR}/{FILE_TO_CHECK}"
    baixar_arquivo_sftp(sftp, remote_file_path, local_file_path)
    sftp.close()

    sftp_df = ler_planilha_sftp(local_file_path)
    usuario_df = ler_planilha_usuario()

    if sftp_df is None or usuario_df is None:
        return

    resultados = buscar_correspondencias(sftp_df, usuario_df)
    salvar_resultados(resultados)
    commit_e_push_resultados()
    enviar_dados_api(resultados, DEPOSITO_ID)

    soma_estoque = resultados['balanco'].sum()
    contagem_ids_diferente_zero = resultados[resultados['balanco'] != 0].shape[0]
    status_corte_estoque = "ativado" if ATIVAR_CORTE_ESTOQUE else "desativado"

    mensagem_email = (
        f"📦 Produtos enviados para a API (balanço ≠ 0): {contagem_ids_diferente_zero}\n"
        f"🧮 Soma total do estoque (balanço): {soma_estoque}\n"
        f"🔒 Corte de Estoque: {status_corte_estoque}\n\n"
        "📎 Segue em anexo o relatório atualizado da Marchon."
    )

    enviar_email_com_anexo(
        "bruno@compreoculos.com.br",
        "Relatório de Estoque Marchon 10",
        mensagem_email,
        os.path.join(os.path.dirname(__file__), "resultado_correspondencias_Marchon10.xlsx")
    )

    print(f"\n✅ Total de IDs processados e subidos: {resultados.shape[0]}")

if __name__ == "__main__":
    main()
