import pandas as pd
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import logging

# Configuração de logging para rastrear execução e erros
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Constantes
CAMINHO_PLANILHA = r"Z:\MARKETING\ENDOMARKETING\Aniversários\LISTA DE ANIVERSÁRIOS ATUALIZADA2.xlsx"
CAMINHO_BASE_ANIVERSARIOS = r"Z:\MARKETING\ENDOMARKETING\Aniversários\2025"
CAMINHO_PERFIL_CHROME = os.path.join(os.path.dirname(__file__), 'chrome_profile')
URL_WHATSAPP_GRUPO = "https://web.whatsapp.com/accept?code=EQDugAsxkaK5HsajvA4GyE"
MENSAGEM_PADRAO = "Muitas felicidades, feliz aniversário! 🎉🎊"

def carregar_planilha(caminho):
    """Carrega a planilha de aniversários e converte a coluna de datas."""
    try:
        df = pd.read_excel(caminho)
        df['Data de Nascimento'] = pd.to_datetime(df['Data de Nascimento'], format='%d/%m/%Y', errors='coerce')
        return df.dropna(subset=['Data de Nascimento'])  # Remove linhas com datas inválidas
    except Exception as e:
        logging.error(f"Erro ao ler a planilha: {e}")
        raise

def configurar_driver():
    """Configura o WebDriver do Chrome com perfil personalizado."""
    servico = Service(ChromeDriverManager().install())
    chrome_options = Options()
    chrome_options.add_argument(f"user-data-dir={CAMINHO_PERFIL_CHROME}")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    return webdriver.Chrome(service=servico, options=chrome_options)

def esperar_qr_code(driver):
    """Aguarda o usuário escanear o QR Code na primeira execução."""
    if not os.path.exists(os.path.join(CAMINHO_PERFIL_CHROME, "Default")):
        logging.info("Por favor, escaneie o QR Code no WhatsApp Web.")
        WebDriverWait(driver, 60).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Caixa de texto de mensagem"]'))
        )
        logging.info("QR Code escaneado com sucesso!")

def enviar_mensagem_whatsapp(nome, caminho_imagem, driver):
    """Envia mensagem com imagem no WhatsApp."""
    try:
        driver.get(URL_WHATSAPP_GRUPO)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        logging.info("Grupo carregado com sucesso!")

        # Abrir menu de anexos
        botao_mais = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="main"]/footer/div[1]/div/span/div/div[1]/div/button/span'))
        )
        botao_mais.click()
        time.sleep(1)

        # Selecionar "Fotos e vídeos"
        fotos_videos = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="app"]/div/span[5]/div/ul/div/div/div[2]'))
        )
        fotos_videos.click()
        time.sleep(1)

        # Upload da imagem
        campo_upload = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
        )
        campo_upload.send_keys(caminho_imagem)
        time.sleep(2)

        # Adicionar legenda
        legenda = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]'))
        )
        legenda.send_keys(f"{MENSAGEM_PADRAO}")  # Personaliza a mensagem com o nome
        time.sleep(1)

        # Enviar mensagem
        botao_enviar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
        )
        botao_enviar.click()
        time.sleep(10)  # Aguarda o envio
        logging.info(f"Mensagem enviada para {nome} com sucesso!")
    except Exception as e:
        logging.error(f"Erro ao enviar mensagem para {nome}: {e}")
        raise

def encontrar_pasta_mes(mes_atual):
    """Encontra a pasta correspondente ao mês atual."""
    try:
        for pasta in os.listdir(CAMINHO_BASE_ANIVERSARIOS):
            if pasta.startswith(f"{mes_atual:02d}."):
                return os.path.join(CAMINHO_BASE_ANIVERSARIOS, pasta)
        return None
    except Exception as e:
        logging.error(f"Erro ao buscar pasta do mês: {e}")
        raise

def main():
    """Função principal para executar o script."""
    hoje = datetime.now()
    dia_atual, mes_atual = hoje.day, hoje.month

    # Carregar planilha
    try:
        df = carregar_planilha(CAMINHO_PLANILHA)
    except Exception:
        return

    # Encontrar pasta do mês
    pasta_mes_atual = encontrar_pasta_mes(mes_atual)
    if not pasta_mes_atual:
        logging.warning(f"Pasta do mês atual ({mes_atual:02d}) não encontrada.")
        return
    logging.info(f"Pasta do mês atual encontrada: {pasta_mes_atual}")

    # Configurar driver
    driver = configurar_driver()
    try:
        driver.get("https://web.whatsapp.com")
        esperar_qr_code(driver)

        # Processar aniversariantes
        for _, row in df.iterrows():
            nome = row['Nome']
            data_nascimento = row['Data de Nascimento']

            if data_nascimento.day == dia_atual and data_nascimento.month == mes_atual:
                caminho_imagem = os.path.join(pasta_mes_atual, f"Aniversário {nome}.png")
                if os.path.exists(caminho_imagem):
                    logging.info(f"Arquivo encontrado para {nome}: {caminho_imagem}")
                    enviar_mensagem_whatsapp(nome, caminho_imagem, driver)
                else:
                    logging.warning(f"Arquivo não encontrado para {nome}: {caminho_imagem}")
    except Exception as e:
        logging.error(f"Erro na execução principal: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()