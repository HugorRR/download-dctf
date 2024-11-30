import os
import sys
import time
import logging
import subprocess
import pandas as pd
from pathlib import Path

import pyautogui
import streamlit as st
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

# Configurações globais
competencia = ""
data_inicial = ""
data_final = ""

# Configuração de logging
logging.basicConfig(
    filename='AUTOMACAO-DCTF.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
pyautogui.PAUSE = 1

# Definição de caminhos
BASE_DIR = Path(__file__).parent
PASTA_COMPETENCIA = BASE_DIR / "Competencias executadas"
IMAGEM_DIR = BASE_DIR / "img"
PLANILHA = BASE_DIR / "Clientes.xlsx"

pasta_competencia = os.path.join(PASTA_COMPETENCIA, competencia)
if not os.path.exists(pasta_competencia):
    os.mkdir(pasta_competencia)

def renomear_arquivo_recente(codigo, competencia):
    try:
        arquivos = list(PASTA_COMPETENCIA.glob("*"))
        arquivo_recente = max(arquivos, key=os.path.getctime)
        novo_nome = PASTA_COMPETENCIA / f"{codigo} DARFWEB {competencia}.pdf"
        arquivo_recente.rename(novo_nome)
        logging.info(f"Arquivo renomeado para: {novo_nome}")
    except Exception as e:
        logging.error(f"Erro ao renomear o arquivo: {e}")

def reconhecimento(imagens_referencia, tempo_limite, confidence=1.0):
    """Espera uma das imagens de referência aparecer na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                logging.info(f"Imagem encontrada: {imagem_referencia}")
                return True
            logging.info(f"Imagem não encontrada: {imagem_referencia}")
        time.sleep(1)
    logging.info("Nenhuma imagem encontrada dentro do tempo limite.")
    return False

def clique(imagens_referencia, tempo_limite, confidence=1.0):
    """Tenta clicar em uma das imagens de referência na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                pyautogui.click(posicao)
                logging.info(f"Clique na imagem: {imagem_referencia}")
                return True
            logging.info(f"Imagem não reconhecida: {imagem_referencia}")
        time.sleep(1)
    return False

def clique2(imagens_referencia, tempo_limite, confidence=1.0, ocorrencia=1):
    """Tenta clicar em uma das imagens de referência na tela dentro de um tempo limite."""
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicoes = list(pyautogui.locateAllOnScreen(imagem_referencia, confidence=confidence))
            logging.info(f"Número de ocorrências da imagem '{imagem_referencia}' encontradas na tela: {len(posicoes)}")
            if len(posicoes) >= ocorrencia:
                pyautogui.click(posicoes[ocorrencia-1])
                logging.info(f"Clique na imagem: {imagem_referencia}")
                return True
        logging.info(f"Imagem não reconhecida: {imagem_referencia}")
        time.sleep(1)
    return False

def ler_planilha():
    df = pd.read_excel(PLANILHA)
    cnpjs = df['CNPJ'].astype(str).tolist()
    codigos = df['COD'].astype(str).tolist()
    df['STATUS'] = ''
    return cnpjs, codigos, df

def configurar_driver():
    """Configura o driver do Chrome com opções personalizadas."""
    CACHE = BASE_DIR / 'perfil-path'
    perfil_path = str(CACHE)
    options = uc.ChromeOptions()
    options.add_argument(f"--user-data-dir={perfil_path}")
    options.add_experimental_option("prefs", {
        "download.default_directory": str(PASTA_COMPETENCIA),
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
    })
    try:
        driver = uc.Chrome(options=options)
        driver.get('https://cav.receita.fazenda.gov.br/autenticacao/login')
        driver.maximize_window()
        driver.implicitly_wait(10)
        return driver
    except Exception as e:
        logging.error(f"Erro ao configurar o driver: {e}")
        raise

def login(driver):
    """Realiza login no sistema."""
    try:
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//input[@alt="Acesso Gov BR"]'))).click()

        time.sleep(2)

        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="login-certificate"]'))).click()

        time.sleep(3)
        pyautogui.press('enter')

    except Exception as e:
        logging.error(f"Erro no login: {e}")
        raise

def navegacao(driver):
    try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))).click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//li[@id="btn214"]'))).click()

        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="containerServicos214"]/div[2]/ul/li[1]/a'))).click()

        time.sleep(2)
        bt_captcha1 = os.path.join(IMAGEM_DIR, 'bt_captcha.png')
        time.sleep(0.5)
        clique([bt_captcha1], 30, confidence=0.8)

        time.sleep(2)
        bt_prosseguir1 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png')
        time.sleep(0.5)
        clique([bt_prosseguir1], 30, confidence=0.8)

        bt_souprocurador = os.path.join(IMAGEM_DIR, 'bt_souprocurador.png')
        time.sleep(0.5)
        clique([bt_souprocurador], 30, confidence=0.8)

        time.sleep(10)

        bt_calendario = os.path.join(IMAGEM_DIR, 'bt_calendario.png')
        time.sleep(1)
        clique([bt_calendario], 30, confidence=0.8)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('delete')

        pyautogui.write(data_inicial)

        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('delete')
        pyautogui.write(data_final)

    except Exception as e:
        logging.error(f'Erro de navegacao: {e}')
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))).click()

def transmissao(cnpjs, codigos, df, driver):
    for cnpj, codigo in zip(cnpjs, codigos):

        status_series = df.loc[df['CNPJ'] == cnpj, 'STATUS']
        if status_series.astype(str).str.contains('Erro no download|Guia baixada|Nenhuma declaração encontrada', case=False, regex=True).any():
            logging.info(f"CNPJ {cnpj} já processado. Pulando...")
            break

        tentativas = 3
        while tentativas > 0:
            try:
                logging.info(f'Iniciando a transmissão da empresa: {cnpj}')

                bt_seta = os.path.join(IMAGEM_DIR, 'bt_seta.png')
                time.sleep(1)
                clique2([bt_seta], 30, confidence=0.8, ocorrencia=1)

                bt_cancelar = os.path.join(IMAGEM_DIR, 'bt_cancelar.png')
                time.sleep(1)
                clique([bt_cancelar], 30, confidence=0.8)

                pyautogui.write(cnpj)
                time.sleep(1)

                pyautogui.press('enter')
                time.sleep(1)

                pyautogui.click(x=6, y=728)
                time.sleep(1)

                bt_pesquisar = os.path.join(IMAGEM_DIR, 'bt_pesquisar.png')
                time.sleep(1)  # Aumentar o tempo de espera
                clique([bt_pesquisar], 30, confidence=0.8)

                time.sleep(1)
                erro_sem_declaracoes = os.path.join(IMAGEM_DIR, 'rc_sem_declaracoes.png')

                time.sleep(5)

                try:
                    reconhecimento([erro_sem_declaracoes], 30, confidence=0.8)

                    logging.info("Nenhuma declaração encontrada.")

                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Nenhuma declaração encontrada'

                    df.to_excel('Clientes.xlsx', index=False)

                    break

                except Exception as e:
                    logging.info('Declaracao encontrada')

                    bt_visualizar = os.path.join(IMAGEM_DIR, 'bt_visualizar.png')
                    time.sleep(1)
                    clique([bt_visualizar], 30, confidence=0.8)

                try:
                    bt_emitir_darf = os.path.join(IMAGEM_DIR, 'bt_emitir_darf.png')
                    time.sleep(1)
                    clique([bt_emitir_darf], 30, confidence=0.8)

                    time.sleep(2)

                    renomear_arquivo_recente(codigo, competencia)

                    logging.info(f"Download successful for {cnpj}")
                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
                    df.to_excel('Clientes.xlsx', index=False)

                    bt_ok = os.path.join(IMAGEM_DIR, 'bt_okk.png')
                    time.sleep(1)
                    clique([bt_ok], 30, confidence=0.8)

                    bt_voltar = os.path.join(IMAGEM_DIR, 'bt_voltar.png')
                    time.sleep(1)
                    clique([bt_voltar], 15, confidence=0.8)

                    bt_captcha2 = os.path.join(IMAGEM_DIR, 'bt_captcha.png')
                    time.sleep(1.5)
                    clique([bt_captcha2], 15, confidence=0.8)

                    time.sleep(1)

                    bt_prossegui2 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png')
                    time.sleep(1.5)
                    clique([bt_prossegui2], 15, confidence=0.8)

                    break

                except Exception as e:
                    logging.error("Nenhuma declaração encontrada.")

                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Possivel erro no e-CAC, verificar cliente manualmente'

                    df.to_excel('Clientes.xlsx', index=False)

                bt_voltar = os.path.join(IMAGEM_DIR, 'bt_voltar.png')
                time.sleep(1)
                clique([bt_voltar], 30, confidence=0.8)

                time.sleep(1.5)

                bt_captcha2 = os.path.join(IMAGEM_DIR, 'bt_captcha.png')
                time.sleep(1.5)
                clique([bt_captcha2], 30, confidence=0.8)

                time.sleep(1)

                bt_prossegui2 = os.path.join(IMAGEM_DIR, 'bt_prosseguir.png')
                time.sleep(1.5)
                clique([bt_prossegui2], 30, confidence=0.8)

                time.sleep(1)

                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'

                df.to_excel('Clientes.xlsx', index=False)

                break

            except Exception as e:
                logging.error(f"Erro no processamento do cliente {cnpj}: {e}")

                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Erro no download'

                df.to_excel('Clientes.xlsx', index=False)

                tentativas -= 1
                if tentativas > 0:
                    logging.info(f"Tentando novamente ({tentativas} tentativas restantes)")
                    navegacao(driver)
                else:
                    logging.error(f"Falha após {3} tentativas para o cliente {cnpj}")

sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def load_clients_spreadsheet():
    """Load the Clientes.xlsx spreadsheet."""
    try:
        return pd.read_excel('Clientes.xlsx')
    except FileNotFoundError:
        st.warning("Planilha de clientes não encontrada. Por favor, verifique o arquivo 'Clientes.xlsx'.")
        return pd.DataFrame(columns=['CNPJ', 'COD', 'STATUS'])

def save_clients_spreadsheet(df):
    """Save the updated DataFrame to Clientes.xlsx."""
    df.to_excel('Clientes.xlsx', index=False)

def run_automation_script():
    """Run the main automation script."""
    driver = configurar_driver()

    cnpjs, codigos, df = ler_planilha()

    login(driver)

    navegacao(driver)

    transmissao(cnpjs, codigos, df, driver)

    df.to_excel('Clientes.xlsx', index=False)

def main():
    st.title("DCTF Automation Dashboard")

    # Sidebar for configuration
    st.sidebar.header("Configurações")

    # Date range selection
    with st.sidebar.expander("Período de Apuração"):
        col1, col2 = st.columns(2)
        with col1:
            data_inicial = st.text_input("Data Inicial", key="initial_date")
        with col2:
            data_final = st.text_input("Data Final", key="final_date")

        competencia = st.text_input("Competência (MM YYYY)",
                                    help="Formato: MM YYYY, ex: 10 2024")

    # Load clients data
    df = load_clients_spreadsheet()

    # Main content area
    tab1, tab2, tab3 = st.tabs(["Status da Automação", "Lista de Clientes", "Executar Automação"])

    with tab1:
        st.header("Status da Automação")

        # Status summary
        status_counts = df['STATUS'].value_counts()
        col1, col2, col3 = st.columns(3)

        with col1:
            st.metric("Total de Clientes", len(df))
        with col2:
            st.metric("Guias Baixadas",
                      status_counts.get('Guia baixada', 0),
                      help="Número de empresas com DARF baixado")
        with col3:
            st.metric("Erros",
                      status_counts.get('Erro no download', 0),
                      help="Número de empresas com erro no download")

        # Pie chart of status
        st.subheader("Distribuição de Status")
        st.bar_chart(status_counts)

    with tab2:
        st.header("Lista de Clientes")

        # Filterable table
        search_term = st.text_input("Filtrar por CNPJ")

        if search_term:
            filtered_df = df[df['CNPJ'].astype(str).str.contains(search_term)]
        else:
            filtered_df = df

        st.dataframe(filtered_df,
                     column_config={
                         "CNPJ": st.column_config.TextColumn("CNPJ"),
                         "STATUS": st.column_config.TextColumn("Status")
                     },
                     hide_index=True)

    with tab3:
        st.header("Executar Automação")

        # Automation controls
        st.warning("""
        Antes de executar:
        - Certifique-se de que o navegador Chrome está fechado
        - Verifique as configurações de data e competência no menu lateral
        - Tenha o certificado digital configurado
        """)

        if st.button("Iniciar Automação", type="primary"):
            with st.spinner("Executando automação..."):
                run_automation_script()

            # Reload data after automation
            df = load_clients_spreadsheet()
            st.rerun()

if __name__ == "__main__":
    main()
