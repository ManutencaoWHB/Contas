# ===============================
# IMPORTS
# ===============================
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import pandas as pd
import calendar
from datetime import date, datetime

# ===============================
# CONFIGURAÇÕES
# ===============================
URL_LOGIN = "https://portal.whbbrasil.com.br/"
URL_HOME = "https://portal.whbbrasil.com.br/Portalhome"
URL_PCP347 = "https://portal.whbbrasil.com.br/pcp347"

USUARIO = "luanfp"
SENHA = os.getenv("SUA_SENHA_PORTAL")

DOWNLOAD_DIR = os.getcwd()

DATA_HOJE = date.today()
DIA_STR = DATA_HOJE.strftime("%d-%m-%Y")

# ===============================
# CHROME OPTIONS
# ===============================
options = webdriver.ChromeOptions()
prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}
options.add_experimental_option("prefs", prefs)
options.add_argument("--headless=new") 
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--window-size=1920,1080")

driver = webdriver.Chrome(
    service=ChromeService(ChromeDriverManager().install()),
    options=options
)

wait = WebDriverWait(driver, 60)

# ===============================
# FUNÇÕES
# ===============================
def periodo_mes_atual():
    primeiro = DATA_HOJE.replace(day=1)
    _, ultimo_dia = calendar.monthrange(DATA_HOJE.year, DATA_HOJE.month)
    ultimo = DATA_HOJE.replace(day=ultimo_dia)
    return primeiro.strftime("%d/%m/%Y"), ultimo.strftime("%d/%m/%Y")

def salvar_html_pagina(nome):
    path = os.path.join(DOWNLOAD_DIR, nome)
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    return path

# ===============================
# EXECUÇÃO
# ===============================
try:
    if not SENHA:
        raise ValueError("Senha não definida!")

    print("🔐 Realizando Login...")
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "login"))).send_keys(USUARIO)
    driver.find_element(By.ID, "senha").send_keys(SENHA)
    driver.find_element(By.ID, "submitButton").click()
    wait.until(EC.url_to_be(URL_HOME))
    print("✅ Login realizado.")

    data_ini, data_fim = periodo_mes_atual()

    # ---------------------------------------------------------
    # 1. PCP347 (ENTRADA)
    # ---------------------------------------------------------
    print("📄 Acessando PCP347...")
    driver.get(URL_PCP347)
    wait.until(EC.url_contains("pcp347"))

    wait.until(EC.element_to_be_clickable((By.ID, "de_data"))).clear()
    driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear()
    driver.find_element(By.ID, "ate_data").send_keys(data_fim)

    Select(driver.find_element(By.ID, "str_fil")).select_by_visible_text("WHB CTBA")
    Select(driver.find_element(By.ID, "str_planta")).select_by_visible_text("USINAGEM CTBA")

    botao_ok_xpath = "//button[.//i[contains(@class,'fa-check')]]"
    wait.until(EC.element_to_be_clickable((By.XPATH, botao_ok_xpath))).click()

    time.sleep(10)
    
    # Lê a tabela direto do HTML da página
    html_path = salvar_html_pagina(f"pcp347_temp.html")
    df_entrada = pd.read_html(html_path, decimal=",", thousands=".")[0]
    print(f"✅ Dados PCP347 (Entrada): {len(df_entrada)} linhas.")

    # ---------------------------------------------------------
    # 2. SD3 (CONSUMO)
    # ---------------------------------------------------------
    print("📊 Acessando SD3...")
    driver.execute_script("wl('/cus027')")
    
    wait.until(EC.url_contains("cus027"))
    wait.until(EC.element_to_be_clickable((By.ID, "de_data"))).clear()
    driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear()
    driver.find_element(By.ID, "ate_data").send_keys(data_fim)

    Select(driver.find_element(By.ID, "str_emp")).select_by_visible_text("WHB AUTOMOTIVE / CURITIBA")
    Select(driver.find_element(By.ID, "str_consumo")).select_by_visible_text("SIM")

    driver.find_element(By.ID, "ate_cod").send_keys("ZZZZZZZZZZZZZZZ")
    driver.find_element(By.ID, "ate_tipo").send_keys("ZZ")

    arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
    wait.until(EC.element_to_be_clickable((By.XPATH, botao_ok_xpath))).click()

    print("⏳ Aguardando download SD3...")
    arquivo_sd3_path = None
    
    for _ in range(60):
        arquivos_agora = set(os.listdir(DOWNLOAD_DIR))
        novos = arquivos_agora - arquivos_antes
        for arquivo in novos:
            if arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):
                arquivo_sd3_path = os.path.join(DOWNLOAD_DIR, arquivo)
                break
        if arquivo_sd3_path:
            break
        time.sleep(1)

    if not arquivo_sd3_path:
        raise Exception("Download SD3 falhou.")

    print(f"✅ SD3 baixado: {arquivo_sd3_path}")
    
    # --- CORREÇÃO DO ERRO ---
    print("🧹 Lendo SD3 (Tentativa Blindada)...")
    try:
        # Tenta ler como Excel binário antigo (.xls) - Requer xlrd
        if arquivo_sd3_path.endswith('.xls'):
            df_consumo = pd.read_excel(arquivo_sd3_path, header=2, engine='xlrd')
        else:
            df_consumo = pd.read_excel(arquivo_sd3_path, header=2)
    except Exception as e:
        print(f"⚠️ Erro ao ler como Excel ({e}). Tentando ler como HTML...")
        # Se falhar, tenta ler como HTML (muito comum o arquivo ser HTML renomeado para .xls)
        # O header=2 aqui tenta pular linhas se o HTML tiver cabeçalho sujo
        try:
            df_consumo = pd.read_html(arquivo_sd3_path, decimal=",", thousands=".", header=2)[0]
        except:
             # Última tentativa: ler HTML sem pular cabeçalho
             df_consumo = pd.read_html(arquivo_sd3_path, decimal=",", thousands=".")[0]
             # Remove as duas primeiras linhas manualmente
             df_consumo = df_consumo.iloc[2:].reset_index(drop=True)
             # Redefine o cabeçalho
             df_consumo.columns = df_consumo.iloc[0]
             df_consumo = df_consumo[1:]
             
    print(f"✅ SD3 lido com sucesso: {len(df_consumo)} linhas.")

    # ---------------------------------------------------------
    # 3. JUNTAR E SALVAR
    # ---------------------------------------------------------
    print("🔄 Gerando 'dados_dashboard.xlsx'...")
    caminho_final = os.path.join(DOWNLOAD_DIR, "dados_dashboard.xlsx")

    with pd.ExcelWriter(caminho_final, engine='openpyxl') as writer:
        df_consumo.to_excel(writer, sheet_name="Consumo", index=False)
        df_entrada.to_excel(writer, sheet_name="Entrada", index=False)

    print(f"🎉 ARQUIVO FINAL: {caminho_final}")

except Exception as e:
    print(f"❌ ERRO CRÍTICO: {e}")
    driver.save_screenshot("erro_debug.png")
    raise e

finally:
    driver.quit()
