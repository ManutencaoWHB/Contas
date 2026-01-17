# ===============================
# IMPORTS
# ===============================
import os
import time
import calendar
import pandas as pd
from datetime import date
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

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

# ===============================
# CONFIGURAÇÃO CHROME (HEADLESS)
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
# FUNÇÕES DE APOIO
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

def ler_tabela_regra_linha_3(caminho_arquivo):
    """
    Lê o arquivo HTML/XLS e aplica a regra estrita:
    O cabeçalho está na LINHA 3 (índice 2).
    """
    print(f"📖 Lendo: {os.path.basename(caminho_arquivo)}")
    
    # 1. Lê todas as tabelas sem assumir cabeçalho (header=None)
    # Isso traz TUDO que está no arquivo bruto
    dfs = pd.read_html(caminho_arquivo, decimal=",", thousands=".", header=None)
    
    if not dfs:
        raise ValueError("Nenhuma tabela encontrada no arquivo.")

    # 2. Pega a maior tabela (onde estão os dados)
    df = max(dfs, key=len)
    
    print(f"   - Linhas brutas encontradas: {len(df)}")
    
    # 3. Aplica a Regra da Linha 3
    # Verifica se tem linhas suficientes
    if len(df) > 3:
        # A linha 3 (índice 2) vira o nome das colunas
        df.columns = df.iloc[2] 
        # Pega da linha 4 (índice 3) para baixo (os dados)
        df = df[3:].reset_index(drop=True)
        print("   - Regra aplicada: Cabeçalho definido na Linha 3.")
    else:
        print("   ⚠️ AVISO: Tabela tem menos de 3 linhas. Retornando bruta.")
    
    print(f"   - Linhas de dados finais: {len(df)}")
    return df

# ===============================
# EXECUÇÃO
# ===============================
try:
    if not SENHA:
        raise ValueError("Senha não definida nas Secrets!")

    print("🔐 Login...")
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "login"))).send_keys(USUARIO)
    driver.find_element(By.ID, "senha").send_keys(SENHA)
    driver.find_element(By.ID, "submitButton").click()
    wait.until(EC.url_to_be(URL_HOME))
    print("✅ Login OK.")

    data_ini, data_fim = periodo_mes_atual()

    # ---------------------------------------------------------
    # 1. PCP347 (ENTRADA)
    # ---------------------------------------------------------
    print("📄 Baixando PCP347 (Entrada)...")
    driver.get(URL_PCP347)
    wait.until(EC.url_contains("pcp347"))

    driver.find_element(By.ID, "de_data").clear()
    driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear()
    driver.find_element(By.ID, "ate_data").send_keys(data_fim)
    
    Select(driver.find_element(By.ID, "str_fil")).select_by_visible_text("WHB CTBA")
    Select(driver.find_element(By.ID, "str_planta")).select_by_visible_text("USINAGEM CTBA")
    
    driver.find_element(By.XPATH, "//button[.//i[contains(@class,'fa-check')]]").click()
    time.sleep(10)
    
    # Salva o HTML atual
    html_pcp = salvar_html_pagina("pcp347_temp.html")
    
    # Processa com a regra da linha 3
    df_entrada = ler_tabela_regra_linha_3(html_pcp)

    # ---------------------------------------------------------
    # 2. SD3 (CONSUMO)
    # ---------------------------------------------------------
    print("📊 Baixando SD3 (Consumo)...")
    driver.execute_script("wl('/cus027')")
    
    wait.until(EC.url_contains("cus027"))
    driver.find_element(By.ID, "de_data").clear()
    driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear()
    driver.find_element(By.ID, "ate_data").send_keys(data_fim)

    Select(driver.find_element(By.ID, "str_emp")).select_by_visible_text("WHB AUTOMOTIVE / CURITIBA")
    Select(driver.find_element(By.ID, "str_consumo")).select_by_visible_text("SIM")
    driver.find_element(By.ID, "ate_cod").send_keys("ZZZZZZZZZZZZZZZ")
    driver.find_element(By.ID, "ate_tipo").send_keys("ZZ")

    arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
    driver.find_element(By.XPATH, "//button[.//i[contains(@class,'fa-check')]]").click()

    print("⏳ Aguardando download...")
    arquivo_sd3 = None
    for _ in range(60):
        novos = set(os.listdir(DOWNLOAD_DIR)) - arquivos_antes
        for f in novos:
            if f.endswith(('.xls', '.xlsx')):
                arquivo_sd3 = os.path.join(DOWNLOAD_DIR, f)
                break
        if arquivo_sd3: break
        time.sleep(1)

    if not arquivo_sd3:
        raise Exception("Download SD3 falhou.")
    
    # Processa com a regra da linha 3
    df_consumo = ler_tabela_regra_linha_3(arquivo_sd3)

    # ---------------------------------------------------------
    # 3. SALVAR FINAL
    # ---------------------------------------------------------
    caminho_final = os.path.join(DOWNLOAD_DIR, "dados_dashboard.xlsx")
    
    print("🔄 Gerando arquivo consolidado...")
    with pd.ExcelWriter(caminho_final, engine='openpyxl') as writer:
        df_consumo.to_excel(writer, sheet_name="Consumo", index=False)
        df_entrada.to_excel(writer, sheet_name="Entrada", index=False)

    print(f"🎉 SUCESSO! Arquivo gerado: {caminho_final}")
    print(f"   - Aba Consumo: {len(df_consumo)} linhas")
    print(f"   - Aba Entrada: {len(df_entrada)} linhas")

except Exception as e:
    print(f"❌ ERRO: {e}")
    driver.save_screenshot("erro_final.png")
    raise e

finally:
    driver.quit()
