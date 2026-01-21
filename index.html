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
wait = WebDriverWait(driver, 90)

# ===============================
# FUNÇÕES INTELIGENTES
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

def normalizar_cabecalho(df_bruto, nome_ref):
    """
    Procura a linha que contém a palavra 'Produto' ou 'Custo' para definir como cabeçalho.
    """
    print(f"   --- Diagnóstico Visual do {nome_ref} ---")
    # Imprime as primeiras 5 linhas para sabermos o que o Python está vendo
    print(df_bruto.head(5).to_string()) 
    print("   ----------------------------------------")

    # 1. Verifica se o cabeçalho já está nas colunas (Python leu direto)
    colunas_str = " ".join([str(c).upper() for c in df_bruto.columns])
    if "PRODUTO" in colunas_str or "CUSTO" in colunas_str:
        print(f"   ✅ Cabeçalho já identificado corretamente nas colunas do {nome_ref}.")
        return df_bruto

    # 2. Procura a linha de cabeçalho nos dados
    idx_cabecalho = -1
    for idx, row in df_bruto.head(15).iterrows():
        # Converte linha para texto maiúsculo
        linha_txt = row.astype(str).str.upper().values
        # Procura palavras chaves que existem nos dois arquivos
        if any("PRODUTO" in str(x) for x in linha_txt) and any("DESC" in str(x) for x in linha_txt):
            idx_cabecalho = idx
            break
            
    if idx_cabecalho != -1:
        print(f"   ✅ Cabeçalho real encontrado na linha índice {idx_cabecalho} do {nome_ref}.")
        df_bruto.columns = df_bruto.iloc[idx_cabecalho] # Define o novo cabeçalho
        df_final = df_bruto[idx_cabecalho + 1:].reset_index(drop=True) # Pega os dados abaixo
        return df_final
    
    # 3. Fallback: Se não achou nada, retorna como está (melhor que zerar)
    print(f"   ⚠️ AVISO: Não encontrei a palavra 'Produto' nas primeiras linhas do {nome_ref}. Mantendo original.")
    return df_bruto

def ler_tabela_inteligente(caminho_arquivo, nome_ref):
    print(f"📖 Lendo {nome_ref}: {os.path.basename(caminho_arquivo)}")
    
    try:
        # Lê todas as tabelas (sem assumir cabeçalho)
        dfs = pd.read_html(caminho_arquivo, decimal=",", thousands=".", header=None)
    except Exception:
        try:
            dfs = [pd.read_excel(caminho_arquivo, header=None)]
        except Exception as e:
            raise ValueError(f"Erro leitura: {e}")

    if not dfs:
        raise ValueError("Nenhuma tabela encontrada.")

    # Escolhe a maior tabela (onde tem dados)
    df_escolhido = max(dfs, key=len)
    print(f"   - Tabela bruta selecionada com {len(df_escolhido)} linhas.")
    
    # Aplica a normalização do cabeçalho
    df_limpo = normalizar_cabecalho(df_escolhido, nome_ref)
    
    print(f"   - Linhas de dados finais: {len(df_limpo)}")
    return df_limpo

# ===============================
# EXECUÇÃO
# ===============================
try:
    if not SENHA: raise ValueError("Senha não definida!")

    print("🔐 Login...")
    driver.get(URL_LOGIN)
    wait.until(EC.presence_of_element_located((By.ID, "login"))).send_keys(USUARIO)
    driver.find_element(By.ID, "senha").send_keys(SENHA)
    driver.find_element(By.ID, "submitButton").click()
    wait.until(EC.url_to_be(URL_HOME))
    
    data_ini, data_fim = periodo_mes_atual()
    
    # 1. PCP347 (ENTRADA)
    print("📄 PCP347 (Entrada)...")
    driver.get(URL_PCP347)
    wait.until(EC.url_contains("pcp347"))
    driver.find_element(By.ID, "de_data").clear(); driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear(); driver.find_element(By.ID, "ate_data").send_keys(data_fim)
    Select(driver.find_element(By.ID, "str_fil")).select_by_visible_text("WHB CTBA")
    Select(driver.find_element(By.ID, "str_planta")).select_by_visible_text("USINAGEM CTBA")
    driver.find_element(By.XPATH, "//button[.//i[contains(@class,'fa-check')]]").click()
    time.sleep(15)
    
    html_pcp = salvar_html_pagina("pcp347_temp.html")
    df_entrada = ler_tabela_inteligente(html_pcp, "PCP347")

    # 2. SD3 (CONSUMO)
    print("📊 SD3 (Consumo)...")
    driver.execute_script("wl('/cus027')")
    wait.until(EC.url_contains("cus027"))
    driver.find_element(By.ID, "de_data").clear(); driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear(); driver.find_element(By.ID, "ate_data").send_keys(data_fim)
    Select(driver.find_element(By.ID, "str_emp")).select_by_visible_text("WHB AUTOMOTIVE / CURITIBA")
    Select(driver.find_element(By.ID, "str_consumo")).select_by_visible_text("SIM")
    driver.find_element(By.ID, "ate_cod").send_keys("ZZZZZZZZZZZZZZZ")
    driver.find_element(By.ID, "ate_tipo").send_keys("ZZ")
    
    arquivos_antes = set(os.listdir(DOWNLOAD_DIR))
    driver.find_element(By.XPATH, "//button[.//i[contains(@class,'fa-check')]]").click()

    print("⏳ Aguardando download SD3...")
    arquivo_sd3 = None
    for _ in range(90):
        novos = set(os.listdir(DOWNLOAD_DIR)) - arquivos_antes
        for f in novos:
            if f.endswith(('.xls', '.xlsx')) and "crdownload" not in f:
                arquivo_sd3 = os.path.join(DOWNLOAD_DIR, f)
                break
        if arquivo_sd3: break
        time.sleep(1)

    if not arquivo_sd3: raise Exception("Download SD3 falhou.")
    time.sleep(2)
    
    df_consumo = ler_tabela_inteligente(arquivo_sd3, "SD3")

    # 3. SALVAR
    caminho_final = os.path.join(DOWNLOAD_DIR, "dados_dashboard.xlsx")
    with pd.ExcelWriter(caminho_final, engine='openpyxl') as writer:
        df_consumo.to_excel(writer, sheet_name="Consumo", index=False)
        df_entrada.to_excel(writer, sheet_name="Entrada", index=False)

    print(f"🎉 FINALIZADO: {caminho_final}")
    print(f"   - Consumo: {len(df_consumo)} linhas")
    print(f"   - Entrada: {len(df_entrada)} linhas")

except Exception as e:
    print(f"❌ ERRO: {e}")
    driver.save_screenshot("erro_final.png")
    raise e
finally:
    driver.quit()
