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
# FUNÇÕES AUXILIARES
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

def encontrar_tabela_correta(arquivo_path, nome_ref):
    """
    Lê o HTML e procura a tabela que contém dados reais.
    Ignora tabelas de layout (cabeçalhos, rodapés vazios).
    """
    try:
        # Lê TODAS as tabelas do arquivo
        dfs = pd.read_html(arquivo_path, decimal=",", thousands=".")
        
        if not dfs:
            raise ValueError(f"Nenhuma tabela encontrada em {nome_ref}")

        print(f"🔎 {nome_ref}: Encontradas {len(dfs)} tabelas internas.")
        
        df_escolhido = None

        # ESTRATÉGIA 1: Procurar tabela que tenha colunas de custo/produto
        for df in dfs:
            # Converte as primeiras linhas para texto para procurar palavras chave
            texto_topo = df.head(10).astype(str).to_string().upper()
            if "CUSTO" in texto_topo or "PRODUTO" in texto_topo or "DESCRIÇÃO" in texto_topo:
                df_escolhido = df
                print(f"✅ {nome_ref}: Tabela identificada por palavras-chave.")
                break
        
        # ESTRATÉGIA 2: Se não achou por palavra, pega a MAIOR tabela (mais linhas)
        if df_escolhido is None:
            print(f"⚠️ {nome_ref}: Palavras-chave não encontradas. Selecionando a maior tabela.")
            df_escolhido = max(dfs, key=len)

        # LIMPEZA DO CABEÇALHO
        # Procura onde começa o cabeçalho real (linha que tem "Custo" ou "Produto")
        # Isso resolve o problema de pular 2, 3 ou 0 linhas dinamicamente
        for i, row in df_escolhido.head(10).iterrows():
            linha_texto = row.astype(str).str.upper().values
            # Se encontrar "CUSTO" ou "PRODUTO" nesta linha, ela é o cabeçalho
            if any("CUSTO" in str(x) for x in linha_texto) or any("PRODUTO" in str(x) for x in linha_texto):
                print(f"🧹 {nome_ref}: Cabeçalho real encontrado na linha {i}.")
                df_escolhido.columns = df_escolhido.iloc[i] # Define essa linha como titulo
                df_escolhido = df_escolhido[i+1:].reset_index(drop=True) # Pega tudo abaixo dela
                break
        
        return df_escolhido

    except Exception as e:
        print(f"❌ Erro ao processar tabelas de {nome_ref}: {e}")
        # Se falhar a leitura HTML, tenta ler como Excel padrão (último recurso)
        try:
             return pd.read_excel(arquivo_path)
        except:
             raise e

# ===============================
# EXECUÇÃO PRINCIPAL
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
    print("📄 Processando PCP347 (Entrada)...")
    driver.get(URL_PCP347)
    wait.until(EC.url_contains("pcp347"))

    wait.until(EC.element_to_be_clickable((By.ID, "de_data"))).clear()
    driver.find_element(By.ID, "de_data").send_keys(data_ini)
    driver.find_element(By.ID, "ate_data").clear()
    driver.find_element(By.ID, "ate_data").send_keys(data_fim)
    
    Select(driver.find_element(By.ID, "str_fil")).select_by_visible_text("WHB CTBA")
    Select(driver.find_element(By.ID, "str_planta")).select_by_visible_text("USINAGEM CTBA")
    
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//i[contains(@class,'fa-check')]]"))).click()
    time.sleep(10)
    
    html_path = salvar_html_pagina("pcp347_temp.html")
    # Usa a nova função inteligente também para o PCP
    df_entrada = encontrar_tabela_correta(html_path, "PCP347")
    print(f"✅ Entrada: {len(df_entrada)} linhas.")

    # ---------------------------------------------------------
    # 2. SD3 (CONSUMO)
    # ---------------------------------------------------------
    print("📊 Processando SD3 (Consumo)...")
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
    wait.until(EC.element_to_be_clickable((By.XPATH, "//button[.//i[contains(@class,'fa-check')]]"))).click()

    print("⏳ Download SD3...")
    arquivo_sd3_path = None
    for _ in range(60):
        arquivos_agora = set(os.listdir(DOWNLOAD_DIR))
        novos = arquivos_agora - arquivos_antes
        for arquivo in novos:
            if arquivo.endswith(".xlsx") or arquivo.endswith(".xls"):
                arquivo_sd3_path = os.path.join(DOWNLOAD_DIR, arquivo)
                break
        if arquivo_sd3_path: break
        time.sleep(1)

    if not arquivo_sd3_path: raise Exception("Timeout download SD3.")
    print(f"✅ Baixado: {arquivo_sd3_path}")
    
    # AQUI ESTAVA O ERRO -> Agora usamos a função inteligente
    df_consumo = encontrar_tabela_correta(arquivo_sd3_path, "SD3")
    print(f"✅ Consumo: {len(df_consumo)} linhas.")

    # ---------------------------------------------------------
    # 3. SALVAR
    # ---------------------------------------------------------
    caminho_final = os.path.join(DOWNLOAD_DIR, "dados_dashboard.xlsx")
    with pd.ExcelWriter(caminho_final, engine='openpyxl') as writer:
        df_consumo.to_excel(writer, sheet_name="Consumo", index=False)
        df_entrada.to_excel(writer, sheet_name="Entrada", index=False)

    print(f"🎉 FINALIZADO: {caminho_final}")

except Exception as e:
    print(f"❌ ERRO: {e}")
    driver.save_screenshot("erro_final.png")
    raise e

finally:
    driver.quit()
