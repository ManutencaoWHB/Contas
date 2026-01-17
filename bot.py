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
# CONFIGURAÇÃO CHROME
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
wait = WebDriverWait(driver, 90) # Aumentado para garantir downloads lentos

# ===============================
# FUNÇÕES INTELIGENTES
# ===============================
def periodo_mes_atual():
    primeiro = DATA_HOJE.replace(day=1)
    _, ultimo_dia = calendar.monthrange(DATA_HOJE.year, DATA_HOJE.month)
    return primeiro.strftime("%d/%m/%Y"), ultimo.strftime("%d/%m/%Y")

def salvar_html_pagina(nome):
    path = os.path.join(DOWNLOAD_DIR, nome)
    with open(path, "w", encoding="utf-8") as f:
        f.write(driver.page_source)
    return path

def ler_tabela_inteligente(caminho_arquivo, nome_ref):
    """
    Lê o arquivo e procura a tabela correta baseada no CONTEÚDO,
    não apenas no tamanho.
    """
    print(f"📖 Lendo {nome_ref}: {os.path.basename(caminho_arquivo)}")
    
    try:
        # Tenta ler como HTML (padrão do portal)
        dfs = pd.read_html(caminho_arquivo, decimal=",", thousands=".", header=None)
    except Exception:
        # Se falhar, tenta ler como Excel real
        try:
            dfs = [pd.read_excel(caminho_arquivo, header=None)]
        except Exception as e:
            raise ValueError(f"Não foi possível ler o arquivo: {e}")

    if not dfs:
        raise ValueError("Nenhuma tabela encontrada.")

    print(f"   - {len(dfs)} tabelas encontradas.")
    
    tabela_escolhida = None
    
    # ESTRATÉGIA: Procura a tabela que tem cabeçalho real
    for i, df in enumerate(dfs):
        # Converte para string para buscar palavras-chave
        texto = df.head(10).astype(str).to_string().upper()
        
        # Se tiver 'PRODUTO', 'CUSTO' ou 'DESCRIÇÃO', é a nossa tabela!
        if "PRODUTO" in texto or "CUSTO" in texto or "DESC" in texto:
            print(f"   - ✅ Tabela {i} identificada por conteúdo.")
            tabela_escolhida = df
            break
    
    # Se não achou por palavra, usa a maior (fallback)
    if tabela_escolhida is None:
        print("   - ⚠️ Aviso: Conteúdo não reconhecido. Usando a maior tabela.")
        tabela_escolhida = max(dfs, key=len)

    # TRATAMENTO DO CABEÇALHO (LINHA 3)
    # Procuramos onde está o cabeçalho "Custo Moeda 1" ou similar
    idx_cabecalho = -1
    for idx, row in tabela_escolhida.head(10).iterrows():
        row_str = row.astype(str).str.upper().values
        if any("CUSTO" in str(x) for x in row_str):
            idx_cabecalho = idx
            break
            
    if idx_cabecalho != -1:
        print(f"   - Cabeçalho detectado na linha {idx_cabecalho + 1}.")
        tabela_escolhida.columns = tabela_escolhida.iloc[idx_cabecalho]
        tabela_escolhida = tabela_escolhida[idx_cabecalho + 1:].reset_index(drop=True)
    else:
        # Se não achou dinamicamente, força a regra da linha 3 (índice 2)
        print("   - Cabeçalho não detectado automaticamente. Forçando Linha 3.")
        if len(tabela_escolhida) > 2:
            tabela_escolhida.columns = tabela_escolhida.iloc[2]
            tabela_escolhida = tabela_escolhida[3:].reset_index(drop=True)

    print(f"   - Linhas de dados finais: {len(tabela_escolhida)}")
    return tabela_escolhida

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
    time.sleep(15) # Mais tempo para garantir carregamento
    
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
    for _ in range(90): # Aumentei timeout para 90s
        novos = set(os.listdir(DOWNLOAD_DIR)) - arquivos_antes
        for f in novos:
            if f.endswith(('.xls', '.xlsx')) and "crdownload" not in f:
                arquivo_sd3 = os.path.join(DOWNLOAD_DIR, f)
                break
        if arquivo_sd3: break
        time.sleep(1)

    if not arquivo_sd3: raise Exception("Download SD3 falhou.")
    
    # IMPORTANTE: Espera 2s extras para garantir que o arquivo foi escrito em disco
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
