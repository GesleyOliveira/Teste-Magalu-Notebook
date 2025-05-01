import os
import time
import re
import pandas as pd
#import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException
import yagmail
from dotenv import load_dotenv



# ========================
# CONFIGURAÇÕES INICIAIS
# ========================
load_dotenv("config/credentials.env")
EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_APP = os.getenv("SENHA_APP")
EMAIL_DESTINATARIO = EMAIL_REMETENTE

output_dir = "Output"
os.makedirs(output_dir, exist_ok=True)
arquivo_excel = os.path.join(output_dir, "Notebooks.xlsx")

# ========================
# INICIAR NAVEGADOR
# ========================
print("[INFO] Iniciando navegador...")
chrome_options = Options()
# chrome_options.add_argument("--headless")  # Ative isso se não quiser abrir o navegador
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 5)

# ========================
# ACESSAR SITE
# ========================
print("[INFO] Acessando o site...")
url = "https://www.magazineluiza.com.br"
for tentativa in range(3):
    try:
        print(f"[INFO] Tentando acessar o site (tentativa {tentativa+1})...")
        driver.get(url)
        #time.sleep(4)
        if "Magazine Luiza" in driver.title:
            break
    except:
        time.sleep(2)
else:
    print("[ERRO] Site fora do ar.")
    with open("log.txt", "w") as log:
        log.write("Site fora do ar\n")
    driver.quit()
    exit(1)

# ========================
# INTERAGIR COM A PÁGINA
# ========================
try:
    # Aceitar cookies se necessário
    #consent_btns = driver.find_elements(By.CSS_SELECTOR, "button[data-testid='privacy-modal-accept']")
    #if consent_btns:
    #    print("[INFO] Aceitando cookies...")
    #    consent_btns[0].click()
    #    time.sleep(2)

    print("[INFO] Buscando barra de pesquisa...")
    barra_pesquisa = wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "input[data-testid='input-search']"))
    )
    barra_pesquisa.clear()
    barra_pesquisa.send_keys("notebooks", Keys.ENTER)
    print("[DEBUG] Comando de busca enviado.")
    wait.until(EC.url_contains("notebook"))
    print("[DEBUG] Página de resultados carregada:", driver.current_url)

except TimeoutException:
    print("[ERRO] Não foi possível encontrar a barra de pesquisa após 20s.")
    print("[DEBUG] HTML atual:")
    print(driver.page_source[:1000])
    driver.quit()
    exit(1)


# ========================
# COLETA DE PRODUTOS
# ========================
print("[INFO] Coletando produtos de todas as páginas...")
dados = []
page = 1  # começamos na página 1

while True:
    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "li.sc-hwhvDX")))
        produtos = driver.find_elements(By.CSS_SELECTOR, "li.sc-hwhvDX")

        print(f"[DEBUG] {len(produtos)} produtos encontrados na página {page}.")

        for p in produtos:
            try:
                nome = p.find_element(By.CSS_SELECTOR, "h2[data-testid='product-title']").text.strip()
                url_prod = p.find_element(By.TAG_NAME, "a").get_attribute("href")

                try:
                    aval_texto = p.find_element(By.CSS_SELECTOR, "span.sc-boZgaH").text
                    match = re.search(r'\((\d+)\)', aval_texto)
                    qtd_aval = int(match.group(1)) if match else 0
                except:
                    qtd_aval = 0

                dados.append([nome, qtd_aval, url_prod])

            except Exception as e:
                print("[AVISO] Produto ignorado:", e)

        # Tenta encontrar o botão da próxima página
        try:
            next_button = driver.find_element(
                By.CSS_SELECTOR, f"a[data-testid='pagination-item'][title='página {page + 1}']"
            )
            driver.execute_script("arguments[0].click();", next_button)
            page += 1
        except Exception as e:
            print("[INFO] Fim da paginação detectado ou próximo botão não encontrado.")
            break

    except Exception as e:
        print("[ERRO] Interrompido devido a erro inesperado:", e)
        break

driver.quit()


print(f"[INFO] {len(dados)} produtos com avaliação coletados.")

# ========================
# EXPORTAÇÃO PARA EXCEL
# ========================
df = pd.DataFrame(dados, columns=["PRODUTO", "QTD_AVAL", "URL"])
df = df[df["QTD_AVAL"].notnull()]
piores = df[df["QTD_AVAL"] < 100]
melhores = df[df["QTD_AVAL"] >= 100]

print("[INFO] Salvando arquivo Excel...")
with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
    piores.to_excel(writer, sheet_name="Piores", index=False)
    melhores.to_excel(writer, sheet_name="Melhores", index=False)
print("[INFO] Arquivo salvo em:", arquivo_excel)

# ========================
# ENVIO DE E-MAIL
# ========================
try:
    print("[INFO] Enviando e-mail com o relatório...")
    yag = yagmail.SMTP(EMAIL_REMETENTE, SENHA_APP)
    yag.send(
        to=EMAIL_DESTINATARIO,
        subject="Relatório Notebooks",
        contents="Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.",
        attachments=arquivo_excel
    )
    print("[INFO] E-mail enviado com sucesso.")
except Exception as e:
    print("[ERRO] Falha ao enviar e-mail:", e)
