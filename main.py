import os
import time

import pandas as pd
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


# INICIAR NAVEGADOR

print("[INFO] Iniciando navegador...")
chrome_options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 5)


# ACESSAR SITE

print("[INFO] Acessando o site...")
url = "https://www.magazineluiza.com.br"
for tentativa in range(3):
    try:
        print(f"[INFO] Tentando acessar o site (tentativa {tentativa+1})...")
        driver.get(url)
        time.sleep(4)
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


# INTERAGIR COM A PÁGINA

try:
    # Aceitar cookies se necessário
    time.sleep(2)
    consent_btns = driver.find_elements(By.CSS_SELECTOR, "button[data-testid='privacy-modal-accept']")
    if consent_btns:
        print("[INFO] Aceitando cookies...")
        consent_btns[0].click()
        time.sleep(2)

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


# COLETA DE PRODUTOS 

print("[INFO] Coletando produtos de todas as páginas...")
dados = []
max_pages = 17  

for page in range(1, max_pages + 1):
    if page == 1:
        page_url = "https://m.magazineluiza.com.br/busca/notebooks/"
    else:
        page_url = f"https://m.magazineluiza.com.br/busca/notebooks/?page={page}"

    print(f"[INFO] Acessando página {page}: {page_url}")
    driver.get(page_url)
    time.sleep(3) 

    try:
        wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[data-testid='product-card-container']")))
        produtos = driver.find_elements(By.CSS_SELECTOR, "a[data-testid='product-card-container']")

        print(f"[DEBUG] {len(produtos)} produtos encontrados na página {page}.")

        for p in produtos:
            try:
                nome = p.find_element(By.CSS_SELECTOR, "h2[data-testid='product-title']").text.strip()
                url_prod = p.get_attribute("href")
                try:
                    review_div = p.find_element(By.CSS_SELECTOR, "div[data-testid='review']")
                    aval_span = review_div.find_element(By.CSS_SELECTOR, "span.sc-boZgaH")
                    aval_texto = aval_span.text  
                    qtd_aval = int(aval_texto.split('(')[-1].replace(')', ''))  
                except:
                    qtd_aval = 0
                
                dados.append([nome, qtd_aval, url_prod])

            except Exception as e:
                if "no such element" in str(e) and "product-title" in str(e):
                    continue  
                else:
                    print("[AVISO] Produto ignorado:", e)

    except Exception as e:
        print(f"[ERRO] Falha ao coletar produtos na página {page}:", e)
        continue  

driver.quit()

print(f"[INFO] {len(dados)} produtos com avaliação coletados.")



# EXPORTAÇÃO PARA EXCEL 

df = pd.DataFrame(dados, columns=["PRODUTO", "QTD_AVAL", "URL"])


df = df[df["QTD_AVAL"] > 0]


piores = df[df["QTD_AVAL"] < 100]
melhores = df[df["QTD_AVAL"] >= 100]

print("[INFO] Salvando arquivo Excel...")
with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
    melhores.to_excel(writer, sheet_name="Melhores", index=False)
    piores.to_excel(writer, sheet_name="Piores", index=False)
print("[INFO] Arquivo salvo em:", arquivo_excel)



# ENVIO DE E-MAIL

try:
    print("[INFO] Enviando e-mail com o relatório...")
    yag = yagmail.SMTP(EMAIL_REMETENTE, SENHA_APP)
    yag.send(
        to=EMAIL_DESTINATARIO,
        subject="Relatório Notebooks",
        contents=[
            "Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.",
            "",
            "",
            "Atenciosamente,",
            "Robô"
        ],
        attachments=arquivo_excel
    )
    print("[INFO] E-mail enviado com sucesso.")
except Exception as e:
    print("[ERRO] Falha ao enviar e-mail:", e)
