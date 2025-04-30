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
import yagmail
from dotenv import load_dotenv

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

load_dotenv("config/credentials.env")

EMAIL_REMETENTE = os.getenv("EMAIL_REMETENTE")
SENHA_APP = os.getenv("SENHA_APP")
EMAIL_DESTINATARIO = EMAIL_REMETENTE

output_dir = "Output"
os.makedirs(output_dir, exist_ok=True)
arquivo_excel = os.path.join(output_dir, "Notebooks.xlsx")

# ========================
# ABRIR NAVEGADOR E ACESSAR SITE
# ========================
print("[INFO] Iniciando navegador...")
chrome_options = Options()
# chrome_options.add_argument("--headless")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# TENTAR ACESSAR O SITE ATÉ 3 VEZES
url = "https://www.magazineluiza.com.br"
attempts = 0
success = False
while attempts < 3:
    try:
        print(f"[INFO] Tentando acessar o site (tentativa {attempts+1})...")
        driver.get(url)
        time.sleep(4)
        if "Magazine Luiza" in driver.title:
            success = True
            break
    except Exception as e:
        attempts += 1
        time.sleep(2)

if not success:
    with open("log.txt", "w") as log:
        log.write("Site fora do ar\n")
    print("[ERRO] Site fora do ar. Encerrando.")
    driver.quit()
    exit(1)

print("[INFO] Pesquisando por notebooks...")
wait = WebDriverWait(driver, 10)

try:
    
    barra_pesquisa = wait.until(EC.presence_of_element_located((By.ID, "input-search")))
    barra_pesquisa.send_keys("notebooks")
    barra_pesquisa.submit()
    time.sleep(5)
except Exception as e:
    print("[ERRO] Não foi possível encontrar a barra de pesquisa:", e)
    driver.quit()
    exit(1)

print("[INFO] Coletando produtos de todas as páginas...")
dados = []
while True:
    produtos = driver.find_elements(By.CSS_SELECTOR, "li[data-testid='product-card']")
    for p in produtos:
        try:
            nome = p.find_element(By.CSS_SELECTOR, "h2").text.strip()
            url_prod = p.find_element(By.TAG_NAME, "a").get_attribute("href")
            try:
                aval = p.find_element(By.CSS_SELECTOR, "span[class*='reviewCount']").text
                qtd_aval = int(aval.strip("()").split()[0])
            except:
                qtd_aval = 0  # Produto sem avaliação visível
            dados.append([nome, qtd_aval, url_prod])
        except Exception as e:
            print("[AVISO] Produto ignorado por erro:", e)
            continue


    # Próxima página
    try:
        proxima = driver.find_element(By.CSS_SELECTOR, "a[data-testid='pagination-button-next']")
        if 'disabled' in proxima.get_attribute("class"):
            break
        driver.execute_script("arguments[0].click();", proxima)
        time.sleep(5)
    except:
        break

driver.quit()
print(f"[INFO] {len(dados)} produtos com avaliação coletados.")

# ========================
# TRATAMENTO E EXPORTAÇÃO
# ========================
df = pd.DataFrame(dados, columns=["PRODUTO", "QTD_AVAL", "URL"])

df = df[df["QTD_AVAL"].notnull()]  # Remove sem avaliações
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
        contents="Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza. Atenciosamente, Robô",
        attachments=arquivo_excel
    )
    print("[INFO] E-mail enviado com sucesso.")
except Exception as e:
    print("[ERRO] Falha ao enviar e-mail:", e)

# ========================
# GARANTIR FECHAMENTO DO NAVEGADOR
# ========================
try:
    driver.quit()
except:
    pass