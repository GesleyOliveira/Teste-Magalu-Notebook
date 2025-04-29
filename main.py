import os
import time
import pandas as pandas
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import yagmail
from dotenv import load_dotenv


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
chrome_options.add_argument("--headless")
driver = webdriver.Chrome(options=chrome_options)

try:
    print("[INFO] Acessando Magazine Luiza...")
    driver.get("https://www.magazineluiza.com.br")
    time.sleep(4)

    print("[INFO] Pesquisando por notebooks...")
    barra_pesquisa = driver.find_element(By.NAME, "q")
    barra_pesquisa.send_keys("notebooks")
    barra_pesquisa.submit()
    time.sleep(5)

    print("[INFO] Coletando produtos da primeira página...")
    produtos = driver.find_elements(By.CSS_SELECTOR, "li[data-testid='product-card']")
    dados = []

    for p in produtos:
        try:
            nome = p.find_element(By.CSS_SELECTOR, "h2").text
            url = p.find_element(By.TAG_NAME, "a").get_attribute("href")
            aval = p.find_element(By.CSS_SELECTOR, "span[class*='reviewCount']").text
            qtd_aval = int(aval.strip("()").split()[0])
            dados.append([nome, qtd_aval, url])
        except Exception:
            continue

    driver.quit()
    print(f"[INFO] {len(dados)} produtos com avaliação coletados.")
except Exception as e:
    driver.quit()
    print("[ERRO] Ocorreu um problema com o navegador:", e)
    exit(1)

# ========================
# TRATAMENTO E EXPORTAÇÃO
# ========================
df = pd.DataFrame(dados, columns=["PRODUTO", "QTD_AVAL", "URL"])
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
