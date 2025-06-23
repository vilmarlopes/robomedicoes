from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import smtplib
from email.message import EmailMessage
import os

# Configurações do navegador Chrome em modo headless
chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--disable-gpu')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
prefs = {"download.default_directory": "/home/seuusuario/"}  # Ajuste para o diretório do PythonAnywhere
chrome_options.add_experimental_option("prefs", prefs)

# Iniciar navegador
driver = webdriver.Chrome(options=chrome_options)

try:
    # Acessar portal e logar
    driver.get("https://origem.virtual360.io/users/sign_in")
    time.sleep(2)
    driver.find_element(By.ID, "user_email").send_keys("vilmarlopes@metaltec.com.br")
    driver.find_element(By.ID, "user_password").send_keys("Vrl$240203")
    driver.find_element(By.NAME, "commit").click()
    time.sleep(5)

    # Navegar até medições de serviços
    driver.get("https://origem.virtual360.io/measurements") # Ajuste se necessário
    time.sleep(5)

    # Clicar no botão "Exportar" (ajuste o seletor conforme necessário)
    export_btn = driver.find_element(By.XPATH, "//a[contains(text(),'Exportar')]")
    export_btn.click()
    time.sleep(15) # Tempo para download terminar (ajuste conforme sua internet e portal)

    # Arquivo baixado (ajuste o nome conforme arquivo gerado)
    filename = "Medicoes_de_Servicos.xlsx"
    filepath = f"/home/vilmarlopes/{filename}"

    # Enviar por e-mail
    msg = EmailMessage()
    msg["Subject"] = "Relatório de Medições Diário"
    msg["From"] = "vilmarlopes@gmail.com"
    msg["To"] = "vilmarlopes@metaltec.com.br"
    msg.set_content("Segue em anexo a planilha de medições exportada automaticamente do portal Origem.")

    with open(filepath, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename=filename)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login("vilmarlopes@gmail.com", "Vrl$1977")
        smtp.send_message(msg)

finally:
    driver.quit()
    # Se quiser, pode apagar o arquivo após envio:
    # os.remove(filepath)
