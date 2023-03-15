import pandas as pd
from time import sleep

"""Extraindo o dataframe da base de dados."""
dadosPlanilha = pd.read_excel("challenge.xlsx")

print(dadosPlanilha)

"""Abrindo o site e realizando os processos de preenchimento dos campos."""
from playwright.sync_api import sync_playwright
with sync_playwright() as syncp:
    browser = syncp.chromium.launch(headless = False)
    pagina = browser.new_page()
    pagina.goto("https://rpachallenge.com")

    """Clica no botão de START."""
    pagina.locator('xpath=/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/button').click()

    """Variáveis iteráveis."""
    col = 0
    linha = 0

    """Se tratando de dez cadastros diferentes, os campos serão preenchidos e submitados até este valor."""
    for numRepeticoes in range(0, 10):

        pagina.get_by_text("First Name").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(dadosPlanilha[dadosPlanilha.columns[col]][linha])

        col += 1

        pagina.get_by_text("Last Name").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(dadosPlanilha[dadosPlanilha.columns[col]][linha])

        col += 1

        pagina.get_by_text("Company Name").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(dadosPlanilha[dadosPlanilha.columns[col]][linha])

        col += 1

        pagina.get_by_text("Role in Company").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(dadosPlanilha[dadosPlanilha.columns[col]][linha])

        col += 1

        pagina.get_by_text("Address").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(dadosPlanilha[dadosPlanilha.columns[col]][linha])

        col += 1

        pagina.get_by_text("Email").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(dadosPlanilha[dadosPlanilha.columns[col]][linha])

        col += 1

        pagina.get_by_text("Phone Number").click(click_count = 3)
        pagina.keyboard.press("Tab")
        pagina.keyboard.type(str(dadosPlanilha[dadosPlanilha.columns[col]][linha]))

        col = 0
        linha += 1

        pagina.locator('xpath=//html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()

    sleep(10)
    browser.close()