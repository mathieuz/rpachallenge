""" Import das libs principais """
import pyautogui as pg
import pandas as pd
import time

""" Definindo métodos de captura de posição à partir da posição X e Y dos elementos na tela. """
def posX(pos):
    string = str(pos)
    indiceChar = string.find("x") + 2
    posX = str()
    
    while (string[indiceChar].isnumeric() == True):
        if (string[indiceChar].isnumeric() == True):
            posX += str(string[indiceChar])
            indiceChar += 1
        else:
            indiceChar += 1
        
    return int(posX)

def posY(pos):
    string = str(pos)
    indiceChar = string.find("y") + 2
    posY = str()
    
    while (string[indiceChar].isnumeric() == True):
        if (string[indiceChar].isnumeric() == True):
            posY += str(string[indiceChar])
            indiceChar += 1
        else:
            if (indiceChar <= string.__len__() - 1):
                indiceChar += 1
        
    return int(posY)

""" Extraindo o Dataframe da base de dados. """
dadosPlanilha = pd.read_excel("challenge.xlsx")
print(dadosPlanilha)

print(dadosPlanilha[dadosPlanilha.columns[2]][5])
print(dadosPlanilha[dadosPlanilha.columns[6]][0])

""" Definindo tempo de execução e habilitando o failsafe. """
pg.PAUSE = 0.5
pg.FAILSAFE = True

""" Abrindo o site rpachallenge.com """
pg.press("win")
time.sleep(2.5)

pg.write("Chrome")
time.sleep(2.5)

pg.press("enter")
time.sleep(2.5)

pg.write("https://rpachallenge.com/")
time.sleep(2.5)

pg.press("enter")
time.sleep(2.5)

pg.hotkey("win", "up")
time.sleep(2.5)

""" Iniciando... """
posStart = pg.locateCenterOnScreen("./imagens-do-input/btn/btn-start.png")
pg.click(posStart)

numColuna = 0
numLinha = 0
numPressSubmit = 0

while (numPressSubmit != 10):

    posFirstName = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-firstname.png", confidence = 0.9)
    pg.click(posX(posFirstName), posY(posFirstName) + 45)
    pg.write(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha])
    numColuna += 1

    posLastName = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-lastname.png", confidence = 0.9)
    pg.click(posX(posLastName), posY(posLastName) + 45)
    pg.write(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha])
    numColuna += 1

    posCompanyName = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-companyname.png", confidence = 0.9)
    pg.click(posX(posCompanyName), posY(posCompanyName) + 45)
    pg.write(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha])
    numColuna += 1

    posRoleInCompany = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-roleincompany.png", confidence = 0.9)
    pg.click(posX(posRoleInCompany), posY(posRoleInCompany) + 45)
    pg.write(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha])
    numColuna += 1

    posAddress = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-address.png", confidence = 0.9)
    pg.click(posX(posAddress), posY(posAddress) + 45)
    pg.write(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha])
    numColuna += 1

    posEmail = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-email.png", confidence = 0.9)
    pg.click(posX(posEmail), posY(posEmail) + 45)
    pg.write(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha])
    numColuna += 1

    posPhoneNumber = pg.locateCenterOnScreen("./imagens-do-input/label-input/label-phonenumber.png", confidence = 0.9)
    pg.click(posX(posPhoneNumber), posY(posPhoneNumber) + 45)
    pg.write(str(dadosPlanilha[dadosPlanilha.columns[numColuna]][numLinha]))
    
    posSubmit = pg.locateCenterOnScreen("./imagens-do-input/btn/btn-submit.png")
    pg.click(posSubmit)

    numLinha += 1
    numColuna = 0

    numPressSubmit += 1


pg.hotkey("win", "down")