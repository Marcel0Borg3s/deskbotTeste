"""
WARNING:

Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.

Also, if you are using PyCharm or another IDE, make sure that you use the SAME Python interpreter
as your IDE.

If you get an error like:
```
ModuleNotFoundError: No module named 'botcity'
```

This means that you are likely using a different Python interpreter than the one used to install the dependencies.
To fix this, you can either:
- Use the same interpreter as your IDE and install your bot with `pip install --upgrade -r requirements.txt`
- Use the same interpreter as the one used to install the bot (`pip install --upgrade -r requirements.txt`)

Please refer to the documentation for more information at
https://documentation.botcity.dev/tutorials/python-automations/desktop/
"""

# Import for the Desktop Bot
from botcity.core import DesktopBot

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

#impotar as bilibliotecas de Excel
from botcity.plugins.excel import BotExcelPlugin

# Função para lidar com elementos não encontrados
def not_found(element_name):
    print(f"Elemento {element_name} não encontrado.")

# Instancia do bot
bot = DesktopBot()

# Instantiate the Excel plugin
bot_excel = BotExcelPlugin()

# Read from an Excel File
dados = bot_excel.read('E:\\RPA\BotCity\\Projetos\\novoDeskbot_teste\\deskbotTeste\\resources\\extrato_bot.xlsx').as_list("Planilha1")[1:]
    
# Executando o aplicativo Banco
bot.execute("E:\\RPA\BotCity\\Drives\\banco\Banco.exe")

for linha in dados:
    #Criar variável Debito/credito
    btnCD = linha[2]

    #Criar variável DescriçãoTransferência Recebida
    strDescricao = linha[1]

    #Criar variável Valor   
    strValor = linha[3]
    convertidostrValor = float(strValor)
    strValor = str(int(convertidostrValor))

    #Criar variável Data
    strData = linha[4]
    strData = strData.strftime("%d/%m/%Y")
    strData = str(strData)

    #Condicional para validar se Deb ou Credt
    if btnCD == "Débito":
        # Mapeando o botão de crédito/débito
        if not bot.find("btnDebito", matching=0.9, waiting_time=10000):
            not_found("btnDebito")
        bot.click()
    else:
        if not bot.find("btnCredito", matching=0.9, waiting_time=10000):
            not_found("btnCredito")
        bot.click()
   
    # Mapeando o campo Descricao
    if not bot.find("campoDescricao2", matching=0.5, waiting_time=10000):
        not_found("campoDescricao2")
    bot.click()
    bot.type_keys(["ctrl", "a"])
    bot.kb_type(strDescricao)

    # Mapeando o campo Valor
    if not bot.find("campoValor", matching=0.5, waiting_time=10000):
        not_found("campoValor")
    bot.click()
    bot.type_keys(["ctrl", "a"])
    bot.kb_type(strValor)

    # Mapeando o campo Data
    if not bot.find("campoData", matching=0.5, waiting_time=10000):
        not_found("campoData")  
    bot.click()
    bot.type_keys(["ctrl", "a"])
    bot.kb_type(strData)
    
    if not bot.find("btnGravar", matching=0.9, waiting_time=10000):
        not_found("btnGravar")
    bot.click()

    
    
    

    