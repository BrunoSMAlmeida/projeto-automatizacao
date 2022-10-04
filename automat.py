import pandas as pd
import pyautogui
import pyperclip
import time
import pandas
import numpy
import openpyxl


pyautogui.PAUSE = 1

# Passo 1: Entrar no sistema da empresa (no nosso caso, o link do drive)
pyautogui.press("win")
pyautogui.write("chrome")
pyautogui.press("enter")
pyperclip.copy("https://drive.google.com/drive/folders/149xknr9JvrlEnhNWO49zPcw0PW5icxga")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(5)

# Passo 2: Navegar no sistema e encontrar nossa base de vendas (entrar na pasta exportar)
# Mapeamos os botões usando pyautogui.position() que mapeia onde vamos clicar, só colocando o mouse em cima do que queremos.
pyautogui.click(x=173, y=269)
time.sleep(3)

# Passo 3: Fazer download da base de vendas
pyautogui.click(x=94, y=263)
time.sleep(5)

# Passo 4: Importar a base de vendas pro python
tabela = pd.read_excel(r"C:\Users\Objectedge\Downloads\Vendas - Dez.xlsx")

# Passo 5: Calcular os indicadores da empresa
faturamento = tabela["Valor Final"].sum()
quantidade = tabela["Quantidade"].sum()

# Passo 6: Enviar o email com os resultados do relatório

pyautogui.hotkey("ctrl", "t")
pyperclip.copy("https://mail.google.com/mail/u/1/#inbox")
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")
time.sleep(7)
pyautogui.click(x=74, y=169)
pyautogui.write("bruno4893@hotmail.com")
pyautogui.press("tab")

pyautogui.press("tab")
pyperclip.copy("Relatório de Vendas")
pyautogui.hotkey("ctrl", "v")

pyautogui.press("tab")
texto = f"""
Prezados(as),
Segue relatório de vendas do dia de hoje:
Faturamento: R$ {faturamento:,.2f}
Quantidade de produtos vendidos: {quantidade:,}

Qualquer dúvida entrar em contato.
Att
Bruno.
"""
pyperclip.copy(texto)
pyautogui.hotkey("ctrl", "v")
pyautogui.hotkey("ctrl", "enter")