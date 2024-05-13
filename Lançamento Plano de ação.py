import os
import pathlib
import pyautogui as gui
import time 
import timeit
import pandas
import pyperclip
import keyboard as kb 
from pathlib import Path

planilha = pathlib.PureWindowsPath()
planilha_nome = ""
planilha_caminho = "" #Pasta do bot

def registrar(p, l):
  if p:
    global planilha, planilha_nome
    planilha_nome = input("Digite o nome do arquivo: ")
    planilha = pathlib.PureWindowsPath((os.path.join(planilha_caminho, planilha_nome)))  #Nome do arquivo
    planilha = planilha.as_posix()
    if not pathlib.Path(planilha).exists() and pathlib.PurePosixPath(planilha).suffix == ".xlsx":
      print("Digite um arquivo valido")
      registrar(True, True)
    else:
      global bd
      bd = pandas.read_excel(planilha)
  if l:
    global iniciar
    iniciar = int(input('Digite a linha para começar, 2 para começar no inicio: '))
    while iniciar < 2 or iniciar == "" or iniciar > len(bd) + 1:
      print("Digite um valor entre 2 - ", len(bd) + 1)
      iniciar = int(input('Digite a linha para começar, 2 para começar no inicio: '))
    iniciar = iniciar - 2
  while not kb.is_pressed('esc'):
      print ("Preparado, só apertar <Esc> para liberar")
      time.sleep(1)
  print("Começando em 3 segs")
  time.sleep(3)


gui.PAUSE = (0.1)
lanc = 0

def preencher_campos(hora1,hora2):
  global lanc 
  gui.hotkey("ctrl", "i")
  gui.write(bd.loc[iniciar, "lanc_colab"]) #funcionario
  gui.press("tab")

  pyperclip.copy(desc)
  gui.hotkey("ctrl","v") #observação
  gui.press("tab")

  pyperclip.copy(data)
  gui.hotkey("ctrl","v") #data prevista
  gui.press("enter")

  gui.hotkey("ctrl","v") #data exec
  gui.press("enter")


  gui.write(str(hora1)) #hora ini
  gui.press("enter")

  gui.write(str(hora2)) #hora fim

  gui.hotkey("ctrl", "o")
  time.sleep(0.2) 
  gui.press("enter")
  time.sleep(0.1)
  lanc += 1

def finalizar():
  fim = timeit.default_timer()
  print("-Finalizado-\nTotal de lançamentos:", lanc, "\nLinhas concluidas:",iniciar+1,"/", len(bd)+1,"\nTempo percorrido:" , (fim - inicio) / 60 , "mins\nMedia de lançamentos" , lanc / ((fim - inicio) / 60),"/min")
  fechar = input("Aperte <enter> para fechar e renomear o arquivo")
  os.rename(planilha, planilha_caminho + "LANÇADO - " + planilha_nome)

registrar(True,True)
inicio = timeit.default_timer()

while iniciar < len(bd):
  if kb.is_pressed('esc'):
    finalizar()
  data = bd.loc[iniciar, "lanc_data"]
  data = str(data.strftime('%d/%m/%Y'))

  if not pandas.isna(bd.loc[iniciar, "lanc_descricao"]):
    desc = bd.loc[iniciar, "lanc_descricao"].replace("_x000D_", " ")
  else:
    desc = " "
  if not kb.is_pressed('esc'):
    preencher_campos(bd.loc[iniciar, "lanc_horaini1"],bd.loc[iniciar, "lanc_horafim1"])
  if not pandas.isna(bd.loc[iniciar, "lanc_horaini2"]) and not kb.is_pressed('esc'): #Verifica se tem lançamento 2
    preencher_campos(bd.loc[iniciar, "lanc_horaini2"],bd.loc[iniciar, "lanc_horafim2"])
  if not pandas.isna(bd.loc[iniciar, "lanc_horaini3"]) and not kb.is_pressed('esc'): #Verifica se tem lançamento 3
    preencher_campos(bd.loc[iniciar, "lanc_horaini3"],bd.loc[iniciar, "lanc_horafim3"])
  print (iniciar+2,"/", len(bd)+1, "concluidos")
  iniciar += 1
  if kb.is_pressed("esc"):
    break
finalizar()


