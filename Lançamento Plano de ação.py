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
    planilha_nome = input("Digite o nome do arquivo: ") #Nome do arquivo 
    planilha = pathlib.PureWindowsPath((os.path.join(planilha_caminho, planilha_nome))) #Caminho da planilha
    planilha = planilha.as_posix()
    if not pathlib.Path(planilha).exists() and pathlib.PurePosixPath(planilha).suffix == ".xlsx": #Verifica se o arquivo existe e se o sufixo está em .xlsx
      print("Digite um arquivo valido (.xlsx)")
      registrar(True, True)
    else:
      global bd
      bd = pandas.read_excel(planilha)
  if l:
    global iniciar
    iniciar = int(input('Digite a linha para começar, 2 para começar no inicio: '))
    while iniciar < 2 or iniciar == "" or iniciar > len(bd) + 1: #Se esta dentro do campo de 2 ao tamanho da planilha
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
  os.rename(planilha, planilha_caminho + "LANÇADO - " + planilha_nome) # Renomeia o arquivo para evitar lançamento duplicados

registrar(True,True)
inicio = timeit.default_timer()

while iniciar < len(bd): # Para cada linha enquanto o iniciar(Codigo para definir a linha) for menor que o tamanho do banco de dados
  if kb.is_pressed('esc'):
    finalizar()
    
  data = bd.loc[iniciar, "lanc_data"]
  data = str(data.strftime('%d/%m/%Y')) #Conversão de dado para o tipo do campo

  if not pandas.isna(bd.loc[iniciar, "lanc_descricao"]): #Verifica se há informação de descrição na linha do banco de dado
    desc = bd.loc[iniciar, "lanc_descricao"].replace("_x000D_", " ") #Remove um codigo de quebra de linha do excel
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


