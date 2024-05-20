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
planilha_caminho = "Z:/Obras/2024/Planejamento de Obras/Revisão de Lançamentos/LAB/Bot Python/" #Pasta do bot
global data
global bd
global iniciar

gui.PAUSE = (0.1)
lanc = 0

time.sleep(2)

def tick():
  iniciar = 0
  while iniciar < len(bd): # Para cada linha enquanto o iniciar(Codigo para definir a linha) for menor que o tamanho do banco de dados
    if kb.is_pressed('esc'):
      finalizar()
      print(input())
    data = bd.loc[iniciar, "lanc_data"]
    data = str(data.strftime('%d/%m/%Y')) #Conversão de dado para o tipo do campo
    if not kb.is_pressed('esc'):
      preencher_campos(bd.loc[iniciar, "lanc_horaini1"],bd.loc[iniciar, "lanc_horafim1"],iniciar,data)
    if not pandas.isna(bd.loc[iniciar, "lanc_horaini2"]) and not kb.is_pressed('esc'): #Verifica se tem lançamento 2
      preencher_campos(bd.loc[iniciar, "lanc_horaini2"],bd.loc[iniciar, "lanc_horafim2"],iniciar,data)
    if not pandas.isna(bd.loc[iniciar, "lanc_horaini3"]) and not kb.is_pressed('esc'): #Verifica se tem lançamento 3
      preencher_campos(bd.loc[iniciar, "lanc_horaini3"],bd.loc[iniciar, "lanc_horafim3"],iniciar,data)
    print (iniciar+2,"/", len(bd)+1, "concluidos")
    iniciar += 1
    if kb.is_pressed("esc"):
      break
  finalizar(iniciar)

def preencher_campos(hora1,hora2,iniciar,data):
  global lanc
  gui.hotkey("ctrl", "i")
  gui.write(bd.loc[iniciar, "lanc_colab"]) #funcionario
  gui.press("tab")
  if not pandas.isna(bd.loc[iniciar, "lanc_descricao"]): #Verifica se há informação de descrição na linha do banco de dado
    pyperclip.copy(bd.loc[iniciar, "lanc_descricao"].replace("_x000D_", " ")) #Remove um codigo de quebra de linha do excel
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
  time.sleep(0.25) 
  gui.press("enter")
  time.sleep(0.15)
  lanc += 1

def finalizar(iniciar):
  fim = timeit.default_timer()
  print("-Finalizado-\nTotal de lançamentos:", lanc, "\nLinhas concluidas:",iniciar+1,"/", len(bd)+1,"\nTempo percorrido:" , (fim - inicio) / 60 , "mins\nMedia de lançamentos" , lanc / ((fim - inicio) / 60),"/min")
  gui.press("f9")
  time.sleep(0.5)
  

def obra(a): #Delega a obra antes de fazer  os ticks de lançamento
  gui.press("tab", presses=3) 
  gui.press("enter")
  gui.write(a)
  gui.press("enter", presses=2)
  time.sleep(0.5)
  gui.press("enter")
  time.sleep(0.5)
  gui.moveTo(x=462, y=210)
  gui.click(x=462, y=230)
  time.sleep(0.5)

planilhas = Path(planilha_caminho)
for child in planilhas.iterdir(): # para todos os arquivos do caminhos 
  if child.suffix == ".xlsx" and child.name[:7] != "Lançado": #verifica se possui o .xlsx e se não tem lançado antes do nome
    bd = pandas.read_excel(child)
    inicio = timeit.default_timer()
    lanc = 0
    obra(child.stem)
    os.rename(child, str(child) + ".EM USO") #adiciona um sufixo .em uso, para impedir alguem de abrir o arquivo enquanto esta sendo lançado
    tick()
    os.rename(str(child) + ".EM USO", planilha_caminho + "Lançado - " + str(child.name)) #adiciona um prefixo de lançado

print("Finalizados com sucesso")