import time, pyautogui, clipboard
import pandas as pd
from tkinter import *
from tkinter import filedialog

# imagens que precisam ser localizadas
add = 'img/add.png'
error = 'img/erro.png'

# selecionar excel para execução
def excel():
    global directory
    excelFile = filedialog.askopenfile(mode='rb', filetypes=[('Excel Files', '*.xlsx')])
    if excel:
        directory = excelFile.name
        fileText.config(text=directory)

# ler velocidade selecionada no optionMenu para envio de mensagens
def speedSelection():
    global timeValue
    option = timeOption.get()
    if option == "3":
        timeValue = 3
    elif option == "5":
        timeValue = 5
    elif option == "10":
        timeValue = 10
    elif option == "15":
        timeValue = 15
    elif option == "25":
        timeValue = 25
    elif option == "30":
        timeValue = 30
    else:
        timeValue = 0

def start():
    speedSelection()
    # verificações se existe excel e o tempo selecionado entre os envios
    if excel:
        if timeValue != 0:
            df = pd.read_excel(directory, sheet_name=0)
            row = len(df) # quantidade de linhas
            time.sleep(3)
            for i in range(row):
                time.sleep(timeValue)
                loc_add = pyautogui.locateOnScreen(add)
                if loc_add is not None:
                    x, y, width, height = loc_add
                    centerx_add = x + width // 2
                    centery_add = y + height // 2
                    pyautogui.click(centerx_add, centery_add)

                    clipboard.copy(df.loc[i, 'Telefone'])
                    pyautogui.hotkey('ctrl', 'v')
                    pyautogui.press("enter")

                    # tempo para tela atualizar e verificar se o número é valido
                    time.sleep(1.5)
                    loc_error = pyautogui.locateOnScreen(error)

                    if loc_error is not None:
                        pyautogui.press("esc")
                        print("Erro na linha {}".format(i))
                    else:
                        print("Mensagem enviada para o número {}".format(df.loc[i, 'Telefone']))
                        clipboard.copy(df.loc[i, 'Texto'])
                        pyautogui.hotkey('ctrl', 'v')
                        pyautogui.press("enter")
                else:
                    print(i)
        else:
            print("Selecione um delay")

root = Tk()
root.title("Robô Whatsapp")
root.geometry("315x135")

timeOption = StringVar()
timeOption.set("Tempo")

fileText = Label(root, text="Excel importado: NULL")
fileText.grid(column=0, row=0, padx=5, pady=7)

fileBtn = Button(root, text="Selecionar", command=excel)
fileBtn.grid(column=1, row=0, padx=5, pady=7)

speed = Label(root, text="Tempo de envio em segundos")
speed.grid(column=0, row=1, padx=5, pady=7)

opSpeed = OptionMenu(root, timeOption, "3","5", "10", "15", "25", "30")
opSpeed.grid(column=1, row=1, padx=5, pady=7)

startBtn = Button(root, text="Iniciar", command=start)
startBtn.grid(column=1, row=2, padx=5, pady=7)

root.mainloop()