import time, pyautogui, clipboard, sys, threading
import pandas as pd
from tkinter import *
from tkinter import filedialog

# Localização de imagem
add = 'img/add.png'
error = 'img/erro.png'

# Interrupção de programa
stopOp = False

# Tipo de execução selecionada
def execType():
    global typeSelec
    typeSelec = var.get()

# Excel selecionado
def excel():
    global directory
    excelFile = filedialog.askopenfile(mode='rb', filetypes=[('Excel Files', '*.xlsx')])
    if excelFile:
        directory = excelFile.name
        fileText.config(text=directory)

# Velocidade de envio selecionado
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

# Definindo a saída para o console dentro da inteface
def consoleText(widget):
    class output:
        def __init__(self, widget):
            self.widget = widget

        def write(self, text):
            self.widget.config(state=NORMAL)
            self.widget.insert(END, text)
            self.widget.config(state=DISABLED)
            self.widget.see(END)

    sys.stdout = output(widget)
    sys.stderr = output(widget)

# Execução da função principal
def start():
    speedSelection()
    execType()
    # Verificação de Excel e capturando o número de linhas
    if directory:
        if timeValue != 0:
            df = pd.read_excel(directory, sheet_name=0)
            row = len(df)
            time.sleep(3)

            def runApp():
                print("..:Iniciando:..")
                # Definindo váriaveis para estátistica
                global stopOp, numImg, numSuc, numIn
                stopOp = True
                numImg = 0
                numSuc = 0
                numIn = 0
                # Executando conforme o método selecionado
                if(typeSelec == '1'):
                    for i in range(row):    
                        if not stopOp:
                            print("..:Interrompido:..")
                            break
                        else:
                            time.sleep(timeValue)
                            loc_add = pyautogui.locateOnScreen(add)
                            if loc_add is not None:
                                x, y, width, height = loc_add
                                centerx_add = x + width // 2
                                centery_add = y + height // 2
                                pyautogui.click(centerx_add, centery_add)

                                clipboard.copy(str(df.iloc[i, 0]))
                                pyautogui.hotkey('ctrl', 'v')
                                pyautogui.press("enter")

                                time.sleep(1.5)
                                loc_error = pyautogui.locateOnScreen(error)

                                if loc_error is not None:
                                    pyautogui.press("esc")
                                    pyautogui.press("esc")
                                    print("Número inexistente: Linha {}".format(i+1))
                                    numIn = numIn + 1
                                else:
                                    print("Linha {} enviada com sucesso.".format(i+1))
                                    clipboard.copy(str(df.iloc[i, 1]))
                                    pyautogui.hotkey('ctrl', 'v')
                                    pyautogui.press("enter")
                                    numSuc = numSuc + 1
                            else:
                                print("Imagem não reconhecida. Linha {} não enviada.".format(i+1))
                                numImg = numImg + 1

                    print("\n")
                    print(".....:RESUMO:.....")
                    print("TOTAL: {} Linhas".format(i+1))
                    print("{} Enviadas".format(numSuc))
                    print("{} Inexistente(s)".format(numIn))
                    print("{} Falha na imagem".format(numImg))
                    print("..................")
                elif(typeSelec == '2'):
                    for i in range(row):     
                        if not stopOp:
                            print("..:Interrompido:..")
                            break
                        else:
                            time.sleep(timeValue)

                            pyautogui.hotkey('ctrl', 'alt', 's')
                            time.sleep(0.5)
                            clipboard.copy(str(df.iloc[i, 0]))
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press("enter")

                            time.sleep(1)
                            clipboard.copy(str(df.iloc[i, 1]))
                            pyautogui.hotkey('ctrl', 'v')
                            pyautogui.press("enter")

                            time.sleep(1)
                            pyautogui.press("esc")
                            pyautogui.press("esc")

                else:
                    print("Método não selecionado")
                
            run_thread = threading.Thread(target=runApp)
            run_thread.start()
        else:
            print("Selecione um delay")

# Parando a execução do programa
def stopApp():
    global stopOp
    stopOp = False

# Configurações da interface
root = Tk()
root.title("Robô Whatsapp")

# Capturando a seleção
timeOption = StringVar()
timeOption.set("Tempo")
var = StringVar()

execLabel = Label(root, text="Método de funcionamento:")
execLabel.grid(column=0, row=0, padx=5, pady=7)

execOption = Radiobutton(root, text="Imagem (Recomendado)", variable=var,  value=1)
execOption.grid(column=1, row= 0, padx=5, pady=7)

execOption2 = Radiobutton(root, text="Atalho", variable=var,  value=2) 
execOption2.grid(column=2, row= 0, padx=5, pady=7) 

fileText = Label(root, text="Excel importado: NULL")
fileText.grid(column=0, row=1, padx=5, pady=7)

fileBtn = Button(root, text="Selecionar", command=excel)
fileBtn.grid(column=1, row=1, columnspan=2, padx=5, pady=7)

speed = Label(root, text="Tempo de envio em segundos")
speed.grid(column=0, row=2, padx=5, pady=7)

opSpeed = OptionMenu(root, timeOption, "3","5", "10", "15", "25", "30")
opSpeed.grid(column=1, row=2, columnspan=2, padx=5, pady=7)

textConsole = Label(root, text="Console")
textConsole.grid(column=0, row=3, columnspan=3)

infoConsole = Text(root, wrap=WORD, height=10, width=50)
infoConsole.grid(column=0, row=4, columnspan=3, padx=5, pady=7)
infoConsole.config(state=DISABLED)

consoleText(infoConsole)

stopBtn = Button(root, text="Parar", command=stopApp)
stopBtn.grid(column=0, row=5, padx=5, pady=7)

startBtn = Button(root, text="Iniciar", command=start)
startBtn.grid(column=1, row=5, columnspan=2, padx=5, pady=7)

root.mainloop()