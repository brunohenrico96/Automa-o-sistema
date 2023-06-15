import pandas as pd
import pyautogui
import clipboard
import time
from Atualiza_P_SIP import try_locate_img_1times, try_locate_img_2times, repeat_key

delay = 0.90




def abre_OS():
#Com o n da OS copiado, abrir o orçamento
    try_locate_img_2times('ASS_OS.PNG') #Clica no campo da OS
    pyautogui.hotkey('ctrl', 'left') 
    time.sleep(0.9 * delay)
    repeat_key('delete', 8) #Delete inf da OS anterior
    pyautogui.hotkey('ctrl', 'v') # Cola OS nova
    time.sleep(2.2 * delay)
    repeat_key('enter', 2) #Repete enter 2x
    time.sleep(1.3 * delay)
    try_locate_img_2times('ASS_add.PNG') #Clica no campo de add ass
    time.sleep(8 * delay)


def escolhe_local():
    try_locate_img_2times('ASS_local.PNG') #Clica no campo do local 
    time.sleep(1.0 * delay)
    pyautogui.hotkey('ctrl', 'v') # Cola local
    time.sleep(1.0 * delay)
    pyautogui.press('enter')
    time.sleep(1.0* delay)
    try_locate_img_2times('ASS_nome.PNG') #Clica no campo do nome 
    time.sleep(1.5 * delay)


def atualizar_assinatura():
    df = pd.read_excel('orçamentospy.xlsx')

    row = 0
    num_rows = int(len(df['OS']))
    
    while row <= num_rows:  #Percorre cada linha da tabela
        n_OS = int(df.loc[row, 'OS']) #identifica o n da OS
        print(n_OS)
        clipboard.copy(n_OS) #copia o n da OS
        time.sleep(2.0 * delay) #espera

        abre_OS()

        local = df.loc[row, 'local'] #identifica o local
        clipboard.copy(local) #copia o local

        if row == 0:
            escolhe_local()

        nome = str(df.loc[row, 'nome']) #identifica o nome
        print(nome)
        clipboard.copy(nome)
        pyautogui.hotkey('ctrl', 'v')
        #pyautogui.typewrite(nome)
        time.sleep(1.5 * delay)
        try_locate_img_2times('ASS_local_ok.PNG')
        #pyautogui.press('enter')
        time.sleep(1.5 * delay)
        pyautogui.typewrite('Autorizacao')
        time.sleep(0.9 * delay)
        try_locate_img_2times('ASS_ok_ass.PNG') #Clica no campo da OS 
        time.sleep(1.9 * delay)          
      
        if int(df.loc[row + 1, 'OS']) != int(df.loc[row, 'OS']):
            row = row + 1


if __name__ == '__main__':
    atualizar_assinatura()
