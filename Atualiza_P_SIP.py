import pandas as pd
import pyautogui
import clipboard
import time
import os

delay = 0.8
#path_arq = os.listdir('C:\Users\103901\Desktop\py\dados')

def try_locate_img_2times(key): #Tenta encontrar a imagem 2 vezes

    import time
    import pyautogui
    try:

        img = pyautogui.locateCenterOnScreen(key, confidence=0.7)
        pyautogui.click(img.x, img.y)

    except:
        time.sleep(1.8)
        img = pyautogui.locateCenterOnScreen(key, confidence=0.5)
        pyautogui.click(img.x, img.y)
        print('ok')


def try_locate_img_1times(key): #Tenta encontrar a imagem 1 vezes

    import time
    import pyautogui
    try:
        img = pyautogui.locateCenterOnScreen(key, confidence=0.7)
        pyautogui.click(img.x, img.y)
        time.sleep(0.8)
    except:
        time.sleep(0.1)


def repeat_key(key, times): #Repete apertar tecla n vezes
    for _ in range(times):
        pyautogui.press(key)


def finaliza_tir_add(vlr_tot): #Ao chegar na última linha, finaliza tir
    pyautogui.press('F2')
    time.sleep(1.3 * delay)
    pyautogui.press('enter')
    time.sleep(1.3 * delay)
    print(vlr_tot)
    clipboard.copy(vlr_tot)

    pyautogui.hotkey('ctrl', 'v')
    time.sleep(1 * delay)

    repeat_key('enter', 5)

    pyautogui.press('F2')
    time.sleep(1 * delay)

    img = pyautogui.locateCenterOnScreen(
        'ORC_oktir.PNG', confidence=0.7)
    pyautogui.click(img.x, img.y)

    time.sleep(2.3 * delay)

    img = pyautogui.locateCenterOnScreen(
        'ORC_sair.PNG', confidence=0.7)
    pyautogui.click(img.x, img.y)

    time.sleep(1.8 * delay)


def if_OS_already_close():

    import time
    import pyautogui
    try:
        img = pyautogui.locateCenterOnScreen('ORC_ja_fechado.PNG', confidence=0.7)
        time.sleep(0.8*delay)
        try_locate_img_2times('ORC_ok2.PNG') #Clica no ok
        time.sleep(0.8*delay)
    except:
        time.sleep(0.1)


def atualizar_precos_SIP():
    df = pd.read_excel('orçamentospy.xlsx')
    row = 0
    num_rows = int(len(df['OS']))

    while row <= num_rows:  #Percorre cada linha da tabela
        n_OS = int(df.loc[row, 'OS']) #identifica o n da OS
        print(n_OS)
        clipboard.copy(n_OS) #copia o n da OS
        time.sleep(2 * delay) #espera
        

        #Com o n da OS copiado, abrir o orçamento
        try_locate_img_2times('ORC_OS.PNG') #Clica no campo da OS
        pyautogui.hotkey('ctrl', 'left') 
        
        time.sleep(0.2 * delay)
        repeat_key('delete', 8) #Delete inf da OS anterior
        pyautogui.hotkey('ctrl', 'v') # Cola OS nova
        time.sleep(1.8 * delay)
        repeat_key('enter', 2) #Repete enter 2x
        time.sleep(7 * delay)

        #Após abrir orçamento, clica no lucro para aplicar valor
        try_locate_img_2times('ORC_L%.PNG') #Clica no campo do lucro
        time.sleep(1.0 * delay)
        pyautogui.press('F2') #Aperta F2
        time.sleep(1.0 * delay)

        if_OS_already_close()

        pyautogui.press('enter') #Desce para linha do v. prazo
        time.sleep(1.0 * delay)
        vlr_tot = str(df.loc[row, 'Total']).replace('.', ',') #encontra o valor total
        print(vlr_tot)
        clipboard.copy(vlr_tot) #Copia vlr total
        pyautogui.hotkey('ctrl', 'v') #Cola vlr total
        time.sleep(1 * delay)
        repeat_key('enter', 5) #Dá enter 5 vezes
        try_locate_img_2times('ORC_ok_F2.PNG') #Clica no ok do botão F2
        time.sleep(0.5 * delay)
        pyautogui.press('F2') #Sai da tela de aplicar lucro
        time.sleep(1 * delay)

        #Salva a alteração de preço na 1° tiragem
        repeat_key('enter', 3) #Dá 3 enter para gravar preço
        time.sleep(1.0 * delay)
        if_OS_already_close()


        try_locate_img_2times('ORC_ok.PNG') #Dá enter para salvar
        if_OS_already_close()
        time.sleep(0.7 * delay)
        try_locate_img_1times('ORC_ok2.PNG')

        time.sleep(2.0 * delay)

        try_locate_img_2times('ORC_ok.PNG') #Clica no ok
        if_OS_already_close()
        time.sleep(0.7 * delay)
        try_locate_img_1times('ORC_ok2.PNG') #Clica no ok
        time.sleep(0.5 * delay)

        #Se não tiver tiragem adicional, muda para a proxima OS
        if row > num_rows:#Para de rodar se última linha
            break

        #Se tiver tiragem adicional, aplica os valores a eles
        if int(df.loc[row + 1, 'OS']) == int(df.loc[row, 'OS']):
            pyautogui.press('F7') #Abre tela de tir_add
            time.sleep(2 * delay)
            try_locate_img_2times('ORC_tiragem.PNG') #classifica as tir do menor para maior
            time.sleep(0.7 * delay)
            pyautogui.click()
            time.sleep(0.7 * delay)
            pyautogui.click()
            time.sleep(0.7 * delay)
            t=0
            if int(df.loc[row + 1, 'OS']) == int(df.loc[row, 'OS']):
                row = row + 1
                #Enquanto for a mesma OS, continua aplicando valor nas tir add
                while int(df.loc[row + 1, 'OS']) == int(df.loc[row, 'OS']):
                    pyautogui.press('F2') #Abre tela de digitar valor
                    time.sleep(2 * delay)
                    pyautogui.press('enter') #desce para valor a prazo
                    time.sleep(1 * delay)
                    vlr_tot = str(df.loc[row, 'Total']).replace('.', ',')
                    print(vlr_tot)
                    clipboard.copy(vlr_tot) #Copia valor total da planilha
                    pyautogui.hotkey('ctrl', 'v') #Cola valor total no SIP
                    time.sleep(0.7 * delay)
                    repeat_key('enter', 5) #Dá enter para aplicar o valor
                    #try_locate_img_2times('ok_F2.PNG') #Clica no ok do botão F2
                    pyautogui.press('F2') #Sai da tela de aplicar valor
                    t=t+1 #conta quantas linhas já foram feitas
                    time.sleep(2.5 * delay)
                    repeat_key('down',t) #Desce a linha dentro do SIP
                    row = row + 1 #Vai para a proxima linha da planilha

                    # Ao chegar na última linha da tabela, executa finaliza_tir_add
                    if row + 1 >= num_rows or int(df.loc[row + 1, 'OS']) != int(df.loc[row, 'OS']):
                        finaliza_tir_add(int(df.loc[row, 'Total']))
                        row = row + 1
                        break


        if int(df.loc[row + 1, 'OS']) != int(df.loc[row, 'OS']):
            row = row + 1



if __name__ == '__main__':
    atualizar_precos_SIP()
