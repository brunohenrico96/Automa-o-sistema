import pandas as pd
import pyautogui
import clipboard
import time
from Atualiza_P_SIP import try_locate_img_1times, try_locate_img_2times, repeat_key

delay = 1

time.sleep(3)
def prop_aut():
    df = pd.read_excel('orçamentospy.xlsx')
    row = 0
    num_rows = int(len(df['OS']))
    tela = int(input('Qual a tela?'))
    env_email = int(input('Enviar email? 1 para SIM e 0 para NÃO'))
    alterar_obs = int(input('Alterar observação? 1 para SIM e 0 para NÃO'))
    while row <= num_rows:  #Percorre cada linha da tabela

        n_OS = int(df.loc[row, 'OS']) #identifica o n da OS
        print(n_OS)
        clipboard.copy(n_OS) #copia o n da OS
        time.sleep(2 * delay) #espera
     
        if tela == 4193:

            #Com o n da OS copiado, abrir o orçamento
            try_locate_img_2times('PRO_OS.PNG') #Clica no campo da OS
            pyautogui.hotkey('ctrl', 'left') 
            time.sleep(0.2 * delay)
            repeat_key('delete', 8) #Delete inf da OS anterior
            pyautogui.hotkey('ctrl', 'v') # Cola OS nova
            time.sleep(1.8 * delay)
            repeat_key('enter', 1) #Repete enter 1x
            time.sleep(2 * delay)

        elif tela == 4141:

            #Com o n da OS copiado, abrir o orçamento
            try_locate_img_2times('PRO_OS4140.PNG') #Clica no campo da OS
            pyautogui.hotkey('ctrl', 'left') 
            time.sleep(0.2 * delay)
            repeat_key('delete', 8) #Delete inf da OS anterior
            pyautogui.hotkey('ctrl', 'v') # Cola OS nova
            time.sleep(1.8 * delay)
            repeat_key('enter', 1) #Repete enter 1x
            time.sleep(2 * delay)

        #clicar em obs
        if alterar_obs == 1:
            img = pyautogui.locateCenterOnScreen('PRO_a.PNG', confidence=0.7)
            pyautogui.moveTo(img.x, img.y)
            pyautogui.click()
            time.sleep(1)
            img = pyautogui.locateCenterOnScreen('PRO_b.PNG', confidence=0.4)
            pyautogui.dragTo(img.x, img.y, duration=0.5)
            repeat_key('backspace', 25)

            #digitar obs
            obs = df.loc[row, u'obs']
            print(obs)
            clipboard.copy(obs)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(1)
        time.sleep(0.5)
        #clicar em ok para salvar
        try_locate_img_1times('PRO_ok.PNG')           
        time.sleep(1)
        #clicar em ok2 para salvar
        try_locate_img_1times('PRO_ok2.PNG') 
        time.sleep(1)        
        #apertar F9
        pyautogui.press('F9')
        time.sleep(2)
        #Apertar ESC para sair da tela de comentario
        pyautogui.press('ESC')  
        time.sleep(1.5)

        if env_email == 1:
            try_locate_img_2times('PRO_enviar_email.PNG')
            time.sleep(1)
            pyautogui.press('up')  
            time.sleep(0.3)
            pyautogui.press('enter')
            time.sleep(0.3)
        

      
        #clicar em env
        time.sleep(0.3)        
        try_locate_img_2times('PRO_env.PNG')      
        #esperar 3 segundos
        time.sleep(2.6) 
             
        #clicar em imp
        time.sleep(1)
        try_locate_img_2times('PRO_imp.PNG')          
        #esperar 15 segundos
        time.sleep(20)
        #clicar em salvar
        try_locate_img_2times('PRO_tela_salvar.PNG')
        time.sleep(0.4)
        try:
            try_locate_img_2times('PRO_pesq_local.PNG')
            pyautogui.write('C:\Proposta')
            time.sleep(0.4)
            pyautogui.press('enter')     
            time.sleep(0.4)    
        except:
            print('Error')
        try_locate_img_1times('PRO_salvar.PNG')  
        try_locate_img_1times('PRO_salvar2.PNG')  
        #esperar 2 segundos
        time.sleep(2)        
        #clicar em salvar_sim
        try_locate_img_2times('PRO_salvar_sim.PNG')
        time.sleep(1)      
        #Apertar ESC para sair da tela de proposta
        if tela == 4193:
            pyautogui.press('ESC')
        #fim       
        time.sleep(2) 






      
        if int(df.loc[row + 1, 'OS']) != int(df.loc[row, 'OS']):
            row = row + 1

  

if __name__ == '__main__':
    prop_aut()