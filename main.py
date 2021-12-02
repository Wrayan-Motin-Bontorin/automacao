import pyautogui as p
import time
import win32clipboard


numex = int(input("insira o numero de execucoes do script: "))

p.alert(text='ESTE PROGRAMA TEM INICIO NA PAGINA DE ITEMS CADASTRADOS!\nPOR FAVOR DEIXE NESSA PAGINA PARA O FUNCIONAMENTO CORRETO DO SCRIPT!', title='AVISO IMPORTANTE', button='OK')
if p.confirm(text='Deseja executar o programa', title='executar?', buttons=["OK", 'cancel']) == "OK":
    p.alert(text='10 SEGUNDO PRA EXECUCAO')
    time.sleep(15)
    while numex > 1:
    # entrando na pagina de items ja cadastrados
        p.press("tab")
        time.sleep(1)
        p.press("space")
        time.sleep(3)
    # entrando para cadastro de novo item
        p.press("insert")
        time.sleep(2)
    # copiando valor da tabela
        p.hotkey('alt', 'tab')
        time.sleep(1)
        p.hotkey('ctrl', 'c')
        time.sleep(1)
        p.hotkey('alt', 'tab')
        time.sleep(1)
        p.hotkey('ctrl', 'v')

        # alternativa caso nao funcione o ctr+V

        # win32clipboard.OpenClipboard()
        # name = win32clipboard.GetClipboardData()
        # win32clipboard.CloseClipboard()
        # p.write(name)
        time.sleep(3)
        p.press('tab')
        p.press('tab')
        time.sleep(1)
        p.press('u')
        p.hotkey('alt', 'return')
        time.sleep(5)
        p.press('esk')
        time.sleep(1)
        p.press('return')
    # cadastrando quantidade
        p.press('tab')
        time.sleep(1)
        p.hotkey('alt', 'tab')
        time.sleep(1)
        p.press('right')
        time.sleep(1)
        p.hotkey('ctrl', 'c')
    # pega a quantidade
        win32clipboard.OpenClipboard()
        qtde = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        p.hotkey('alt', 'tab')
        p.write(qtde)
        time.sleep(1)
        p.press('tab')
        time.sleep(1)
    # pegar o preco do item
        p.hotkey('alt', 'tab')
        time.sleep(1)
        p.press('right')
        time.sleep(1)
        p.hotkey('ctrl', 'c')
    # pega o valor de price da area de transferencia
        win32clipboard.OpenClipboard()
        price = win32clipboard.GetClipboardData()
        win32clipboard.CloseClipboard()
        time.sleep(1)
    # cadastro do preco
        p.hotkey('alt', 'tab')
        time.sleep(1)
        p.write(price)
        time.sleep(1)
        p.press('tab')
        time.sleep(1)
        p.hotkey('alt', 'return')
        time.sleep(2)
    # reposicionando a planilha
        p.hotkey('alt', 'tab')
        time.sleep(0.5)
        p.press('left')
        time.sleep(0.5)
        p.press('left')
        time.sleep(0.5)
        p.press('down')
        time.sleep(0.5)
        p.hotkey('alt', 'tab')
        numex = numex - 1
        print(numex)


