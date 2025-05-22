import pandas as pd
import pyautogui
import keyboard
import time
import tkinter as tk
from threading import Thread

# Caminho da planilha
file_path = r''

# Lê os dados da aba "Base limpa"
df = pd.read_excel(file_path, sheet_name='Base limpa', engine='openpyxl')

# Inicializa o contador e o tempo inicial
contador = 0
start_time = time.time()

# Função para atualizar a janela externa
def update_window():
    while True:
        elapsed_time = time.time() - start_time
        time_label.config(text=f"Tempo decorrido: {int(elapsed_time)} segundos")
        count_label.config(text=f"Registros preenchidos: {contador}")
        time.sleep(1)

# Função para criar a janela externa
def criar_janela():
    global time_label, count_label
    janela = tk.Tk()
    janela.title("Status da Automação")
    janela.geometry("300x100")

    time_label = tk.Label(janela, text="Tempo decorrido: 0 segundos")
    time_label.pack()

    count_label = tk.Label(janela, text="Registros preenchidos: 0")
    count_label.pack()

    Thread(target=update_window, daemon=True).start()
    janela.mainloop()

# Inicia a janela em um thread separado
Thread(target=criar_janela, daemon=True).start()

print("O script começará em 10 segundos. Posicione o navegador com o formulário.")
print("Pressione ESC a qualquer momento para interromper a execução.")
time.sleep(10)

for index, row in df.iterrows():
    if keyboard.is_pressed('esc'):
        print("Execução interrompida pelo usuário.")
        break

    # 1. Clica no campo 'Usuário' para garantir o foco
    pyautogui.click(x=564, y=344)
    pyautogui.write(str(row['usuário']))
    pyautogui.press('tab')

    # 2. Seleciona 'Dealer'
    pyautogui.write('D')
    pyautogui.press('tab')

    # 3. Nome completo
    pyautogui.write(str(row['nome inteiro']))
    pyautogui.press('tab')

    # 4. E-mail
    pyautogui.write(str(row['e-mail']))
    pyautogui.press('tab')

    # 5. Clica no botão 'Salvar'
    pyautogui.click(x=963, y=664)

    contador += 1
    print(f"Registro {contador} preenchido com sucesso.")

    time.sleep(1)

print(f"Script finalizado. Total de registros preenchidos: {contador}")
