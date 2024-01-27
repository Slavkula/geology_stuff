import tkinter as tk
import subprocess

def open_card_script():
    subprocess.call(['python', 'карточки.py'])

window = tk.Tk()
window.title("Открытие скрипта")

open_card_button = tk.Button(window, text="Создание карточек", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'белые нули.py'])

open_card_button = tk.Button(window, text="Белые нули в ведомость", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'сделать csv.py'])

open_card_button = tk.Button(window, text="Сделать .csv из экселя", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'красные числа.py'])

open_card_button = tk.Button(window, text="Убрать красные числа", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'Заполнение ведомости.py'])

open_card_button = tk.Button(window, text="Заполнение ведомости", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'вкладки3.py'])

open_card_button = tk.Button(window, text="Создание вкладок3", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'карточки4.py'])

open_card_button = tk.Button(window, text="Создание карточек4", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'вкладки4.py'])

open_card_button = tk.Button(window, text="Создание вкладок4", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'карточки5.py'])

open_card_button = tk.Button(window, text="Создание карточек5", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'карточки5.py'])

open_card_button = tk.Button(window, text="Создание карточек5", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

def open_card_script():
    subprocess.call(['python', 'карточки6.py'])

open_card_button = tk.Button(window, text="Создание карточек6", command=open_card_script, font=('Arial', 20), width=30)
open_card_button.pack()

window.mainloop()
