import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showerror, showwarning, showinfo

import tools
import generate
  
root = tk.Tk()
root.eval('tk::PlaceWindow . center')

def helloCallBack():
    showinfo( "Hello Python", "Hello World")

def generate_students():
    generate.make_students_excel('students.xlsx')
    showinfo('Создание файла', 'Файл с логинами успешно создан. Название файла: students.xlsx')

def check_file():
    filetypes = (
        ('excel files', '*.xlsx'),
    )
    path = fd.askopenfilename(
        title='Выберите файл',
        filetypes=filetypes
    )
    print(path)

    fio = ''
    try:
        fio = tools.check_file(path)
        if fio:
            showinfo('Результат проверки', f'Файл корректный, ФИО студента: {fio}')
    except KeyError as e:
        showinfo('Результат проверки', 'Вероятно, передан файл, не сгенерированный программой')
    except Exception as e:
        showinfo('Результат проверки', 'Неизвестная ошибка. Вероятно, изменены свойства файла')


generate_students_button = tk.Button(root, text="Создать список логинов для студентов", command=generate_students)
check_file_button = tk.Button(root, text="Проверить файл студента", command=check_file)

generate_students_button.pack(side=tk.LEFT)
check_file_button.pack(side=tk.RIGHT)
root.mainloop()