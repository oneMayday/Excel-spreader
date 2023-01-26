import customtkinter
import tkinter

from tkinter import messagebox
from PIL import Image
from spreader import create_new_file, create_result_file


class App(customtkinter.CTk):
    """ Базовый класс, описывающий окно приложения и его функционал"""

    def __init__(self, main_func, create_additional_file_func):
        """ Инициализатор дополнительно принимает основную рабочую функцию и функцию создания дополнительного файла """

        super().__init__()
        self.main_func = main_func
        self.create_additional_file_fun = create_additional_file_func

        customtkinter.set_appearance_mode("dark-blue")
        customtkinter.set_default_color_theme("dark-blue")

        # Размеры окна, заголовок и блокировка изменения размеров окна
        self.geometry("400x280")
        self.title("Spreader")
        self.resizable(False, False)

        # Вставка логотипа
        self.logo = customtkinter.CTkImage(dark_image=Image.open("app/spreader.png"), size=(400, 80))
        self.logo_label = customtkinter.CTkLabel(master=self, text='', image=self.logo, anchor=tkinter.CENTER)
        self.logo_label.pack()

        # Фрейм с интерфейсом
        self.main_frame = customtkinter.CTkFrame(master=self, width=350, height=180, corner_radius=10)
        self.main_frame.place(relx=0.5, rely=0.6, anchor=tkinter.CENTER)

        # Поля ввода имени файла и стартовой ячейки
        self.entry_file_name = customtkinter.CTkEntry(master=self.main_frame, width=300,
                                                      placeholder_text='Введите имя файла', font=('Roboto', 12))
        self.entry_file_name.place(x=20, y=20)

        self.entry_start_cell = customtkinter.CTkEntry(master=self.main_frame, width=300,
                                                       placeholder_text='Введите адрес первой ячейки с SAP кодом',
                                                       font=('Roboto', 12))
        self.entry_start_cell.place(x=20, y=60)

        # Создаем чекбокс для создания сводной таблицы результатов
        self.check_creating = tkinter.IntVar(value=1)
        self.create_additional_file = customtkinter.CTkCheckBox(master=self.main_frame,
                                                                text='Создать сводный файл с результатами',
                                                                font=('Roboto', 12), variable=self.check_creating)
        self.create_additional_file.place(x=20, y=100)

        # Кнопка старта программы
        start_program = customtkinter.CTkButton(master=self.main_frame, command=self.start_main, text='Начать')
        start_program.place(relx=0.5, y=150, anchor=tkinter.CENTER)

    def start_main(self):
        """ Функция выполняется при нажатии на кнопку. Внутри происходит проверка введенных полей и их обработка.
        Если поля заполнены верно, выполняется основная рабочая и вспомогательная (если проставлен чекбокс) функции
        """

        file_name = self.entry_file_name.get()
        start_cell = self.entry_start_cell.get()
        check_result_table = self.check_creating.get()

        if not file_name or not start_cell:
            messagebox.showwarning(title='Ошибка', message="Введите имя файла и адрес ячейки!")
            return

        try:
            result_values_list, result_calculated_values = self.main_func(file_name, start_cell)
            create_new_file(file_name, result_values_list)
        except Exception:
            messagebox.showerror(title='Что-то пошло не так!', message="Проверьте имя файла и адрес ячейки!")
            return

        if check_result_table:
            create_result_file(result_calculated_values)

        messagebox.showinfo(title='Все прошло удачно', message="Готово!")
