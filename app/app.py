import customtkinter
import sys
import tkinter

from os import path
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

        def resource_path(relative_path):
            """ Получение абсолютного пути к файлу, для работы из IDE и при использовании pyinstaller """

            try:
                # Pyinstaller создает временный каталог в _MEIPASS
                base_path = sys._MEIPASS
                return path.join(base_path, relative_path[10:])
            except Exception:
                base_path = path.abspath(".")
                return path.join(base_path, relative_path)

        # Вставка логотипа
        self.logo = customtkinter.CTkImage(dark_image=Image.open(resource_path("resources/spreader.png")),
                                           size=(400, 80))
        self.logo_label = customtkinter.CTkLabel(master=self, text='', image=self.logo, anchor=tkinter.CENTER)
        self.logo_label.pack()

        # Фрейм с интерфейсом
        self.main_frame = customtkinter.CTkFrame(master=self, width=350, height=180, corner_radius=10)
        self.main_frame.place(relx=0.5, rely=0.6, anchor=tkinter.CENTER)

        # Поля ввода имени файла и стартовой ячейки
        self.entry_file_name = customtkinter.CTkEntry(master=self.main_frame, width=300,
                                                      placeholder_text='Введите имя файла', font=('Roboto', 12))
        self.entry_file_name.place(x=25, y=20)

        self.entry_start_cell = customtkinter.CTkEntry(master=self.main_frame, width=300,
                                                       placeholder_text='Введите адрес первой ячейки с SAP кодом',
                                                       font=('Roboto', 12))
        self.entry_start_cell.place(x=25, y=60)

        # Создаем чекбокс для создания сводной таблицы результатов
        self.check_creating = tkinter.IntVar(value=1)
        self.create_additional_file = customtkinter.CTkCheckBox(master=self.main_frame,
                                                                text='Создать сводный файл с результатами',
                                                                fg_color='#264ab5', font=('Roboto', 12),
                                                                variable=self.check_creating)
        self.create_additional_file.place(x=25, y=100)

        # Кнопка старта программы
        self.start_program = customtkinter.CTkButton(master=self.main_frame, fg_color='#264ab5',
                                                     command=self.start_main, text_color='#191a19', text='Старт')
        self.start_program.place(relx=0.5, y=150, anchor=tkinter.CENTER)

        # Кнопка информации
        self.help_info = customtkinter.CTkButton(master=self.main_frame, command=self.show_info,
                                                 text='Help', width=10,
                                                 text_color='#191a19', fg_color='#264ab5', font=('Roboto', 12))
        self.help_info.place(x=285, y=137)

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

    @staticmethod
    def show_info():
        """ Окно помощи """

        window_info = customtkinter.CTkToplevel()
        window_info.geometry("560x240")
        window_info.title("Help")
        window_info.resizable(False, False)

        window_info_frame = customtkinter.CTkTextbox(window_info)
        window_info_frame.place(relx=0.5, rely=0.5, relwidth=0.95, relheight=0.9, anchor=tkinter.CENTER)

        window_info_frame.insert('0.0', 'Для корректной работы программы необходимо:\n\n'
                                        '1) Вводить имя файла необходимо с расширением, например: "some_file.xlsx".\n'
                                        '2) В поле "Введите имя ячейки" необходимо ввести адрес первой ячейки,'
                                        ' в которой\n'
                                        '   находится первый SAP код первого товара.\n'
                                        '3) Для правильной работы программы исходная матрица должен иметь\n'
                                        '   следующую структуру:\n'
                                        '       - Первый столбец (А) должен содержать SAP-кода.\n'
                                        '       - Второй столбец (Б) должен содержать наименования товаров.\n'
                                        '       - Столбцы со значениями должны начинаться со столбца "C".\n'
                                        '       - Ячейки не должны содержать формул, только чистые значения.\n'
                                        '       - После последнего столбца со значениями не должно быть каких-либо\n '
                                        '         других значений.')
