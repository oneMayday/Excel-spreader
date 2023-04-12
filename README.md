<h2> Excel cell spreader</h2>

- Приложение предназначено для распределения максимальных значений в таблице между другими ячейками.


- Приложение подходит только для таблиц определенного вида (подробнее можно прочитать в окне "help" программы.



Для сборки приложения в exe-файл можно использовать pyinstaller (актуально для версии 5.7.0).
В терминале нужно ввести следующую команду:

    Для Windows:
    pyinstaller --noconfirm --onefile -n spreader --windowed --icon=resources/fav.ico --add-data "resources/spreader.png;." --add-data "venv/lib/site-packages/customtkinter;customtkinter/" main.py

    Для Линукс:
    pyinstaller --noconfirm --onefile -n spreader --windowed --icon=resources/fav.ico --add-data "resources/spreader.png:." --add-data "venv/lib/site-packages/customtkinter:customtkinter/" main.py

Если сборка с первого раза не получилась - нужно внимательно посмотреть пути до папки customtkinter, возможно в путь нужно добавить дополнителую ветку 'python'
Скрин интерфейса:





[![spreader-window.png](https://i.postimg.cc/N0gHxdvn/spreader-window.png)](https://postimg.cc/SnvRS7sL)
