<h2> Excel cell spreader</h2>

- Приложение предназначено для распределения максимальных значений в таблице между другими ячейками.


- Приложение подходит только для таблиц определенного вида, информацию о структуре таблицы можно получить, нажав кнопку "help" в окне программы



Для сборки приложения в exe-файл можно использовать pyinstaller (актуально для версии 5.7.0).
В терминале нужно ввести следующую команду:

pyinstaller --noconfirm --onefile -n spreader --windowed --icon=resources\fav.ico --add-data "resources\spreader.png;." --add-data "PATH_TO_YOUR_PROJECT\venv\lib\site-packages\customtkinter;customtkinter/" main.py 

Скрин интерфейса:





[![spreader-window.png](https://i.postimg.cc/N0gHxdvn/spreader-window.png)](https://postimg.cc/SnvRS7sL)
