import os
import time
import subprocess
import keyboard
import platform
import sys
import customtkinter
import requests
from bs4 import BeautifulSoup
import pyttsx3
from datetime import datetime
import webbrowser
from fuzzywuzzy import fuzz
import re
import speech_recognition as sr
from colorama import init, Fore, Style
import calendar
import pygetwindow as gw
import pyautogui
import threading
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk, ImageSequence
import ctypes
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL


# Инициализация colorama
init(autoreset=True)

# Инициализация блокировки для pyttsx3
speak_lock = threading.Lock()

# Настройки
opts = {
    "alias": ('миша', 'миш', 'мишь', 'мих', 'михаил', 'мишка', 'медведь', 'мишенька', "миша", "мишу"),
    "tbr": ('скажи', 'расскажи', 'покажи', 'сколько', 'произнеси',
            'ищи в гугл', 'найди в гугле', 'поиск google', 'найди в гугл', 'поиск в гугл',
            'ищи в google', 'найди в google',
            'найди википедия', 'поиск в википедии', 'что такое', 'найди гугл', 'в гугл',
            "скажи", "покажи", "скрой", "заверши", "удали","закрой окно"),
    "cmds": {
        "ctime": ('текущее время', 'сейчас времени', 'который час'),
        "radio": ('включи музыку', 'воспроизведи радио', 'включи радио', 'включи youtube', 'открой youtube'),
        "data": ('дата', 'какое сегодня число'),
        "stupid1": ('расскажи анекдот', 'рассмеши меня', 'ты знаешь анекдоты'),
        "google": ('открой гугл', 'запусти браузер', 'открой браузер'),
        "spasibo": ('спасибо', 'благодарю'),
        "kak_dela": ('как дела',),
        "steam": ('запусти steam', 'запусти стим', 'открой steam'),
        "ds": ('открой дискорд', 'открой дс', 'открой discord'),
        "yes": ('да', 'конечно', 'ага'),
        "no": ('нет', 'неа', 'отмена'),
        "dis": ('открой нагрузку', 'открой диспетчер задач'),
        "wikipedia": ('найди википедия', 'поиск в википедии', 'что такое', 'найди в википедии'),
        "open_window": ('открой окно', 'активируй окно', 'покажи окно'),
        "close_window": ('закрой окно', 'удали окно'),
        "minimize_window": ('сверни окно', 'убери окно'),
        "ang": ('английский', 'смени язык на английский', 'английский язык'),
        "rus": ('руский', 'смени язык на руский', 'руский язык'),
        "list_windows": ('покажи окна', 'какие окна открыты', 'список окон'),
        'change_volume': ('звук на ', 'измени звук на '),
        'googl_sk': ('поиск', 'ищи')
    }
}

# Инициализация pyttsx3
speak_engine = pyttsx3.init()
speak_engine.setProperty('voice', speak_engine.getProperty('voices')[3].id)

# Инициализация распознавателя речи
recognizer = sr.Recognizer()
microphone = sr.Microphone()

# Флаг для ожидания ответа на запрос подтверждения
awaiting_confirmation = False
recording = False
recording_file = os.path.join(os.getcwd(), 'записи.txt')


user32 = ctypes.WinDLL('user32', use_last_error=True)

# Печать текста зелёным цветом
def print_green(text):
    print(Fore.GREEN + text + Style.RESET_ALL)


def speak(what):
    print_green(what)
    with speak_lock:
        speak_engine.say(what)
        speak_engine.runAndWait()


def recognize_cmd(cmd):
    RC = {'cmd': '', 'percent': 0}
    for c, v in opts['cmds'].items():
        for x in v:
            vrt = fuzz.ratio(cmd, x)
            if vrt > RC['percent']:
                RC['cmd'] = c
                RC['percent'] = vrt

    if RC['percent'] > 50:
        return RC
    else:
        return {'cmd': '', 'percent': 0}


def list_windows():
    try:
        windows = gw.getAllTitles()
        active_windows = [w for w in windows if w.strip()]
        print(f"[log] Активные окна: {active_windows}")
        # speak("Активные окна: " + ", ".join(active_windows))
        return active_windows
    except Exception as e:
        speak(f"Произошла ошибка при получении списка окон: {e}")
        return []

def close_window(title):
    title = clean_command(title, window_cmd=True)
    print(f"[log] Закрытие окна: '{title}'")
    try:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            window = windows[0]
            if window.isMinimized:
                window.restore()
            if not window.isActive:
                window.activate()
            window.close()
            speak(f"Окно '{title}' закрыто.")
        else:
            speak(f"Окно с названием '{title}' не найдено.")
    except Exception as e:
        speak(f"Произошла ошибка при закрытии окна: {e}")


def open_window(title):
    title = clean_command(title, window_cmd=True)
    print(f"[log] Открытие окна: '{title}'")
    active_windows = list_windows()  # Получить список активных окон
    if not any(title.lower() in window.lower() for window in active_windows):
        speak(f"Окно с названием '{title}' не найдено.")
        return
    try:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            window = windows[0]
            if window.isMinimized:
                window.restore()
            window.activate()
            pyautogui.hotkey('alt', 'tab')
            if window.isActive:
                speak(f"Окно '{title}' открыто.")
            else:
                speak(f"Не удалось активировать окно '{title}'.")
        else:
            speak(f"Окно с названием '{title}' не найдено.")
    except Exception as e:
        speak(f"Произошла ошибка при открытии окна: {e}")

def minimize_window(title):
    title = clean_command(title, window_cmd=True)
    print(f"[log] Сворачивание окна: '{title}'")
    active_windows = list_windows()  # Получить список активных окон
    if not any(title.lower() in window.lower() for window in active_windows):
        speak(f"Окно с названием '{title}' не найдено.")
        return
    try:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            window = windows[0]
            if window.isMinimized:
                speak(f"Окно '{title}' уже свернуто.")
                return
            window.activate()
            time.sleep(0.5)  # Wait a bit to ensure window is activated
            window.minimize()
            if window.isMinimized:
                speak(f"Окно '{title}' свернуто.")
            else:
                speak(f"Не удалось свернуть окно '{title}'.")
        else:
            speak(f"Окно с названием '{title}' не найдено.")
    except Exception as e:
        speak(f"Произошла ошибка при сворачивании окна: {e}")


words_to_numbers = {
    'ноль': 0,
    'один': 1,
    'два': 2,
    'три': 3,
    'четыре': 4,
    'пять': 5,
    'шесть': 6,
    'семь': 7,
    'восемь': 8,
    'девять': 9,
    'десять': 10,
    'пятнадцать': 15,
    'двадцать': 20,
    'тридцать': 30,
    'сорок': 40,
    'пятьдесят': 50,
    'шестьдесят': 60,
    'семьдесят': 70,
    'восемьдесят': 80,
    'девяносто': 90,
    'сто': 100
}


def set_volume(volume_level):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = interface.QueryInterface(IAudioEndpointVolume)

    # Установите уровень громкости (от 0.0 до 1.0)
    volume.SetMasterVolumeLevelScalar(volume_level / 100, None)
    speak(f"Громкость установлена на {volume_level} процентов.")


def get_volume_from_voice_input(voice_input):
    # Сначала попытаемся найти числовые значения
    match = re.search(r'\d+', voice_input)
    if match:
        return int(match.group())

    # Если не найдено, попробуем найти словесные числа
    for word, number in words_to_numbers.items():
        if word in voice_input:
            return number

    # Если ничего не найдено, возвращаем None
    return None


# def execute_cmd(cmd, voice_input=None):
#     global awaiting_confirmation, recording
#     print_green(f"[log] Выполнение команды: {cmd}")
#     commands = {
#         'ctime': lambda: speak(f"Сейчас {datetime.now().strftime('%H:%M')}"),
#         'data': data,
#         'radio': lambda: webbrowser.open("https://www.youtube.com/watch?v=7x3RW2BCG68&list=PL-SnCU0L3jIkKCh8o3W9hY8qTmHEYTtpt"),
#         'dis': lambda: os.system('taskmgr'),
#         'google': lambda: webbrowser.open_new('chrome'),
#         'spasibo': lambda: speak("Всегда пожалуйста"),
#         'kak_dela': lambda: speak("Все хорошо, спасибо что спросили"),
#         'steam': open_steam,
#         'ds': open_discord,
#         'wikipedia': lambda: speak(get_wikipedia_summary(voice_input)),
#         'open_window': lambda: open_window(clean_command(voice_input, window_cmd=True)),
#         'close_window': lambda: close_window(clean_command(voice_input, window_cmd=True)),
#         'minimize_window': lambda: minimize_window(clean_command(voice_input, window_cmd=True)),
#         'list_windows': list_windows,
#         'rus': switch_to_russian,
#         'ang': switch_to_english,
#         'change_volume': lambda: set_volume(get_volume_from_voice_input(voice_input)),
#         'googl_sk': search_google
#     }
#
#     if cmd in commands:
#         commands[cmd]()
#     else:
#         print("Извините, я не понимаю эту команду.")



def search_google():
    recongnizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Говорите...")
        audio = recongnizer.listen(source)
        try:
            print("Поиск...")
            query = recongnizer.recognize_google(audio,language='ru-Ru')
            print(f"Вы сказали: {query}")
            search_url = f"https://www.google.com/search?q={query}"
            webbrowser.open(search_url)
        except sr.UnknownValueError:
            print("Ошибка")

def data():
    speak(f"Сегодня {datetime.now().strftime('%d %B %Y')}")

def get_keyboard_layout():
    h_wnd = user32.GetForegroundWindow()
    thread_id = user32.GetWindowThreadProcessId(h_wnd, 0)
    layout_id = user32.GetKeyboardLayout(thread_id)
    language_id = layout_id & (2 ** 16 - 1)
    return language_id


# Функция для переключения на английскую раскладку
def switch_to_english():
    if get_keyboard_layout() != 0x0409:  # 0x0409 - идентификатор английской (США) раскладки
        keyboard.send("alt+shift")
        print("Switched to English")
    else:
        print("Already English layout")

# Функция для переключения на русскую раскладку
def switch_to_russian():
    if get_keyboard_layout() != 0x0419:  # 0x0419 - идентификатор русской раскладки
        keyboard.send("alt+shift")
        print("Переключились на русский")
    else:
        print("Уже русская раскладка")


def open_steam():
    try:
        subprocess.Popen(["C:\\Program Files (x86)\\Steam\\Steam.exe"])
        speak("Steam запущен.")
    except Exception as e:
        speak(f"Произошла ошибка при запуске Steam: {e}")


def open_discord():
    try:
        subprocess.Popen(["C:\\Users\\alkak\\AppData\\Local\\Discord\\Update.exe", "--processStart", "Discord.exe"])
        speak("Discord запущен.")
    except Exception as e:
        speak(f"Произошла ошибка при запуске Discord: {e}")


def get_wikipedia_summary(voice_input):
    print(f"[log] Обработка запроса: {voice_input}")
    reg_ex = re.search('что такое (.*)', voice_input)
    if reg_ex:
        topic = reg_ex.group(1).strip()
        print(f"[log] Поиск информации по теме: {topic}")
        try:
            topic = topic.replace(" ", "_")
            url = f'https://ru.wikipedia.org/wiki/{topic}'
            response = requests.get(url)
            print(f"[log] Запрос к Википедии по URL: {url}")
            if response.status_code == 200:
                html = response.text
                soup = BeautifulSoup(html, 'html.parser')
                paragraphs = soup.select("p")
                summary = next((paragraph.text for paragraph in paragraphs if paragraph.text.strip()), "Описание не найдено.")
                return summary
            else:
                return "Не удалось найти информацию по вашему запросу в Википедии."
        except Exception as e:
            return f"Произошла ошибка при поиске в Википедии: {e}"
    return "Не удалось найти информацию по вашему запросу в Википедии."

def clean_command(cmd, window_cmd=False):
    cmd = cmd.lower()
    for alias in opts['alias']:
        cmd = cmd.replace(alias, '')
    for tbr in opts['tbr']:
        cmd = cmd.replace(tbr, '')
    cleaned_cmd = ' '.join(cmd.split())  # Удалить лишние пробелы
    if window_cmd:
        # Удаление избыточных слов
        for word in ['окно', 'открой', 'закрой', 'сверни']:
            cleaned_cmd = cleaned_cmd.replace(word, '').strip()
    return cleaned_cmd


class VoiceAssistantApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Mihail 0.2")
        self.geometry("300x500")

        ctk.set_default_color_theme("dark-blue")

        # Инициализация объектов распознавания речи
        self.recognizer = sr.Recognizer()
        self.microphone = sr.Microphone()

        # Кнопка для выхода из приложения
        quit_button = ctk.CTkButton(self, text="Quit", command=self.quit_app)
        quit_button.pack(pady=10)

        # Переключатель для управления прослушиванием
        self.listen_switch = ctk.CTkSwitch(self, text="Listening", command=self.toggle_listening)
        self.listen_switch.pack(pady=20)

        self.listening = False
        self.listening_thread = None

    def quit_app(self):
        self.stop_listening()  # Останавливаем прослушивание перед выходом
        self.quit()
        try:
            self.master.destroy()
        except AttributeError:
            self.destroy()

    def toggle_listening(self):
        if self.listen_switch.get() == 1:
            self.start_listening()
        else:
            self.stop_listening()

    def start_listening(self):
        if not self.listening:
            self.listening = True
            self.listening_thread = threading.Thread(target=self.listen)
            self.listening_thread.start()

    def stop_listening(self):
        if self.listening:
            self.listening = False
            if self.listening_thread is not None:
                self.listening_thread.join()  # Ожидание завершения потока
                self.listening_thread = None  # Очистка ссылки на поток

    def listen(self):
        with self.microphone as source:
            self.recognizer.adjust_for_ambient_noise(source, duration=1)  # Калибровка уровня окружающего шума
            while self.listening:
                try:
                    print("Ожидание фразы...")
                    audio = self.recognizer.listen(source, timeout=10)  # Установка тайм-аута на 10 секунд
                    command = self.recognizer.recognize_google(audio, language="ru-RU").lower()
                    print(f"[log] Распознано: {command}")

                    # Используем методы recognize_cmd и execute_cmd
                    cmd_data = self.recognize_cmd(command)
                    if cmd_data:
                        self.execute_cmd(cmd_data['cmd'], command)

                except sr.WaitTimeoutError:
                    print("Таймаут ожидания фразы.")
                except sr.UnknownValueError:
                    print("Извините, я не расслышал команду.")
                except sr.RequestError as e:
                    print(f"Ошибка сервиса распознавания речи; {e}")
                except Exception as e:
                    print(f"Неизвестная ошибка: {e}")
                finally:
                    if not self.listening:
                        break  # Выход из цикла при остановке прослушивания

    def recognize_cmd(self, command):
        # Определение команды из распознанной фразы
        for cmd, phrases in opts['cmds'].items():
            if any(phrase in command for phrase in phrases):
                return {"cmd": cmd}
        return None

    def execute_cmd(self, cmd, voice_input=None):
        commands = {
            'ctime': lambda: speak(f"Сейчас {datetime.now().strftime('%H:%M')}"),
            'data': lambda: speak(f"Сегодня {datetime.now().strftime('%d %B %Y')}"),
            'radio': lambda: webbrowser.open(
                "https://www.youtube.com/watch?v=7x3RW2BCG68&list=PL-SnCU0L3jIkKCh8o3W9hY8qTmHEYTtpt"),
            'dis': lambda: os.system('taskmgr'),
            'google': lambda: webbrowser.open_new('chrome'),
            'spasibo': lambda: speak("Всегда пожалуйста"),
            'kak_dela': lambda: speak("Все хорошо, спасибо что спросили"),
            'steam': open_steam,
            'ds': open_discord,
            'wikipedia': lambda: speak(get_wikipedia_summary(voice_input)),
            'open_window': lambda: open_window(clean_command(voice_input, window_cmd=True)),
            'close_window': lambda: close_window(clean_command(voice_input, window_cmd=True)),
            'minimize_window': lambda: minimize_window(clean_command(voice_input, window_cmd=True)),
            'list_windows': list_windows,
            'rus': switch_to_russian,
            'ang': switch_to_english,
            'change_volume': lambda: set_volume(get_volume_from_voice_input(voice_input)),
            'googl_sk': search_google
        }

        command_func = commands.get(cmd, lambda: print("Извините, я не понимаю эту команду."))
        command_func()

    def clean_command(self, voice_input, window_cmd=False):
        cmd = voice_input.lower()  # Преобразуем вход в нижний регистр
        for alias in opts['alias']:
            cmd = cmd.replace(alias, '')
        for tbr in opts['tbr']:
            cmd = cmd.replace(tbr, '')
        cleaned_cmd = ' '.join(cmd.split())  # Удалить лишние пробелы
        if window_cmd:
            # Удаление избыточных слов
            for word in ['окно', 'открой', 'закрой', 'сверни']:
                cleaned_cmd = cleaned_cmd.replace(word, '').strip()
        return cleaned_cmd


if __name__ == "__main__":
    app = VoiceAssistantApp()
    app.mainloop()

    # class VoiceAssistantApp(ctk.CTk):
    #     def __init__(self):
    #         super().__init__()
    #
    #         self.title("Mihail 0.2")
    #         self.geometry("300x500")
    #
    #         customtkinter.set_default_color_theme("dark-blue")
    #
    #         # Кнопка для выхода из приложения
    #         quit_button = ctk.CTkButton(self, text="Quit", command=self.quit_app)
    #         quit_button.pack(pady=10)
    #
    #         # Кнопка для начала прослушивания
    #         self.listen_button = ctk.CTkButton(self, text="Start Listening", command=self.toggle_listening)
    #         self.listen_button.pack(pady=20)
    #
    #     def quit_app(self):
    #         self.quit()  # Завершение работы окна приложения
    #
    #         # Попробуйте уничтожить окно с помощью self, если self.master не установлен
    #         try:
    #             self.master.destroy()
    #         except AttributeError:
    #             self.destroy()  # Уничтожение окна через self
    #
    #     def toggle_listening(self):
    #         if self.listen_button.cget("text") == "Start Listening":
    #             self.listen_button.configure(text="Stop Listening")
    #             self.start_listening()
    #         else:
    #             self.listen_button.configure(text="Start Listening")
    #             self.stop_listening()
    #
    #     def start_listening(self):
    #         threading.Thread(target=self.listen).start()
    #
    #     def stop_listening(self):
    #         self.listening = False
    #
    #     def listen(self):
    #         self.listening = True
    #         while self.listening:
    #             with microphone as source:
    #                 recognizer.adjust_for_ambient_noise(source)
    #                 audio = recognizer.listen(source)
    #
    #             try:
    #                 command = recognizer.recognize_google(audio, language="ru-RU").lower()
    #                 print(f"[log] Распознано: {command}")
    #                 execute_cmd(recognize_cmd(command)['cmd'], command)
    #             except sr.UnknownValueError:
    #                 print("Извините, я не расслышал команду.")
    #             except sr.RequestError as e:
    #                 print(f"Ошибка сервиса распознавания речи; {e}")
    #
    #
    # if __name__ == "__main__":
    #     app = VoiceAssistantApp()
    #     app.mainloop()
