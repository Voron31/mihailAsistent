# библиотеки
import os
import time
import subprocess
import platform
import requests
from bs4 import BeautifulSoup
import pyttsx3
import datetime
import webbrowser
from fuzzywuzzy import fuzz
import re
import speech_recognition as sr
from colorama import init, Fore, Style
import calendar
from datetime import datetime
import pygetwindow as gw
import pyautogui
import ctypes
import sys
import locale
import re
import requests
import json
import wikipediaapi


# Инициализация colorama
init(autoreset=True)

# Настройки
opts = {
    "alias": ('миша', 'миш', 'мишь', 'мих', 'михаил', 'мишка', 'медведь', 'мишенька',"миша", "мишу"),
    "tbr": ('скажи', 'расскажи', 'покажи', 'сколько', 'произнеси',
            'ищи в гугл', 'найди в гугле', 'поиск google', 'найди в гугл', 'поиск в гугл',
            'ищи в google', 'найди в google',
            'найди википедия', 'поиск в википедии', ' что такое', 'найди гугл', 'в гугл',
            "скажи", "покажи", "скрой", "заверши", "удали"),
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
        "wikipedia": ('найди википедия', 'поиск в википедии', 'что такое','Что такое', 'найди в википедии'),
        "search_google": ('найди в гугл', 'найди в гугле', 'поиск в гугл', 'ищи в гугл', 'найди в google', 'найди', 'найди что такое'),
        "open_window": ('открой окно', 'активируй окно', 'покажи окно'),
        "close_window": ('закрой окно', 'удали окно'),
        "minimize_window": ('сверни окно', 'убери окно'),
        "list_windows": ('покажи окна', 'какие окна открыты', 'список окон')
    }
}

# Инициализация pyttsx3
speak_engine = pyttsx3.init()
speak_engine.setProperty('voice', speak_engine.getProperty('voices')[0].id)

# Инициализация распознавателя речи
recognizer = sr.Recognizer()
microphone = sr.Microphone()

# Путь к файлу для сохранения данных
DATA_FILE = "currency_data.json"

# Флаг для ожидания ответа на запрос подтверждения
awaiting_confirmation = False
recording = False
recording_file = os.path.join(os.getcwd(), 'записи.txt')
# письмо зелёным
def print_green(text):
    print(Fore.GREEN + text + Style.RESET_ALL)  # Печать текста зеленым цветом

def speak(what):
    print_green(what)
    speak_engine.say(what)
    speak_engine.runAndWait()

def recognize_cmd(cmd):
    RC = {'cmd': '', 'percent': 0}
    
    # Проверяем на наличие запроса к Википедии
    if "что такое" in cmd or "поиск в википедии" in cmd:
        return {'cmd': 'wikipedia', 'percent': 100}
    
    # Остальная логика
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

# функция лист октивных окон
def list_windows():
    try:
        windows = gw.getAllTitles()
        active_windows = [w for w in windows if w.strip()]
        print(f"[log] Активные окна: {active_windows}")
        speak("Активные окна: " + ", ".join(active_windows))
        return active_windows
    except Exception as e:
        speak(f"Произошла ошибка при получении списка окон: {e}")
        return []

# функция открытия окон
def open_window(title):
    title = clean_command(title, window_cmd=True)
    print(f"[log] Открытие окна: '{title}'")
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

# функция закрытия окон
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

# свёртование окон
def minimize_window(title):
    title = clean_command(title, window_cmd=True)
    print(f"[log] Сворачивание окна: '{title}'")
    try:
        windows = gw.getWindowsWithTitle(title)
        if windows:
            window = windows[0]
            if not window.isMinimized:
                window.minimize()
                if window.isMinimized:
                    speak(f"Окно '{title}' свернуто.")
                else:
                    speak(f"Не удалось свернуть окно '{title}'.")
            else:
                speak(f"Окно '{title}' уже свернуто.")
        else:
            speak(f"Окно с названием '{title}' не найдено.")
    except Exception as e:
        speak(f"Произошла ошибка при сворачивании окна: {e}")
# выполнение команд
def execute_cmd(cmd, voice_input=None):
    global awaiting_confirmation, recording
    print_green(f"[log] Выполнение команды: {cmd}")

    # Команды
    commands = {
        'ctime': lambda: speak(f"Сейчас {datetime.now().strftime('%H:%M')}"),#раб
        'data': data,#раб
        'radio': lambda: webbrowser.open("https://www.youtube.com"),#раб
        'stupid1': lambda: speak("Колобок повесился"),#раб
        'dis': lambda: os.system('taskmgr'),#раб
        'google': lambda: webbrowser.open_new('chrome'),#раб
        'spasibo': lambda: speak("Всегда к вашим услугам"),#раб
        'kak_dela': lambda: speak("Все хорошо, спасибо что спросили"),#раб
        'steam': open_steam,#раб
        'ds': open_discord,#раб
        'wikipedia': lambda: speak(get_wikipedia_summary(clean_wikipedia_command(voice_input))),#раб
        'open_window': lambda: open_window(clean_command(voice_input, window_cmd=True)),#раб
        'close_window': lambda: close_window(clean_command(voice_input, window_cmd=True)),#раб
        'minimize_window': lambda: minimize_window(clean_command(voice_input, window_cmd=True)),#раб
        'list_windows': list_windows#раб
    }

    # Обработка команды "что такое"
    if "что такое" in voice_input.lower():
        query = voice_input.lower().replace("википедия", "").strip()
        summary = get_wikipedia_summary(query)  # Передаем очищенный запрос
        speak(summary)  # Отправляем результат обратно

    if cmd in commands:
        commands[cmd]()
    else:
        speak("Команда не распознана, повторите!")



def clean_command(cmd, window_cmd=False):
    cmd = cmd.lower()
    
    # Удаление обращения к ассистенту
    for alias in opts['alias']:
        cmd = cmd.replace(alias, '')
    
    # Удаление мусорных слов
    for tbr in opts['tbr']:
        cmd = cmd.replace(tbr, '')
    
    # Удаление лишних пробелов
    cleaned_cmd = ' '.join(cmd.split())
    
    # Если команда связана с окнами, удаляем "служебные" слова, но оставляем название окна
    if window_cmd:
        # Используем регулярное выражение для удаления слов "окно", "открой", и т.д. только в начале строки
        cleaned_cmd = re.sub(r'\b(открой|закрой|окно|сверни)\b', '', cleaned_cmd).strip()
    
    # Возвращаем очищенную команду
    return cleaned_cmd






# википедия
def get_wikipedia_summary(voice_input):
    print(f"[log] Обработка запроса: {voice_input}")
    reg_ex = re.search('что такое (.*)', voice_input)
    if reg_ex:
        topic = reg_ex.group(1).strip()
        
        print(f"[log] Поиск информации по теме: {topic}")
        try:
            topic = topic.replace(" ", "_")
            url = f'https://ru.wikipedia.org/wiki/{topic}'  # Замените 'звук' на 'topic'
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


def clean_wikipedia_command(cmd):
    # Очистка команды, убираем лишние слова
    cmd = cmd.lower().replace("миша", "").replace("поиск в википедии", "").replace("что такое", "").strip()
    return cmd


# открытие стим
def open_steam():
    if platform.system() == "Windows":
        subprocess.Popen(['C:\\Program Files (x86)\\Steam\\Steam.exe'])
    else:
        subprocess.Popen(['steam'])
# дискорд
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def open_discord():
    if platform.system() == "Windows":
        discord_path = 'C:\\Users\\alkak\\AppData\\Local\\Discord\\Update.exe --processStart Discord.exe'
        if is_admin():
            try:
                subprocess.Popen([discord_path])
            except Exception as e:
                print(f"Произошла ошибка при запуске Discord: {e}")
        else:
            speak("Дискорд не обноружен")
    else:
        subprocess.Popen(['discord'])

def get_currency_rates():
    url = "https://api.monobank.ua/bank/currency"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            data = response.json()
            save_currency_data(data)
            return data
        else:
            print("Не удалось получить данные от Monobank.")
            return load_saved_data()
    except requests.exceptions.RequestException:
        print("Ошибка сети. Используются сохраненные данные.")
        return load_saved_data()
    
def save_currency_data(data):
    # Сохраняем данные в файл с текущей датой обновления
    with open(DATA_FILE, 'w') as f:
        json.dump({
            "date": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            "rates": data
        }, f)


def load_saved_data():
    # Загружаем данные из файла, если он существует
    try:
        with open(DATA_FILE, 'r') as f:
            saved_data = json.load(f)
            print(f"Используются данные от {saved_data['date']}")
            return saved_data['rates']
    except (FileNotFoundError, json.JSONDecodeError):
        print("Нет сохраненных данных.")
        return None

def print_currency_rates(data):
    currency_ids = [840, 978]  # USD и EUR
    currency_names = {840: "USD", 978: "EUR"}
    
    for entry in data:
        if entry['currencyCodeA'] in currency_ids and entry['currencyCodeB'] == 980:  # UAH
            currency_code = entry['currencyCodeA']
            rate_buy = entry.get('rateBuy')
            rate_sell = entry.get('rateSell')
            
            if rate_buy is not None and rate_sell is not None:
                print(f"1 {currency_names[currency_code]} = {rate_buy} UAH (покупка), {rate_sell} UAH (продажа)")


def data():
    try:
        # Попробуйте установить другую локализацию
        locale.setlocale(locale.LC_TIME, 'ru_RU')
    except locale.Error:
        print("Локализация не поддерживается на этой системе")
    today_date = datetime.now().strftime('%d %B %Y года')
    today_weekday = datetime.now().strftime('%A')
    speak(f"Сегодня {today_date}, {today_weekday}.")
    #speak(f"Сегодня {datetime.now().strftime('%d %B %Y года')}, {calendar.day_name[datetime.now().weekday()]}.")

def greet_in_soviet_style_1():
    try:
        locale.setlocale(locale.LC_TIME, 'ru_RU')  # Устанавливаем русскую локализацию
    except locale.Error:
        print("Локализация не поддерживается на этой системе")
    
    today_date = datetime.now().strftime('%d %B %Y года')
    today_weekday = datetime.now().strftime('%A')
    
    speak("Здравствуйте, товарищ!")
    speak(f"Сегодня на дворе {today_date}, {today_weekday}.")
    print_green("Наступило время великих свершений и трудовых подвигов во славу СССР!")
    data = get_currency_rates()
    if data:
        print_green("Текущие курсы валют, необходимые для планирования народного хозяйства:")
        print_currency_rates(data)
        
    else:
        print("Не удалось загрузить курсы валют.")


def main():
    try:
        greet_in_soviet_style_1()
        while True:
            with microphone as source:
                print_green("[log] Ожидание команды...")
                recognizer.adjust_for_ambient_noise(source)
                audio = recognizer.listen(source)
            try:
                voice_input = recognizer.recognize_google(audio, language="ru-RU")
                print_green(f"[log] Распознано: {voice_input}")
                
                if any(alias in voice_input.lower() for alias in opts["alias"]):
                    cmd = recognize_cmd(voice_input)  # Используйте `voice_input` без очистки
                    if cmd['cmd']:
                        execute_cmd(cmd['cmd'], voice_input)
                    else:
                        speak("Команда не распознана, повторите!")
                elif "википедия" in voice_input.lower():
                    query = voice_input.lower().replace("википедия", "").strip()  # Извлечение запроса
                    cmd = {"cmd": "wikipedia"}
                    execute_cmd(cmd['cmd'], query)  # Передаем только запрос
                elif recording:
                    write_recording(voice_input)
                time.sleep(0.1)
            except sr.UnknownValueError:
                print_green("[log] Не удалось распознать речь")
            except sr.RequestError as e:
                print_green(f"[log] Ошибка сервиса распознавания речи; {e}")
            except Exception as e:
                print_green(f"[log] Произошла ошибка во время обработки: {e}")
    except KeyboardInterrupt:
        print_green("[log] Остановлено пользователем")
    except Exception as e:
        print_green(f"[log] Произошла ошибка: {e}")

if __name__ == '__main__':
    main()





