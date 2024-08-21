import subprocess
import sys

def install_dependencies():
    dependencies = [
        'pillow',
        'pywin32',
        'pyperclip',
        'tqdm'
    ]
    
    print("Установка зависимостей...")
    for package in dependencies:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Успешно установлен пакет: {package}")
        except subprocess.CalledProcessError:
            print(f"Не удалось установить пакет: {package}")
    print("Установка зависимостей завершена.")

# Устанавливаем зависимости перед импортом
install_dependencies()

# Теперь импортируем остальные модули
import json
import os
import shutil
import re
from datetime import datetime
import win32com.client
from PIL import Image
import pyperclip
import glob
import tqdm

# Словарь для хранения тем
THEMES = {
    2: "Ручная чистка",
    4: "Каналопромывка",
    6: "Ежедневные отчеты",
    9: "АВАРИИ",
    12: "Отчет о выполнении работ по объектам",
    18: "Работа наемной техники",
    58: "Объект Днепропетровское шоссе",
    121: "Стройка ливневки колодцы"
}

def update_themes(filename):
    """
    Обновляет словарь THEMES новыми темами из файла JSON.

    Args:
        filename: Имя файла JSON.
    """
    global THEMES
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for message in data['messages']:
        if message['type'] == 'service' and message['action'] == 'topic_created':
            THEMES[int(message['id'])] = message['title']

    print(f"Обновлен словарь тем. Всего тем: {len(THEMES)}")
    print("Текущий список тем:")
    for id, theme in THEMES.items():
        print(f"{id}: {theme}")

def get_date_range(messages):
    """Находит самую раннюю и самую позднюю даты сообщений."""
    dates = [int(msg['date_unixtime']) for msg in messages if 'date_unixtime' in msg]
    if not dates:
        return None, None
    earliest = datetime.fromtimestamp(min(dates))
    latest = datetime.fromtimestamp(max(dates))
    return earliest.strftime("%d-%m-%y"), latest.strftime("%d-%m-%y")

def process_json(json_file):
    """Обрабатывает JSON файл и сортирует сообщения по темам."""
    global sorted_folder_name  # Объявляем переменную как глобальную
    
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    messages = data['messages']
    print("Начали обработку сообщений")

    photos_path = os.path.join(os.path.dirname(__file__), 'photos')
    
    # Получаем диапазон дат и создаем новое имя для папки sorted
    start_date, end_date = get_date_range(messages)
    sorted_folder_name = f"sorted_{start_date}_to_{end_date}"
    sorted_path = os.path.join(os.path.dirname(__file__), sorted_folder_name)
    print(f"Путь к папке с отсортированными файлами: {sorted_path}")

    theme_folders = {}
    processed_photos = set()  # Множество для отслеживания обработанных фотографий

    for message in messages:
        # Пропускаем сообщения, если 'from_id' или 'reply_to_message_id' отсутствуют
        if 'from_id' not in message or 'reply_to_message_id' not in message:
            continue

        # Пропускаем сообщения с group_id = 6
        if 'group_id' in message and message['group_id'] == 6:
            continue

        # Пропускаем сообщения, если поле text больше 255 символов
        if 'text' in message and len(message['text']) > 235:
            continue

        print(f"Обрабатываем сообщение с id: {message['id']}")
        reply_to_id = int(message['reply_to_message_id'])
        if reply_to_id in THEMES:
            theme = THEMES[reply_to_id]
            theme_path = os.path.join(sorted_path, theme.replace(" ", "_"))

            if theme_path not in theme_folders:
                theme_folders[theme_path] = {}

            # Группируем сообщения
            grouped_messages = [message]
            for msg in messages:
                # Преобразуем значения date_unixtime в числа
                msg_date_unixtime = int(msg['date_unixtime'])
                message_date_unixtime = int(message['date_unixtime'])

                # Условие группировки по разнице во времени не более 3 секунд
                if (
                        abs(msg_date_unixtime - message_date_unixtime) <= 3
                        and msg['from_id'] == message['from_id']
                        and msg['reply_to_message_id'] == message['reply_to_message_id']
                ):
                    grouped_messages.append(msg)
                    print(f"Добавлено сообщение в группу: {msg['id']}")

            # Проверяем, есть ли в группе сообщения с текстом
            subfolder_name = "Без_текста"
            for msg in grouped_messages:
                if 'text' in msg and msg['text']:
                    # Заменяем /n на _ в тексте
                    subfolder_name = msg['text'].replace(" ", "_").replace("\n", "_")
                    # Очищаем имя папки от недопустимых символов
                    subfolder_name = re.sub(r'[^\w\s-]', '_', subfolder_name).strip()
                    break

            subfolder_path = os.path.join(theme_path, subfolder_name)
            
            # Проверяем, есть ли фотографии для перемещения
            photos_to_move = [msg for msg in grouped_messages if 'photo' in msg and msg['photo'] not in processed_photos]
            if photos_to_move:
                if subfolder_path not in theme_folders[theme_path]:
                    theme_folders[theme_path][subfolder_path] = []
                theme_folders[theme_path][subfolder_path].extend(photos_to_move)

    # Создаем папки и перемещаем фотографии
    for theme_path, subfolders in theme_folders.items():
        if subfolders:
            os.makedirs(theme_path, exist_ok=True)
            print(f"Создана папка темы: {theme_path}")
            
            for subfolder_path, photos in subfolders.items():
                os.makedirs(subfolder_path, exist_ok=True)
                print(f"Создана подпапка: {subfolder_path}")
                
                for msg in photos:
                    if msg['photo'] not in processed_photos:  # Проверяем, не была ли фотография уже обработана
                        source_file = os.path.join(photos_path, os.path.basename(msg['photo']))
                        target_file = os.path.join(subfolder_path, os.path.basename(msg['photo']))

                        if os.path.exists(source_file):
                            shutil.move(source_file, target_file)
                            processed_photos.add(msg['photo'])
                            print(f"Перемещен файл: {source_file} в {target_file}")
                        else:
                            print(f"Файл не найден: {source_file}")
    clear_terminal()
    print("Обработка сообщений завершена")
    print(f"Всего обработано уникальных фотографий: {len(processed_photos)}")

    # Проверяем, существует ли папка photos, и удаляем ее, если она пуста
    if os.path.exists(photos_path):
        if not os.listdir(photos_path):
            os.rmdir(photos_path)
            print("Папка photos удалена, так как она пуста.")
        else:
            print("ВНИМАНИЕ: В папке photos остались нерассортированные фотографии.")
    else:
        print("Папка photos не существует.")

def clear_terminal():
    os.system('cls' if os.name == 'nt' else 'clear')

def parse_folder_selection(input_string, max_index):
    selected_indices = set()
    parts = input_string.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            start, end = map(int, part.split('-'))
            selected_indices.update(range(start - 1, min(end, max_index)))
        else:
            try:
                index = int(part) - 1
                if 0 <= index < max_index:
                    selected_indices.add(index)
            except ValueError:
                pass
    return sorted(list(selected_indices))

def sanitize_folder_name(name):
    """Очищает имя папки от недопустимых символов."""
    # Запрещенные символы в Windows
    forbidden_chars = r'<>:"/\|?*'
    # Заменяем запрещенные символы на подчеркивание
    for char in forbidden_chars:
        name = name.replace(char, '_')
    # Удаляем точки и пробелы в начале и конце имени
    name = name.strip('. ')
    # Ограничиваем длину имени папки до 255 символов
    return name[:255]

def process_photos_and_create_docx(folder_path, text_value):
    """Обрабатывает фотографии и создает docx файл без переименования файлов."""
    word = win32com.client.Dispatch("Word.Application")
    status = ""

    if not text_value:
        # Если текст не указан, используем имя папки
        text_value = os.path.basename(folder_path)
    elif len(text_value) > 30:
        return "Ошибка: текст превышает 30 символов"

    # Копируем путь к папке и имя файла в буфер обмена, разделяя их символом |
    pyperclip.copy(f"{folder_path}|{text_value}")

    image_files = glob.glob(os.path.join(folder_path, '*.jpg'))
    image_files.sort(key=lambda x: os.path.getctime(x))
    total_images = len(image_files)

    print("Обработка изображений:")
    for file_path in tqdm.tqdm(image_files, total=total_images, unit="фото"):
        try:
            with Image.open(file_path) as img:
                width, height = img.size
                if width > height:
                    img = img.transpose(method=Image.ROTATE_90)
                    img.save(file_path)
        except Exception as e:
            status += f"Не удалось обработать изображение: {os.path.basename(file_path)}. Ошибка: {str(e)}\n"

    status += f"Обработано {total_images} изображений.\n"

    script_dir = os.path.dirname(os.path.abspath(__file__))
    docm_file_path = os.path.join(script_dir, 'word_jpg_auto_v5.docm')

    if not os.path.exists(docm_file_path):
        return "Ошибка: файл word_jpg_auto_v5.docm не найден"

    print("Создание docx файла...")
    word.Documents.Open(docm_file_path)
    word.Application.Run("Макрос1")
    word.ActiveDocument.Close()
    word.Quit()

    # Создаем папку docx, если она не существует
    docx_folder = os.path.join(script_dir, 'docx')
    os.makedirs(docx_folder, exist_ok=True)

    for file_name in os.listdir(folder_path):
        if file_name.endswith(".docx"):
            source_path = os.path.join(folder_path, file_name)
            destination_path = os.path.join(docx_folder, file_name)
            shutil.move(source_path, destination_path)
            pyperclip.copy(destination_path)
            status += f"Сформирован {file_name} и перемещен в папку docx\n"

    return status

def merge_folders(sorted_folder_name):
    clear_terminal()
    photos_path = os.path.join(os.path.dirname(__file__), 'photos')
    if os.path.exists(photos_path) and os.listdir(photos_path):
        print("ВНИМАНИЕ: ЕСТЬ ФОТО НЕ РАССОРТИРОВАННЫЕ ПО РАЗДЕЛАМ (в папке photos)")
        print("=" * 50)

    sorted_path = os.path.join(os.path.dirname(__file__), sorted_folder_name)
    themes = [f for f in os.listdir(sorted_path) if os.path.isdir(os.path.join(sorted_path, f))]
    
    while True:
        print("Список доступных тем:")
        for i, theme in enumerate(themes, 1):
            print(f"{i}. {theme}")
        print("0. Выход")
        
        theme_choice = input("Выберите номер темы или '0' для выхода: ").strip()
        if theme_choice == '0':
            return 'exit'  # Возвращаем 'exit' для завершения программы
        
        try:
            theme_index = int(theme_choice) - 1
            selected_theme = themes[theme_index]
        except (ValueError, IndexError):
            print("Неверный выбор. Попробуйте снова.")
            input("Нажмите Enter для продолжения...")
            continue
        
        while True:
            clear_terminal()
            theme_path = os.path.join(sorted_path, selected_theme)
            subfolders = sorted([f for f in os.listdir(theme_path) if os.path.isdir(os.path.join(theme_path, f))])
            
            print(f"Подпапки темы '{selected_theme}':")
            for i, folder in enumerate(subfolders, 1):
                print(f"{i}. {folder}")
            print("0. Вернуться к выбору темы")
            
            choices = input("Введите номера подпапок для слияния/переименования (через запятую или диапазон, например 1,3-5,7), 'д' для создания docx или '0' для возврата: ").strip()
            if choices == '0':
                clear_terminal()
                break
            elif choices.lower() == 'д':
                clear_terminal()
                while True:
                    if not subfolders:
                        print("В этой теме нет подпапок.")
                        input("Нажмите Enter для продолжения...")
                        break
                    
                    print("Выберите папку для создания docx:")
                    for i, folder in enumerate(subfolders, 1):
                        print(f"{i}. {folder}")
                    print("0. Вернуться в предыдущее меню")
                    folder_choice = input("Введите номер папки для создания docx или '0' для возврата: ").strip().lower()
                    
                    if folder_choice == '0':
                        clear_terminal()
                        break
                    
                    try:
                        folder_index = int(folder_choice) - 1
                        if 0 <= folder_index < len(subfolders):
                            selected_folder = subfolders[folder_index]
                            folder_path = os.path.join(theme_path, selected_folder)
                            text_value = input("Введите имя файла для docx (или нажмите Enter для использования имени папки): ").strip()
                            
                            status = process_photos_and_create_docx(folder_path, text_value)
                            print(status)
                            input("Нажмите Enter для продолжения...")
                            clear_terminal()
                        else:
                            print("Неверный выбор папки.")
                            input("Нажмите Enter для продолжения...")
                    except ValueError:
                        print("Неверный ввод. Введите число.")
                        input("Нажмите Enter для продолжения...")
                continue

            folder_indices = parse_folder_selection(choices, len(subfolders))
            if not folder_indices:
                print("Неверный выбор. Попробуйте снова.")
                input("Нажмите Enter для продолжения...")
                continue
            
            selected_folders = [subfolders[i] for i in folder_indices]
            
            if len(selected_folders) == 1:
                # Переименование одной папки
                old_name = selected_folders[0]
                while True:
                    new_name = input(f"Введите новое имя для папки '{old_name}' (или '0' для отмены): ").strip()
                    if new_name == '0':
                        break
                    
                    sanitized_name = sanitize_folder_name(new_name)
                    if sanitized_name != new_name:
                        print(f"Имя папки было изменено на '{sanitized_name}' из-за недопустимых символов.")
                        choice = input("Продолжить с этим именем? (да/нет): ").lower()
                        if choice != 'да':
                            continue

                    old_path = os.path.join(theme_path, old_name)
                    new_path = os.path.join(theme_path, sanitized_name)
                    
                    if os.path.exists(new_path):
                        print(f"Папка с именем '{sanitized_name}' уже существует. Выберите другое имя.")
                        continue
                    
                    os.rename(old_path, new_path)
                    break
            
            elif len(selected_folders) >= 2:
                # Слияние нескольких папок
                while True:
                    target_folder = input("Введите имя новой подпапки для слияния (или '0' для отмены): ").strip()
                    if target_folder == '0':
                        break
                    
                    sanitized_name = sanitize_folder_name(target_folder)
                    if sanitized_name != target_folder:
                        print(f"Имя папки было изменено на '{sanitized_name}' из-за недопустимых символов.")
                        choice = input("Продолжить с этим именем? (да/нет): ").lower()
                        if choice != 'да':
                            continue

                    target_path = os.path.join(theme_path, sanitized_name)
                    
                    if not os.path.exists(target_path):
                        os.makedirs(target_path, exist_ok=True)
                    
                    for folder in selected_folders:
                        if folder != sanitized_name:
                            source_path = os.path.join(theme_path, folder)
                            for item in os.listdir(source_path):
                                s = os.path.join(source_path, item)
                                d = os.path.join(target_path, item)
                                if os.path.isdir(s):
                                    shutil.copytree(s, d, dirs_exist_ok=True)
                                else:
                                    if os.path.exists(d):
                                        base, extension = os.path.splitext(d)
                                        counter = 1
                                        while os.path.exists(d):
                                            d = f"{base}_{counter}{extension}"
                                            counter += 1
                                    shutil.copy2(s, d)
                            shutil.rmtree(source_path)
                    break

            # Обновляем список подпапок после операции
            subfolders = sorted([f for f in os.listdir(theme_path) if os.path.isdir(os.path.join(theme_path, f))])

    return None  # Возвращаем None, если пользователь закончил работу с темами

# Основной код
json_file = os.path.join(os.path.dirname(__file__), 'result.json')

# Шаг 1: Обновляем словарь тем
update_themes(json_file)

# Шаг 2: Обрабатываем JSON файл и создаем папки
process_json(json_file)

# Шаг 3: Предлагаем пользователю объединить подпапки в темах
merge_folders(sorted_folder_name)

print("Программа завершена.")