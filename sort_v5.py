import sys
import subprocess
import os
import re
import json
from datetime import datetime
import glob
import shutil
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter
import win32com.client

def install_dependencies():
    dependencies = [
        'pillow',
        'pywin32',
        'pyperclip',
        'tqdm',
        'requests',
        'openpyxl'
    ]
    
    print("Установка зависимостей...")
    for package in dependencies:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
            print(f"Успешно установлен пакет: {package}")
        except subprocess.CalledProcessError:
            print(f"Не удалось установить пакет: {package}")
    print("Установка зависимостей завершена.")

try:
    # Попытка импорта необходимых модулей
    import requests
    from PIL import Image
    import win32com.client
    import pyperclip
    import tqdm
except ImportError as e:
    print(f"Ошибка импорта: {e}")
    print("Попытка установки недостающих зависимостей...")
    install_dependencies()
    # Повторная попытка импорта после установки
    import requests
    from PIL import Image
    import win32com.client
    import pyperclip
    import tqdm

#  Добавляем версии
SCRIPT_VERSION = "5.1.3"
DOCM_VERSION = "1.0.6"
EXCEL_TEMPLATE_VERSION = "1.0.0"  # Добавляем версию для шаблона Excel

GITHUB_REPO = "cyberkek587/drain_tg"
SCRIPT_NAME = "sort_v5.py"
DOCM_NAME = "word_jpg_auto_v5.docm"
EXCEL_TEMPLATE_NAME = "excel_summary_template.xlsm"

def check_for_updates():
    """Проверяет наличие обновлений скрипта, .docm файла и шаблона Excel на GitHub."""
    try:
        requests.get("https://github.com", timeout=5)
    except requests.ConnectionError:
        print("Нет подключения к интернету. Пропускаем проверку обновлений.")
        return False

    try:
        script_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{SCRIPT_NAME}"
        docm_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{DOCM_NAME}"
        excel_template_url = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{EXCEL_TEMPLATE_NAME}"
        
        script_response = requests.get(script_url, timeout=10)
        docm_response = requests.get(docm_url, timeout=10)
        excel_template_response = requests.get(excel_template_url, timeout=10)
        
        update_available = False
        
        if script_response.status_code == 200:
            remote_script = script_response.text
            remote_script_version = re.search(r'SCRIPT_VERSION\s*=\s*"([\d.]+)"', remote_script)
            if remote_script_version and remote_script_version.group(1) > SCRIPT_VERSION:
                print(f"Доступно обновление скрипта. Текущая версия: {SCRIPT_VERSION}, новая версия: {remote_script_version.group(1)}")
                update_available = True
        
        if docm_response.status_code == 200:
            remote_docm = docm_response.content
            remote_docm_version = re.search(r'DOCM_VERSION\s*=\s*"([\d.]+)"', remote_script)
            if remote_docm_version and remote_docm_version.group(1) > DOCM_VERSION:
                print(f"Доступно обновление .docm файла. Текущая версия: {DOCM_VERSION}, новая версия: {remote_docm_version.group(1)}")
                update_available = True
        
        if excel_template_response.status_code == 200:
            remote_excel_template = excel_template_response.content
            remote_excel_template_version = re.search(r'EXCEL_TEMPLATE_VERSION\s*=\s*"([\d.]+)"', remote_script)
            if remote_excel_template_version and remote_excel_template_version.group(1) > EXCEL_TEMPLATE_VERSION:
                print(f"Доступно обновление шаблона Excel. Текущая версия: {EXCEL_TEMPLATE_VERSION}, новая версия: {remote_excel_template_version.group(1)}")
                update_available = True
        
        if update_available:
            choice = input("Хотите обновить файлы? (д/н): ").lower()
            if choice == 'д':
                print("Обновляем файлы...")
                try:
                    # Создаем резервную копию текущего скрипта
                    backup_name = f"{SCRIPT_NAME}.bak"
                    shutil.copy(__file__, backup_name)
                    print(f"Создана резервная копия скрипта: {backup_name}")

                    with open(__file__, 'w', encoding='utf-8') as file:
                        file.write(remote_script)
                    with open(DOCM_NAME, 'wb') as file:
                        file.write(remote_docm)
                    
                    # Скачиваем шаблон Excel
                    excel_template_path = os.path.join(os.path.dirname(__file__), EXCEL_TEMPLATE_NAME)
                    with open(excel_template_path, 'wb') as file:
                        file.write(excel_template_response.content)
                    print("Шаблон Excel обновлен.")
                    
                    print("Файлы обновлены. Перезапустите программу.")
                    input("Нажмите Enter для завершения...")
                    return True
                except Exception as e:
                    print(f"Ошибка при обновлении файлов: {e}")
                    print("Восстанавливаем предыдущую версию скрипта...")
                    shutil.copy(backup_name, __file__)
                    print("Предыдущая версия скрипта восстановлена.")
                    print("Продолжаем работу с текущими версиями.")
                    input("Нажмите Enter для продолжения...")
        else:
            print(f"У вас установлены последние версии скрипта ({SCRIPT_VERSION}), .docm файла ({DOCM_VERSION}) и шаблона Excel ({EXCEL_TEMPLATE_VERSION}).")
        
        return False
    except Exception as e:
        print(f"Ошибка при проверке обновлений: {e}")
        print("Продолжаем работу с текущими версиями файлов.")
        input("Нажмите Enter для продолжения...")
        return False

# Вызываем функцию проверки обновлений
if check_for_updates():
    sys.exit()

# Проверяем наличие файлов и скачиваем их, если отсутствуют
def check_and_download_file(file_name, url):
    if not os.path.exists(file_name):
        try:
            response = requests.get(url, timeout=10)
            if response.status_code == 200:
                with open(file_name, 'wb') as file:
                    file.write(response.content)
                print(f"Файл {file_name} успешно скачан.")
            else:
                print(f"Не удалось скачать файл {file_name}. Код ответа: {response.status_code}")
        except Exception as e:
            print(f"Ошибка при скачивании файла {file_name}: {e}")

check_and_download_file(DOCM_NAME, f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{DOCM_NAME}")
check_and_download_file(EXCEL_TEMPLATE_NAME, f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/{EXCEL_TEMPLATE_NAME}")

# Словарь для хранения тем
THEMES = {
    2: "Ручная чистка",
    4: "Перехваты Авраменко Святогеоргиевска",
    6: "Ежедневные отчеты",
    9: "АВАРИИ",
    12: "Отчет о выполнении работ по объектам",
    18: "Работа наемной техники",
    58: "Объект Днепропетровское шоссе",
    121: "Стройка ливневки колодцы"
}

def update_themes(filename):
    """
    Обновляет словарь THEMES новыми темами из файла JSON,
    сохраняя приоритет существующих названий для указанных ID.

    Args:
        filename: Имя файла JSON.
    """
    global THEMES
    with open(filename, 'r', encoding='utf-8') as f:
        data = json.load(f)

    existing_themes = THEMES.copy()
    new_themes = {}
    updated_themes = []

    for message in data['messages']:
        if message['type'] == 'service' and message['action'] == 'topic_created':
            theme_id = int(message['id'])
            if theme_id not in existing_themes:
                new_themes[theme_id] = message['title']
                updated_themes.append((theme_id, message['title']))

    # Обновляем THEMES, сохраняя приоритет существующих названий
    THEMES.update(new_themes)

    print(f"Обновлен словарь тем. Всего тем: {len(THEMES)}")
    if updated_themes:
        print("Новые темы:")
        for id, theme in updated_themes:
            print(f"{id}: {theme}")
    else:
        print("Новых тем не обнаружено.")
    
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
            photos_to_move = [msg for msg in grouped_messages if 'photo' in msg]
            if photos_to_move:
                if subfolder_path not in theme_folders[theme_path]:
                    theme_folders[theme_path][subfolder_path] = []
                theme_folders[theme_path][subfolder_path].extend(photos_to_move)

    # Создаем папки и копируем фотографии
    for theme_path, subfolders in theme_folders.items():
        if subfolders:
            os.makedirs(theme_path, exist_ok=True)
            print(f"Создана папка темы: {theme_path}")
            
            for subfolder_path, photos in subfolders.items():
                os.makedirs(subfolder_path, exist_ok=True)
                print(f"Создана подпапка: {subfolder_path}")
                
                for msg in photos:
                    source_file = os.path.join(photos_path, os.path.basename(msg['photo']))
                    target_file = os.path.join(subfolder_path, os.path.basename(msg['photo']))

                    if os.path.exists(source_file):
                        shutil.copy2(source_file, target_file)
                        print(f"Скопирован файл: {source_file} в {target_file}")
                        processed_photos.add(source_file)
                    else:
                        print(f"Файл не найден: {source_file}")

    # Удаляем обработанные фотографии из папки photos
    for photo in processed_photos:
        try:
            os.remove(photo)
            print(f"Удален файл: {photo}")
        except OSError as e:
            print(f"Ошибка при удалении файла {photo}: {e}")

    # Проверяем, существует ли папка photos, и удаляем ее, если она пуста
    if os.path.exists(photos_path):
        if not os.listdir(photos_path):
            os.rmdir(photos_path)
            print("Папка photos удалена, так как она пуста.")
        else:
            print("ВНИМАНИЕ: В папке photos остались нерассортированные фотографии.")
    else:
        print("Папка photos не существует.")

    clear_terminal()
    print("Обработка сообщений завершена")
    print(f"Всего обработано уникальных фотографий: {len(processed_photos)}")

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

def update_excel(docx_files, prefix):
    template_path = os.path.join(os.path.dirname(__file__), EXCEL_TEMPLATE_NAME)
    if not os.path.exists(template_path):
        print("Шаблон Excel не найден. Пожалуйста, запустите проверку обновлений.")
        return

    excel_path = os.path.join(os.path.dirname(__file__), 'docx', 'Содержание.xlsm')
    
    if not os.path.exists(excel_path):
        shutil.copy(template_path, excel_path)
        print(f"Создан новый файл {excel_path} на основе шаблона.")
    wb = openpyxl.load_workbook(excel_path, keep_vba=True)

    sheet_name = prefix if prefix else "Без префикса"
    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        ws.append(["№", "Название объекта", "Количество", "Примечание", "Печать"])
        for col in range(1, 6):
            ws.cell(row=1, column=col).font = Font(bold=True)
    else:
        ws = wb[sheet_name]

    # Определяем стили для границ
    thin_border = Border(left=Side(style='thin'), 
                         right=Side(style='thin'), 
                         top=Side(style='thin'), 
                         bottom=Side(style='thin'))

    # Удаляем существующую итоговую строку, если она есть
    if ws.cell(row=ws.max_row, column=2).value == "Всего":
        ws.delete_rows(ws.max_row)

    # Добавляем новые данные
    for docx_file in docx_files:
        file_name = os.path.splitext(os.path.basename(docx_file))[0]
        relative_path = os.path.relpath(docx_file, start=os.path.dirname(excel_path))
        ws.append(["", f'=HYPERLINK("{relative_path}", "{file_name}")', "", "", ""])

    # Добавляем итоговую строку
    ws.append(["", "Всего", "", "", ""])
    ws.cell(row=ws.max_row, column=2).font = Font(bold=True)

    # Нумеруем непустые строки и определяем диапазон суммирования
    row_number = 1
    first_data_row = None
    last_data_row = None
    for row in range(2, ws.max_row):
        if ws.cell(row=row, column=2).value and ws.cell(row=row, column=2).value != "Всего":
            ws.cell(row=row, column=1).value = row_number
            row_number += 1
            if first_data_row is None:
                first_data_row = row
            last_data_row = row

    # Обновляем формулу в итоговой строке
    if first_data_row and last_data_row:
        ws.cell(row=ws.max_row, column=3).value = f"=SUM(C{first_data_row}:C{last_data_row})"

    # Применяем границы и выравнивание ко всем ячейкам
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=5):
        for cell in row:
            cell.border = thin_border
            if cell.column == 1:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif cell.column == 2:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Применяем автоподбор ширины столбцов
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column
        for cell in column_cells:
            if cell.row == 1 or cell.row == ws.max_row:  # Заголовок или итоговая строка
                max_length = max(max_length, len(str(cell.value)))
            elif column == 2:  # Второй столбец (с гиперссылками)
                if cell.value and cell.value.startswith('=HYPERLINK'):
                    # Извлекаем текст гиперссылки
                    display_text = cell.value.split('"')[3]
                    max_length = max(max_length, len(display_text))
            else:
                max_length = max(max_length, len(str(cell.value)))
        adjusted_width = max_length + 4
        ws.column_dimensions[get_column_letter(column)].width = adjusted_width

    # Добавляем "кнопку" для печати
    button_cell = ws.cell(row=1, column=6)
    button_cell.value = "ПЕЧАТЬ"
    button_cell.font = Font(bold=True, color="FFFFFF")
    button_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    button_cell.alignment = Alignment(horizontal='center', vertical='center')
    button_cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Устанавливаем ширину столбца для кнопки
    ws.column_dimensions['F'].width = 15

    wb.save(excel_path)
    print(f"Файл Excel обновлен: {excel_path}")

    # Добавляем макрос к кнопке с помощью win32com
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = False
    wb = xl.Workbooks.Open(excel_path)
    ws = wb.Worksheets(sheet_name)
    
    # Добавляем кнопку и привязываем к ней макрос
    button = ws.Buttons().Add(ws.Range("F1").Left, ws.Range("F1").Top, ws.Range("F1").Width, ws.Range("F1").Height)
    button.OnAction = "ПечатьДокументов"
    button.Caption = "ПЕЧАТЬ"
    button.Name = "ПечатьДокументов"

    wb.Save()
    xl.Quit()

    print("Кнопка для запуска макроса печати добавлена.")

def process_photos_and_create_docx(folder_path, folder_name, prefix=""):
    """Обрабатывает фотографии и создает docx файл с добавлением префикса к имени файла."""
    word = win32com.client.Dispatch("Word.Application")
    status = ""

    # Используем имя папки в качестве текста
    text_value = f"{prefix}_{folder_name}" if prefix else folder_name

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

    docx_folder = os.path.join(script_dir, 'docx')
    os.makedirs(docx_folder, exist_ok=True)

    docx_files = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".docx"):
            source_path = os.path.join(folder_path, file_name)
            new_file_name = f"{prefix}_{file_name}" if prefix and not file_name.startswith(prefix) else file_name
            destination_path = os.path.join(docx_folder, new_file_name)
            shutil.move(source_path, destination_path)
            pyperclip.copy(destination_path)
            status += f"Сформирован {new_file_name} и перемещен в папку docx\n"
            docx_files.append(destination_path)

    update_excel(docx_files, prefix)

    return status

def merge_folders(sorted_folder_name):
    clear_terminal()
    sorted_path = os.path.join(os.path.dirname(__file__), sorted_folder_name)
    themes = [f for f in os.listdir(sorted_path) if os.path.isdir(os.path.join(sorted_path, f))]
    
    while True:
        print("Список доступных тем:")
        for i, theme in enumerate(themes, 1):
            print(f"{i}. {theme}")
        print("0. Выход")
        
        theme_choice = input("Выберите номер темы или '0' для выхода: ").strip()
        if theme_choice == '0':
            return 'exit'
        
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
            
            choices = input("Введите номера подпапок для слияния/переименования (через запятую или диапазон, например 1,3-5,7),\n"
                            "'д' для создания docx\n"
                            "'0' для возврата: ").strip()
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
                    
                    print("Выберите папки для создания docx:")
                    for i, folder in enumerate(subfolders, 1):
                        print(f"{i}. {folder}")
                    print("0. Вернуться в предыдущее меню")
                    folder_choices = input("Введите номера папок для создания docx (через запятую или диапазон, например 1,3-5,7) или '0' для возврата: ").strip().lower()
                    
                    if folder_choices == '0':
                        clear_terminal()
                        break
                    
                    folder_indices = parse_folder_selection(folder_choices, len(subfolders))
                    if not folder_indices:
                        print("Неверный выбор. Попробуйте снова.")
                        input("Нажмите Enter для продолжения...")
                        continue
                    
                    prefix = input("Введите префикс для имен файлов (нажмите Enter для пропуска): ").strip()
                    
                    for folder_index in folder_indices:
                        selected_folder = subfolders[folder_index]
                        folder_path = os.path.join(theme_path, selected_folder)
                        
                        status = process_photos_and_create_docx(folder_path, selected_folder, prefix)
                        print(f"Обработка папки '{selected_folder}':")
                        print(status)
                        print("-" * 50)
                    
                    clear_terminal()
                continue

            folder_indices = parse_folder_selection(choices, len(subfolders))
            if not folder_indices:
                print("Неверный выбор. Попробуйте снова.")
                input("Нажмите Enter для продолжения...")
                continue
            
            selected_folders = [subfolders[i] for i in folder_indices]
            
            if len(selected_folders) == 1:
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

            subfolders = sorted([f for f in os.listdir(theme_path) if os.path.isdir(os.path.join(theme_path, f))])

    return None

json_file = os.path.join(os.path.dirname(__file__), 'result.json')

update_themes(json_file)

process_json(json_file)

merge_folders(sorted_folder_name)

print("Программа завершена.")
