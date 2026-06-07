import telebot
import pandas as pd
import os
import requests
from datetime import datetime, time, timedelta
from dateutil import parser
import time as time_module

API_TOKEN = ''
bot = telebot.TeleBot(API_TOKEN)

HOLIDAYS = [
    datetime(2025, 1, 1).date(),
    datetime(2025, 1, 2).date(),
    datetime(2025, 1, 3).date(),
    datetime(2025, 1, 4).date(),
    datetime(2025, 1, 5).date(),
    datetime(2025, 1, 6).date(),
    datetime(2025, 1, 7).date(),
    datetime(2025, 1, 8).date(),
    datetime(2025, 2, 23).date(),
    datetime(2025, 2, 24).date(),
    datetime(2025, 3, 8).date(),
    datetime(2025, 3, 9).date(),
    datetime(2025, 5, 1).date(),
    datetime(2025, 5, 2).date(),
    datetime(2025, 5, 9).date(),
    datetime(2025, 6, 12).date(),
    datetime(2025, 6, 13).date(),
    datetime(2025, 11, 3).date(),
    datetime(2025, 11, 4).date(),
]

WEBHOOK = 'https://mospoly.bitrix24.ru/rest/66740/1pjeusppmp1w81v2/'

@bot.message_handler(commands=['start'])
def send_welcome(message):
    bot.reply_to(message, "Привет! Пришли мне файл (.xlsx/.xls/.html/.htm), и я обработаю его содержимое.")

def detect_file_type_and_read(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()

    try:
        with open(file_path, 'rb') as f:
            first_bytes = f.read(8192)
    except Exception as e:
        raise Exception(f"Не удалось прочитать файл: {str(e)}")

    if first_bytes.startswith(b'\xef\xbb\xbf'):
        first_bytes_without_bom = first_bytes[3:]
    elif first_bytes.startswith(b'\xff\xfe'):
        first_bytes_without_bom = first_bytes[2:]
    elif first_bytes.startswith(b'\xfe\xff'):
        first_bytes_without_bom = first_bytes[2:]
    else:
        first_bytes_without_bom = first_bytes

    first_bytes_lower = first_bytes_without_bom[:200].lower()
    if (b'<html' in first_bytes_lower or
            b'<!doctype html' in first_bytes_lower or
            b'<table' in first_bytes_lower or
            b'<meta' in first_bytes_lower):

        print(f"Обнаружен HTML-файл (расширение: {file_extension})")
        try:
            dfs = None
            for encoding in ['utf-8', 'cp1251', 'latin1']:
                try:
                    dfs = pd.read_html(file_path, encoding=encoding)
                    if len(dfs) > 0:
                        print(f"HTML успешно прочитан с кодировкой: {encoding}")
                        break
                except:
                    continue

            if dfs is None or len(dfs) == 0:
                raise Exception("HTML файл не содержит читаемых таблиц")

            df = dfs[0]
            return df, 'html'
        except Exception as e:
            raise Exception(f"Ошибка чтения HTML: {str(e)}")

    if first_bytes.startswith(b'PK\x03\x04'):
        print(f"Обнаружен XLSX-файл")
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            return df, 'excel'
        except Exception as e:
            raise Exception(f"Ошибка чтения XLSX: {str(e)}")

    if first_bytes.startswith(b'\xd0\xcf\x11\xe0'):
        print(f"Обнаружен XLS-файл (старый формат)")
        try:
            df = pd.read_excel(file_path, engine='xlrd')
            return df, 'excel'
        except Exception as e:
            raise Exception(f"Ошибка чтения XLS: {str(e)}")

    if file_extension in ['.xlsx', '.xls']:
        print(f"Попытка чтения как Excel по расширению: {file_extension}")
        try:
            if file_extension == '.xlsx':
                df = pd.read_excel(file_path, engine='openpyxl')
            else:
                df = pd.read_excel(file_path, engine='xlrd')
            return df, 'excel'
        except Exception as e:
            print(f"Excel не удалось прочитать, пробуем как HTML...")
            try:
                dfs = pd.read_html(file_path)
                if len(dfs) == 0:
                    raise Exception("Нет таблиц")
                print(f"Успешно прочитан как HTML")
                return dfs[0], 'html'
            except:
                raise Exception(f"Файл не удалось прочитать ни как Excel, ни как HTML: {str(e)}")

    if file_extension in ['.html', '.htm']:
        print(f"Чтение HTML по расширению")
        try:
            dfs = pd.read_html(file_path)
            if len(dfs) == 0:
                raise Exception("HTML файл не содержит таблиц")
            return dfs[0], 'html'
        except Exception as e:
            raise Exception(f"Ошибка чтения HTML: {str(e)}")

    raise Exception(f"Неподдерживаемый формат файла. Первые байты: {first_bytes[:20]}")

def get_working_hours(date_obj):
    if date_obj in HOLIDAYS:
        return None

    month = date_obj.month
    weekday = date_obj.weekday()

    if month in [11, 12, 1, 2, 3]:
        if weekday < 5:
            return (time(9, 0), time(20, 0))
        else:
            return None

    elif month in [4, 5, 6]:
        if weekday < 5:
            return (time(9, 0), time(20, 0))
        else:
            return (time(9, 0), time(18, 0))

    elif month in [7, 8]:
        return (time(9, 0), time(20, 0))

    elif month in [9, 10]:
        if weekday < 5:
            return (time(9, 0), time(20, 0))
        else:
            return (time(9, 0), time(18, 0))

    return (time(9, 0), time(18, 0))

def working_time_diff(start_str, end_str):
    try:
        start_dt = parser.parse(start_str)
        end_dt = parser.parse(end_str)

        if end_dt < start_dt:
            start_dt, end_dt = end_dt, start_dt

        total = timedelta(0)
        current_day = start_dt.date()

        while current_day <= end_dt.date():
            working_hours = get_working_hours(current_day)

            if working_hours is None:
                current_day += timedelta(days=1)
                continue

            work_start, work_end = working_hours

            work_start_dt = datetime.combine(current_day, work_start, tzinfo=start_dt.tzinfo)
            work_end_dt = datetime.combine(current_day, work_end, tzinfo=start_dt.tzinfo)

            interval_start = max(start_dt, work_start_dt)
            interval_end = min(end_dt, work_end_dt)

            if interval_start < interval_end:
                total += (interval_end - interval_start)

            current_day += timedelta(days=1)

        return total

    except Exception as e:
        print(f"Ошибка в working_time_diff: {e}")
        return timedelta(0)

def bitrix_request_with_retry(url, params, max_retries=3, pause=0.5):
    """
    Отправляет запрос к Bitrix24 API с повторными попытками при ошибках.
    Возвращает (response_json, error_string).
    """
    for attempt in range(1, max_retries + 1):
        try:
            response = requests.post(url, params=params, timeout=30)

            if response.status_code in (429, 503):
                wait = 2 * attempt
                print(f"Rate limit (HTTP {response.status_code}), жду {wait}с (попытка {attempt}/{max_retries})")
                time_module.sleep(wait)
                continue

            if response.status_code != 200:
                return None, f"HTTP {response.status_code}"

            data = response.json()

            if 'error' in data:
                error_code = data.get('error', '')
                error_desc = data.get('error_description', '')

                if 'QUERY_LIMIT_EXCEEDED' in str(error_code).upper():
                    wait = 2 * attempt
                    print(f"QUERY_LIMIT_EXCEEDED, жду {wait}с (попытка {attempt}/{max_retries})")
                    time_module.sleep(wait)
                    continue

                return None, f"API error: {error_code} — {error_desc}"

            return data, None

        except requests.exceptions.Timeout:
            print(f"Timeout, попытка {attempt}/{max_retries}")
            time_module.sleep(1)

        except requests.exceptions.RequestException as e:
            print(f"Сетевая ошибка: {e}, попытка {attempt}/{max_retries}")
            time_module.sleep(1)

    return None, "Все попытки исчерпаны"

@bot.message_handler(content_types=['document'])
def handle_docs(message):
    try:
        file_info = bot.get_file(message.document.file_id)
        file_name = message.document.file_name

        supported_extensions = ['.xlsx', '.xls', '.html', '.htm']
        file_extension = os.path.splitext(file_name)[1].lower()

        if file_extension not in supported_extensions:
            bot.reply_to(message,
                         "Поддерживаемые форматы: Excel (.xlsx, .xls) и HTML (.html, .htm)")
            return

        downloaded_file = bot.download_file(file_info.file_path)

        with open(file_name, 'wb') as new_file:
            new_file.write(downloaded_file)

        bot.reply_to(message, f"Файл '{file_name}' получен. Определяю тип и читаю содержимое... ")

        try:
            df, file_type = detect_file_type_and_read(file_name)
            bot.send_message(message.chat.id, f"Файл определен как {file_type.upper()} и успешно прочитан! ")

        except Exception as e:
            bot.send_message(message.chat.id, f"Ошибка чтения файла: {str(e)} ")
            return

        if '№' not in df.columns:
            possible_columns = ['№', 'ID', 'Id', 'id', 'номер', 'Номер', 'NUMBER']
            found_column = None

            for col in possible_columns:
                if col in df.columns:
                    found_column = col
                    break

            if not found_column:
                available_columns = ', '.join(df.columns.tolist())
                bot.send_message(message.chat.id,
                                 f"Столбец '№' не найден. Доступные столбцы: {available_columns}")
                return
            else:
                df = df.rename(columns={found_column: '№'})
                bot.send_message(message.chat.id, f"Использую столбец '{found_column}' как '№'")

        session = []
        for index, row in df.iterrows():
            session.append(row.to_dict())

        operators = {}

        skipped_nan = 0
        skipped_api_error = 0
        skipped_no_result = 0
        skipped_no_dates = 0
        processed_ok = 0

        bot.send_message(message.chat.id, f"Начинаю обработку {len(session)} записей... ")

        for idx, i in enumerate(session):
            try:
                if pd.isna(i['№']):
                    skipped_nan += 1
                    continue

                line_id = str(i['№']).strip()
                if not line_id or line_id == 'nan':
                    skipped_nan += 1
                    continue

                session_id = line_id
                params = {'SESSION_ID': session_id}

                data_json, error = bitrix_request_with_retry(
                    WEBHOOK + 'imopenlines.session.history.get',
                    params,
                    max_retries=3,
                    pause=0.5
                )

                time_module.sleep(0.5)

                if data_json is None:
                    print(f"Сессия {session_id} пропущена: {error}")
                    skipped_api_error += 1
                    continue

                if 'result' not in data_json or 'message' not in data_json['result']:
                    skipped_no_result += 1
                    continue

                data = data_json['result']['message']
                keys_list = list(data.keys())

                data1 = 0
                data2 = 0
                datap = 0
                oper = "Неизвестный оператор"

                for key in keys_list:
                    if 'Передаю ваш вопрос оператору.' in data[key]['text']:
                        data1 = data[key]['date']
                    if 'забрал диалог у' in data[key]['text'] or 'начал работу с диалогом' in data[key]['text']:
                        data2 = datap
                        try:
                            oper = data[key]['text'].split(']')[1].split('[')[0].strip()
                        except IndexError:
                            oper = "Неизвестный оператор"
                    if data1 != 0 and data2 != 0:
                        break
                    datap = data[key]['date']

                if data1 == 0 or data2 == 0:
                    data1 = 0
                    data2 = 0
                    datap = 0
                    for key in keys_list:
                        if 'Обращение направлено' in data[key]['text']:
                            data1 = data[key]['date']
                        if ('начал работу с диалогом' in data[key]['text'] or
                                'начала работу с диалогом' in data[key]['text']):
                            data2 = datap
                            try:
                                oper = data[key]['text'].split(']')[1].split('[')[0].strip()
                            except IndexError:
                                oper = "Неизвестный оператор"
                        if data1 != 0 and data2 != 0:
                            break
                        datap = data[key]['date']

                if data1 != 0 and data2 != 0:
                    diff = working_time_diff(data1, data2)

                    if oper in operators:
                        operators[oper]['count'] += 1
                        operators[oper]['sec'] += diff.total_seconds()
                    else:
                        operators[oper] = {
                            'count': 1,
                            'sec': diff.total_seconds()
                        }
                    processed_ok += 1
                else:
                    skipped_no_dates += 1

                if (idx + 1) % max(1, round(len(session) * 0.1)) == 0:
                    percentage = round(((idx + 1) / len(session)) * 100)
                    bot.send_message(message.chat.id,
                                     f'Прогресс: {percentage}% ({idx + 1}/{len(session)}) ')

            except Exception as e:
                print(f"Ошибка обработки записи {idx}: {e}")
                skipped_api_error += 1
                continue

        bot.send_message(message.chat.id,
                         f"Диагностика обработки:\n"
                         f"   Всего строк в файле: {len(session)}\n"
                         f"   Успешно обработано: {processed_ok}\n"
                         f"   Пропущено (пустые/NaN): {skipped_nan}\n"
                         f"   Ошибки API / сеть: {skipped_api_error}\n"
                         f"   Нет result/message: {skipped_no_result}\n"
                         f"   Не найдены даты: {skipped_no_dates}")

        if not operators:
            bot.send_message(message.chat.id, "Не удалось найти данные для анализа. ")
            return

        total_sec = 0
        total_count = 0

        bot.send_message(message.chat.id, "**Результаты анализа:**\n")

        for oper_name in sorted(operators.keys()):
            stats = operators[oper_name]
            if stats['count'] > 0:
                avg_seconds = stats['sec'] / stats['count']
                td = timedelta(seconds=avg_seconds)
                total_seconds_int = int(td.total_seconds())
                hours = total_seconds_int // 3600
                minutes = (total_seconds_int % 3600) // 60
                secs = total_seconds_int % 60

                total_sec += stats['sec']
                total_count += stats['count']

                bot.send_message(message.chat.id,
                                 f"{oper_name}\n"
                                 f"   Обработано: {stats['count']} сессий\n"
                                 f"   Среднее время: {hours:02d}:{minutes:02d}:{secs:02d}")

        if total_count > 0:
            td = timedelta(seconds=total_sec / total_count)
            total_seconds_int = int(td.total_seconds())
            hours = total_seconds_int // 3600
            minutes = (total_seconds_int % 3600) // 60
            secs = total_seconds_int % 60
            bot.send_message(message.chat.id,
                             f"\n**ИТОГО:**\n"
                             f"   Всего сессий: {total_count}\n"
                             f"   Среднее время по всем: {hours:02d}:{minutes:02d}:{secs:02d}")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {str(e)} ")
        print(f"Общая ошибка: {e}")

    finally:
        if 'file_name' in locals() and os.path.exists(file_name):
            os.remove(file_name)

if __name__ == "__main__":
    print("Бот запущен и поддерживает форматы: Excel (.xlsx, .xls) и HTML (.html, .htm)")
    bot.polling(none_stop=True)
