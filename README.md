# sla_bitrix — SLA-бот Bitrix24

Telegram-бот, который считает SLA по задачам из Bitrix24 с учётом рабочего календаря и праздников и выгружает отчёт в Excel.

## Файлы

- `sla_bitriks.py` — весь проект: Telegram-бот (pyTelegramBotAPI / telebot), запросы к Bitrix24, расчёт SLA по рабочему времени с учётом списка праздников, выгрузка в Excel (pandas).

## Технологии

pyTelegramBotAPI (telebot), pandas, requests, python-dateutil.

## Запуск

```bash
pip install pyTelegramBotAPI pandas requests python-dateutil openpyxl
# задайте токен бота и параметры Bitrix24 через .env
python sla_bitriks.py
```

## ⚠️ Безопасность

Токен Telegram-бота, ранее захардкоженный в коде, заменён на пустой плейсхолдер. Старый токен был опубликован — перевыпустите его в @BotFather и храните новый в `.env`. Список праздников (`HOLIDAYS`) при необходимости обновляйте под текущий год.


## Зависимости

Зависимости проекта вынесены в `requirements.txt`. Установка:

```bash
pip install -r requirements.txt
```
