import os
import json
import asyncio
import time
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.types import InputFile
from aiogram.filters import Command
from tkinter import Tk, filedialog, Button
from PIL import Image, ImageDraw, ImageFont
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger
import numpy as np


# Telegram Bot Token
BOT_TOKEN = "8190891058:AAH4i-tGLp5y_a-YhUH4qvlcTyWRh-J66IU"
GROUP_CHAT_ID = -4527992004  # Укажите ID вашей группы

# Файл для хранения пути и блокировки
CONFIG_FILE = "file_path.json"
LOCK_FILE = "lockfile.lock"

# Создание Telegram-бота
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

# Чтение пути из файла
def read_file_path():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as file:
            data = json.load(file)
            return data.get("path", "")
    return ""

# Запись пути в файл
def save_file_path(path):
    with open(CONFIG_FILE, "w") as file:
        json.dump({"path": path}, file)

# Проверка и управление блокировкой
def is_locked():
    return os.path.exists(LOCK_FILE)

def lock():
    with open(LOCK_FILE, "w") as lock_file:
        lock_file.write(str(os.getpid()))

def unlock():
    if is_locked():
        os.remove(LOCK_FILE)

def is_lock_expired():
    if not is_locked():
        return False
    lock_time = os.path.getmtime(LOCK_FILE)
    return (time.time() - lock_time) > 1  # 15 секунд

def handle_expired_lock():
    if is_lock_expired():
        print("Старая блокировка удалена.")
        unlock()

# Преобразование Excel в изображение
def excel_to_image(file_path):
    try:
        df = pd.read_excel(file_path)

        # Заменяем названия столбцов на A, B, C, ...
        df.columns = [chr(65 + i) for i in range(len(df.columns))]

        # Заменяем NaN на пустые строки
        df = df.replace(np.nan, "")

        # Приводим числовые значения к целым там, где это возможно
        df = df.applymap(lambda x: int(x) if isinstance(x, float) and x.is_integer() else x)

        # Настройки для отображения
        font_path = "C:/Windows/Fonts/arial.ttf"
        font = ImageFont.truetype(font_path, 16)
        padding = 20
        row_height = 40
        col_padding = 10

        max_col_widths = [
            max(len(str(val)) for val in df[col].values) for col in df.columns
        ]
        max_col_widths = [
            max(len(str(col)), width) + 2 for col, width in zip(df.columns, max_col_widths)
        ]
        col_widths = [width * col_padding for width in max_col_widths]

        image_width = sum(col_widths) + padding * 2
        image_height = (len(df) + 1) * row_height + padding * 2

        image = Image.new("RGB", (image_width, image_height), "white")
        draw = ImageDraw.Draw(image)

        y = padding
        x_start = padding

        for col_idx, col in enumerate(df.columns):
            col_width = col_widths[col_idx]
            x_end = x_start + col_width

            draw.rectangle(
                [(x_start, y), (x_end, y + row_height)],
                fill="lightgrey",
                outline="black",
                width=1,
            )

            bbox = draw.textbbox((0, 0), str(col), font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            text_x = x_start + (col_width - text_width) / 2
            text_y = y + (row_height - text_height) / 2
            draw.text((text_x, text_y), str(col), font=font, fill="black")

            x_start = x_end

        y += row_height
        for row in df.values.tolist():
            x_start = padding
            for col_idx, cell in enumerate(row):
                col_width = col_widths[col_idx]
                x_end = x_start + col_width

                draw.rectangle(
                    [(x_start, y), (x_end, y + row_height)],
                    outline="black",
                    width=1,
                )

                cell_text = str(cell)
                bbox = draw.textbbox((0, 0), cell_text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]
                text_x = x_start + (col_width - text_width) / 2
                text_y = y + (row_height - text_height) / 2
                draw.text((text_x, text_y), cell_text, font=font, fill="black")
                x_start = x_end
            y += row_height

        temp_image_path = "temp_image.png"
        image.save(temp_image_path)
        return temp_image_path
    except Exception as e:
        print(f"Ошибка: {e}")
        return str(e)

@dp.message(Command("start"))
async def start(message: types.Message):
    file_path = read_file_path()
    if not file_path or not os.path.exists(file_path):
        await message.reply("Файл не выбран или не существует. Пожалуйста, выберите файл через GUI.")
    else:
        await message.reply("Бот готов к работе. Используйте команду /sendexsel для отправки файла в группу.")

@dp.message(Command("sendexsel"))
async def sendexsel(message: types.Message):
    await send_excel_to_group()

async def send_excel_to_group():
    try:
        handle_expired_lock()
        if is_locked():
            print("Отправка уже выполняется другим процессом. Ждем 10 секунд...")
            await asyncio.sleep(1)
            return await send_excel_to_group()

        lock()

        file_path = read_file_path()
        if not file_path or not os.path.exists(file_path):
            print("Ошибка: файл не выбран или не существует.")
            return

        image_path = excel_to_image(file_path)
        if os.path.exists(image_path):
            from aiogram.types import FSInputFile
            await bot.send_photo(GROUP_CHAT_ID, photo=FSInputFile(image_path))
            os.remove(image_path)
            print("Файл успешно отправлен.")
        else:
            print(f"Ошибка преобразования файла: {image_path}")
    except Exception as e:
        print(f"Ошибка при отправке: {e}")
    finally:
        unlock()

def gui():
    def choose_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            save_file_path(file_path)
            print(f"Выбран файл: {file_path}")

    def stop_bot():
        root.destroy()
        os._exit(0)

    root = Tk()
    root.title("Выбор файла")
    Button(root, text="Выбрать файл", command=choose_file).pack(pady=10)
    Button(root, text="Остановить бота", command=stop_bot).pack(pady=10)
    root.mainloop()

async def main():
    dp.message.register(start)
    dp.message.register(sendexsel)

    scheduler = AsyncIOScheduler(timezone="Europe/Moscow")
    scheduler.add_job(send_excel_to_group, CronTrigger(hour=18, minute=8))
    scheduler.start()

    import threading
    threading.Thread(target=gui, daemon=True).start()

    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
