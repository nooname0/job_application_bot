import asyncio
import logging

from aiogram import Bot, Dispatcher, types
from aiogram.filters import Command
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton, FSInputFile
from openpyxl import Workbook
import os
# ================== SOZLAMALAR ==================
BOT_TOKEN = "7966503499:AAHlh6Y4KwsQOdUc13MAMOIzJq9OuyGFEjI"
ADMIN_ID = 6140962854  # @userinfobot orqali tekshirilgan

logging.basicConfig(level=logging.INFO)

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()

applications = []
user_step = {}
user_data = {}

FILIALS = [
    "Niyazbosh",
    "Olmazor",
    "Chinoz",
    "Kasblar",
    "Gulbahor",
    "Konditeriski",
    "Mevazor"
]

steps = [
    "Lavozimni kiriting:",
    "F.I.SH:",
    "Tug‘ilgan yil:",
    "Tug‘ilgan oy:",
    "Tug‘ilgan kun:",
    "Otasi F.I.SH:",
    "Otasi tug‘ilgan yil:",
    "Otasi tug‘ilgan oy:",
    "Otasi tug‘ilgan kun:",
    "Onasi F.I.SH:",
    "Onasi tug‘ilgan yil:",
    "Onasi tug‘ilgan oy:",
    "Onasi tug‘ilgan kun:",
    "Telefon raqam (hodimniki):"
]

keys = [
    "lavozim", "fio",
    "tyil", "toy", "tkun",
    "ofio", "oyil", "ooy", "okun",
    "mfio", "myil", "moy", "mkun",
    "phone"
]

# ================== KLAWITURA ==================
def filial_keyboard():
    return InlineKeyboardMarkup(
        inline_keyboard=[
            [InlineKeyboardButton(text=f, callback_data=f"filial:{f}")]
            for f in FILIALS
        ]
    )

# ================== START ==================
@dp.message(Command("start"))
async def start(message: types.Message):
    user_step[message.chat.id] = 0
    user_data[message.chat.id] = {}

    await message.answer(
        "Filialni tanlang:",
        reply_markup=filial_keyboard()
    )

# ================== EXCEL (MUHIM: TEPADA) ==================
@dp.message(Command("excel"))
async def export_excel(message: types.Message):
    if message.chat.id != ADMIN_ID:
        await message.answer("⛔ Siz admin emassiz")
        return

    wb = Workbook()
    ws = wb.active

    ws.append([
        "№","Filial","Lavozim","F.I.SH",
        "Tug'ilgan yil","Oy","Kun",
        "Otasi F.I.SH","Otasi yil","Oy","Kun",
        "Onasi F.I.SH","Onasi yil","Oy","Kun",
        "Telefon"
    ])

    for i, app in enumerate(applications, 1):
        ws.append([
            i,
            app["filial"], app["lavozim"], app["fio"],
            app["tyil"], app["toy"], app["tkun"],
            app["ofio"], app["oyil"], app["ooy"], app["okun"],
            app["mfio"], app["myil"], app["moy"], app["mkun"],
            app["phone"]
        ])

    file_name = "arizalar.xlsx"
    wb.save(file_name)

    await message.answer_document(FSInputFile(file_name))

# ================== FILIAL TANLASH ==================
@dp.callback_query(lambda c: c.data.startswith("filial:"))
async def filial_chosen(callback: types.CallbackQuery):
    chat_id = callback.message.chat.id
    filial = callback.data.split(":")[1]

    user_data[chat_id]["filial"] = filial
    user_step[chat_id] = 0

    await callback.message.edit_text(f"✅ Tanlangan filial: {filial}")
    await bot.send_message(chat_id, steps[0])
    await callback.answer()

# ================== FORMA (ENG OXIRDA) ==================
@dp.message()
async def form_handler(message: types.Message):
    if message.text.startswith("/"):
        return

    chat_id = message.chat.id

    if chat_id not in user_step:
        return

    step = user_step[chat_id]

    user_data[chat_id][keys[step]] = message.text
    step += 1

    if step < len(steps):
        user_step[chat_id] = step
        await message.answer(steps[step])
    else:
        applications.append(user_data[chat_id])
        await message.answer("✅ Arizangiz qabul qilindi.")
        del user_step[chat_id]
        del user_data[chat_id]

# ================== RUN ==================
async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
