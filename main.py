import openpyxl
import re
import xml.etree.ElementTree as ET
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command
import asyncio
import os
from aiohttp import web

# ✔ Сохраняем namespace v1 для корректного вывода
ET.register_namespace("v1", "v1.snt")

TOKEN = os.getenv("TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher()


# ------------------ ЧТЕНИЕ EXCEL ------------------
def load_mapping():
    wb = openpyxl.load_workbook("gtin.xlsx")
    ws = wb.active

    mapping = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        productName, gtin, ntin = row[:3]

        if productName:
            mapping[str(productName).strip()] = {
                "gtin": str(gtin) if gtin else "",
                "ntin": str(ntin) if ntin else ""
            }

    return mapping


mapping = load_mapping()


# ------------------ XML FORMATTER ------------------
def indent(elem, level=0):
    """Красивое форматирование XML"""
    i = "\n" + level * "  "
    if len(elem):
        if not elem.text or not elem.text.strip():
            elem.text = i + "  "
        for child in elem:
            indent(child, level + 1)
        if not elem.tail or not elem.tail.strip():
            elem.tail = i
    else:
        if level and (not elem.tail or not elem.tail.strip()):
            elem.tail = i


# ------------------ ОБРАБОТКА XML ------------------
def process_xml_with_cdata(xml_text: str) -> str:

    # ---- СОХРАНЯЕМ XML-ДЕКЛАРАЦИЮ ----
    xml_decl = ""
    if xml_text.strip().startswith("<?xml"):
        end_decl = xml_text.find("?>") + 2
        xml_decl = xml_text[:end_decl].strip()
        xml_text_no_decl = xml_text[end_decl:].lstrip()
    else:
        xml_text_no_decl = xml_text

    # ---- ИЗВЛЕКАЕМ CDATA ----
    cdata_pattern = r"<!\[CDATA\[(.*)\]\]>"
    match = re.search(cdata_pattern, xml_text_no_decl, re.DOTALL)

    if not match:
        raise ValueError("В файле нет CDATA со вложенным XML.")

    inner_xml = match.group(1).strip()

    # ---- ПАРСИМ ВНУТРЕННИЙ XML ----
    root = ET.fromstring(inner_xml)

    # namespace
    ns = {"v1": "v1.snt"}

    # Находим <product>
    products = root.findall(".//product", ns) or root.findall(".//product")

    # ---- ДОБАВЛЯЕМ GTIN/NTIN ----
    for p in products:
        name_el = p.find("productName")
        if not name_el:
            continue

        name = name_el.text.strip()

        if name in mapping:
            gtin_el = ET.SubElement(p, "gtin")
            gtin_el.text = mapping[name]["gtin"]

            ntin_el = ET.SubElement(p, "ntin")
            ntin_el.text = mapping[name]["ntin"]

    # ---- ФОРМАТИРУЕМ ----
    indent(root)

    # Получаем XML как строку
    updated_inner_xml = ET.tostring(root, encoding="unicode")

    # ---- ВОССТАНАВЛИВАЕМ ПРЕФИКС v1 ----
    updated_inner_xml = updated_inner_xml.replace("ns0:", "v1:")
    updated_inner_xml = updated_inner_xml.replace('xmlns:ns0="v1.snt"', 'xmlns:v1="v1.snt"')

    # ---- ВСТАВЛЯЕМ ОБРАТНО В CDATA ----
    updated_cdata = f"<![CDATA[\n{updated_inner_xml}\n]]>"
    final_xml = re.sub(cdata_pattern, updated_cdata, xml_text_no_decl, flags=re.DOTALL)

    # ---- ВОССТАНАВЛИВАЕМ XML-ДЕКЛАРАЦИЮ ----
    if xml_decl:
        final_xml = xml_decl + "\n" + final_xml

    return final_xml


# ------------------ TELEGRAM BOT ------------------
@dp.message(Command("start"))
async def start(message: types.Message):
    await message.answer("Отправьте XML-файл — я добавлю внутрь GTIN и NTIN.")


@dp.message(lambda msg: msg.document and msg.document.file_name.endswith(".xml"))
async def handle_xml(message: types.Message):
    file = await bot.download(message.document)
    original_xml = file.read().decode("utf-8")

    try:
        updated_xml = process_xml_with_cdata(original_xml)
    except Exception as e:
        await message.answer(f"Ошибка: {e}")
        return

    output_path = "updated.xml"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(updated_xml)

    await message.answer_document(FSInputFile(output_path), caption="Готово! Обновлённый XML:")


# ------------------ WEB SERVER ДЛЯ RENDER (FREE) ------------------
async def handle(request):
    return web.Response(text="Bot is running.")


async def start_web_server():
    app = web.Application()
    app.router.add_get("/", handle)

    port = int(os.getenv("PORT", 8080))
    runner = web.AppRunner(app)
    await runner.setup()

    site = web.TCPSite(runner, "0.0.0.0", port)
    await site.start()


# ------------------ MAIN ------------------
async def main():
    await start_web_server()   # важно для Render Free
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

