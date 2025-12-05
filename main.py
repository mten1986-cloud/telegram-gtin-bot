import openpyxl
import re
import xml.etree.ElementTree as ET
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command
import asyncio
import os
from aiohttp import web

# ✔ Сохраняем namespace v1:snt
ET.register_namespace("v1", "v1.snt")

TOKEN = os.getenv("TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher()


# ---------- ЧТЕНИЕ EXCEL ----------
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


# ---------- КРАСИВОЕ ФОРМАТИРОВАНИЕ XML ----------
def indent(elem, level=0):
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


# ---------- ОБРАБОТКА CDATA И ДОБАВЛЕНИЕ GTIN/NTIN ----------
def process_xml_with_cdata(xml_text: str) -> str:
    """Извлекает XML внутри CDATA, модифицирует его, возвращает обратно."""
    
    cdata_pattern = r"<!\[CDATA\[(.*)\]\]>"
    match = re.search(cdata_pattern, xml_text, re.DOTALL)

    if not match:
        raise ValueError("В файле нет CDATA со вложенным XML.")

    inner_xml = match.group(1).strip()

    # Парсим с учётом namespace
    root = ET.fromstring(inner_xml)
    tree = ET.ElementTree(root)

    # Путь поиска тегов в пространстве имён v1
    ns = {"v1": "v1.snt"}

    products = root.findall(".//product", namespaces=ns) or root.findall(".//product")

    for p in products:
        name_el = p.find("productName")
        if name_el is None:
            continue

        name = name_el.text.strip()

        if name in mapping:
            gtin_el = ET.SubElement(p, "gtin")
            gtin_el.text = mapping[name]["gtin"]

            ntin_el = ET.SubElement(p, "ntin")
            ntin_el.text = mapping[name]["ntin"]

    # Красивое форматирование
    indent(root)
    updated_inner_xml = ET.tostring(root, encoding="unicode")

    updated_cdata = f"<![CDATA[\n{updated_inner_xml}\n]]>"
    final_xml = re.sub(cdata_pattern, updated_cdata, xml_text, flags=re.DOTALL)

    return final_xml


# ---------- TELEGRAM BOT ----------
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


# ---------- AIOHTTP SERVER (для Render Free) ----------
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


# ---------- MAIN ----------
async def main():
    await start_web_server()   # важно, чтобы Render видел порт
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
