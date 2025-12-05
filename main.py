import openpyxl
import re
import xml.etree.ElementTree as ET
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command
import asyncio
import os

TOKEN = os.getenv("TOKEN")  # токен телеграм-бота из Render Environment

bot = Bot(token=TOKEN)
dp = Dispatcher()


# ---------- ЧТЕНИЕ EXCEL ----------
def load_mapping():
    wb = openpyxl.load_workbook("gtin.xlsx")
    ws = wb.active

    mapping = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Берём только первые 3 значения, даже если в Excel больше колонок
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
    cdata_pattern = r"<!\[CDATA\[(.*)\]\]>"
    match = re.search(cdata_pattern, xml_text, re.DOTALL)

    if not match:
        raise ValueError("В файле нет CDATA со вложенным XML.")

    inner_xml = match.group(1).strip()

    inner_tree = ET.ElementTree(ET.fromstring(inner_xml))
    inner_root = inner_tree.getroot()

    products = inner_root.findall(".//product")

    for p in products:
        name_el = p.find("productName")
        if name_el is None:
            continue

        name = name_el.text.strip()

        if name in mapping:
            gtin_value = mapping[name]["gtin"]
            ntin_value = mapping[name]["ntin"]

            # Добавляем перенос строки
            gtin_el = ET.SubElement(p, "gtin")
            gtin_el.text = gtin_value

            ntin_el = ET.SubElement(p, "ntin")
            ntin_el.text = ntin_value

    # Красиво отформатировать XML
    indent(inner_root)

    updated_inner_xml = ET.tostring(inner_root, encoding="unicode")

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
        await message.answer(f"Ошибка обработки XML: {e}")
        return

    output_path = "updated.xml"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(updated_xml)

    await message.answer_document(
        FSInputFile(output_path),
        caption="Готово! Обновлённый XML:"
    )


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())