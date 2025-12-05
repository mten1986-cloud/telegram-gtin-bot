import pandas as pd
import re
import xml.etree.ElementTree as ET
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command
import asyncio
import os

TOKEN = os.getenv("TOKEN")  # токен из переменных окружения на Render

bot = Bot(token=TOKEN)
dp = Dispatcher()

# Загружаем Excel файл (лежит в той же директории, где main.py)
mapping = pd.read_excel("gtin.xlsx")
mapping.columns = ["productName", "gtin", "ntin"]


def process_xml_with_cdata(xml_text: str) -> str:
    """
    1) Находит CDATA с внутренним XML
    2) Извлекает XML
    3) Добавляет gtin и ntin в каждый <product>
    4) Собирает XML обратно
    """

    cdata_pattern = r"<!\[CDATA\[(.*)\]\]>"
    match = re.search(cdata_pattern, xml_text, re.DOTALL)

    if not match:
        raise ValueError("В исходном файле нет CDATA с вложенным XML.")

    inner_xml = match.group(1).strip()

    # Парсим внутренний XML
    inner_tree = ET.ElementTree(ET.fromstring(inner_xml))
    inner_root = inner_tree.getroot()

    products = inner_root.findall(".//product")

    for p in products:
        name_el = p.find("productName")
        if name_el is None:
            continue

        name = name_el.text.strip()
        row = mapping[mapping["productName"] == name]

        if not row.empty:
            gtin_el = ET.Element("gtin")
            gtin_el.text = str(row.iloc[0]["gtin"])
            p.append(gtin_el)

            ntin_el = ET.Element("ntin")
            ntin_el.text = str(row.iloc[0]["ntin"])
            p.append(ntin_el)

    updated_inner_xml = ET.tostring(inner_root, encoding="unicode")

    updated_cdata = f"<![CDATA[\n{updated_inner_xml}\n]]>"

    final_xml = re.sub(cdata_pattern, updated_cdata, xml_text, flags=re.DOTALL)

    return final_xml


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

    await message.answer_document(FSInputFile(output_path), caption="Готово! Ваш файл обновлён.")


async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())