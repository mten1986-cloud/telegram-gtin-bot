import openpyxl
import re
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command
import asyncio
import os
from aiohttp import web

TOKEN = os.getenv("TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher()


# ---------------------------------------------------------
#                 ЧТЕНИЕ EXCEL МАППИНГА
# ---------------------------------------------------------
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


# ---------------------------------------------------------
#       ОБРАБОТКА ТОЛЬКО <product>...</product>
# ---------------------------------------------------------
def process_xml_with_cdata(xml_text: str) -> str:
    # ---- СОХРАНЯЕМ XML-ДЕКЛАРАЦИЮ ----
    xml_decl = ""
    if xml_text.strip().startswith("<?xml"):
        end = xml_text.find("?>") + 2
        xml_decl = xml_text[:end].strip()
        xml_text_no_decl = xml_text[end:].lstrip()
    else:
        xml_text_no_decl = xml_text

    # ---- ИЗВЛЕКАЕМ CDATA ----
    cdata_pattern = r"<!\[CDATA\[(.*)\]\]>"
    match = re.search(cdata_pattern, xml_text_no_decl, re.DOTALL)
    if not match:
        raise ValueError("CDATA не найден.")

    inner_xml = match.group(1)

    # ---- ФУНКЦИЯ ДЛЯ ИЗМЕНЕНИЯ ОДНОГО PRODUCT ----
    def replace_product(block):
        product_block = block.group(1)

        # Ищем productName
        m = re.search(r"<productName>(.*?)</productName>", product_block, re.DOTALL)
        if not m:
            return product_block

        name = m.group(1).strip()
        if name not in mapping:
            return product_block

        gtin = mapping[name]["gtin"]
        ntin = mapping[name]["ntin"]

        # ---- вычисляем отступы ----
        # отступ перед </product>
        m_indent = re.search(r"(\s*)</product>", product_block)
        indent = m_indent.group(1) if m_indent else ""

        # отступ для внутренних тегов: на один уровень глубже
        inner_indent = indent + "    "

        # ---- вставляем gtin/ntin с правильными отступами ----
        updated = re.sub(
            r"</product>",
            f"{inner_indent}<gtin>{gtin}</gtin>\n"
            f"{inner_indent}<ntin>{ntin}</ntin>\n"
            f"{indent}</product>",
            product_block
        )

        return updated

    # ---- МЕНЯЕМ ТОЛЬКО PRODUCT ----
    product_pattern = r"(<product[\s\S]*?</product>)"
    updated_inner_xml = re.sub(product_pattern, lambda m: replace_product(m), inner_xml)

    # ---- СОБИРАЕМ НОВУЮ CDATA ----
    updated_cdata = f"<![CDATA[{updated_inner_xml}]]>"
    final_xml = re.sub(cdata_pattern, updated_cdata, xml_text_no_decl, flags=re.DOTALL)

    # ---- ВОССТАНАВЛИВАЕМ XML-ДЕКЛАРАЦИЮ ----
    if xml_decl:
        final_xml = xml_decl + "\n" + final_xml

    return final_xml


# ---------------------------------------------------------
#                   TELEGRAM BOT HANDLERS
# ---------------------------------------------------------
@dp.message(Command("start"))
async def start(message: types.Message):
    await message.answer("Отправьте XML-файл — я добавлю GTIN и NTIN внутрь <product>.")


@dp.message(lambda msg: msg.document and msg.document.file_name.endswith(".xml"))
async def handle_xml(message: types.Message):

    original_filename = message.document.file_name  # сохраняем имя файла

    # скачиваем файл
    file = await bot.download(message.document)
    xml_text = file.read().decode("utf-8")

    # обрабатываем XML
    try:
        updated_xml = process_xml_with_cdata(xml_text)
    except Exception as e:
        await message.answer(f"Ошибка обработки XML: {e}")
        return

    # сохраняем с тем же именем
    with open(original_filename, "w", encoding="utf-8") as f:
        f.write(updated_xml)

    # отправляем обратно под тем же именем
    await message.answer_document(
        FSInputFile(original_filename),
        caption="Готово! Обновлённый файл."
    )


# ---------------------------------------------------------
#      ВЕБ-СЕРВЕР ДЛЯ RENDER (БЕСПЛАТНЫЙ ТАРИФ)
# ---------------------------------------------------------
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


# ---------------------------------------------------------
#                        MAIN
# ---------------------------------------------------------
async def main():
    await start_web_server()   # важно для Render Free
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())

