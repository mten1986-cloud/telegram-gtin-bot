import openpyxl
import re
import xml.etree.ElementTree as ET
from aiogram import Bot, Dispatcher, types
from aiogram.types import FSInputFile
from aiogram.filters import Command
import asyncio
import os
from aiohttp import web

TOKEN = os.getenv("TOKEN")

bot = Bot(token=TOKEN)
dp = Dispatcher()

# ---------- РЕГИСТРАЦИЯ НAMESPACE ДЛЯ СОХРАНЕНИЯ <v1:snt> ----------
ET.register_namespace("v1", "v1.snt")


# ---------- ЧТЕНИЕ EXCEL ----------
def load_mapping():
    wb = openpyxl.load_workbook("gtin.xlsx")
    ws = wb.active

    mapping = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        productName, gtin, ntin = row[:3]

        if productName:
            mapping[str(productName).strip()] = {
                "gtin": str(gtin).strip() if gtin else "",
                "ntin": str(ntin).strip() if ntin else ""
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

    # Парсим с учётом namespace
    inner_tree = ET.ElementTree(ET.fromstring(inner_xml))
    inner_root = inner_tree.getroot()

    products = inner_root.findall(".//product")

    for p in products:
        name_el = p.find("productName")
        if name_el is None:
            continue

        name = name_el.text.strip()

        if name in mapping