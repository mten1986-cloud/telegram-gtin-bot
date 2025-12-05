def process_xml_with_cdata(xml_text: str) -> str:
    cdata_pattern = r"<!\[CDATA\[(.*)\]\]>"
    match = re.search(cdata_pattern, xml_text, re.DOTALL)

    if not match:
        raise ValueError("В файле нет CDATA со вложенным XML.")

    inner_xml = match.group(1).strip()

    root = ET.fromstring(inner_xml)
    tree = ET.ElementTree(root)

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

    indent(root)

    updated_inner_xml = ET.tostring(root, encoding="unicode")

    # ----------- ПРИНУДИТЕЛЬНО ВОССТАНАВЛИВАЕМ v1: ----------
    updated_inner_xml = updated_inner_xml.replace("ns0:", "v1:")
    updated_inner_xml = updated_inner_xml.replace('xmlns:ns0="v1.snt"', 'xmlns:v1="v1.snt"')

    updated_cdata = f"<![CDATA[\n{updated_inner_xml}\n]]>"
    final_xml = re.sub(cdata_pattern, updated_cdata, xml_text, flags=re.DOTALL)

    return final_xml

