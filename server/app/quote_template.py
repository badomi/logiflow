"""C트랙 출력 전용: 원본 XLSX의 도형·로고·인쇄 설정을 보존하며 셀 값만 채운다."""

import io
from xml.dom import minidom
from zipfile import ZipFile


def save_filled_template(workbook, template) -> bytes:
    # openpyxl은 일반 도형을 저장하지 못한다. 셀/스타일만 가져오고 나머지는
    # 원본 ZIP의 구성 요소를 그대로 복사한다. 원본 파일에는 쓰지 않는다.
    filled = io.BytesIO()
    workbook.save(filled)
    output = io.BytesIO()
    with ZipFile(template) as original, ZipFile(filled) as changed, ZipFile(output, "w") as result:
        sheet = minidom.parseString(original.read("xl/worksheets/sheet1.xml"))
        cells = minidom.parseString(changed.read("xl/worksheets/sheet1.xml"))
        old = sheet.getElementsByTagName("sheetData")[0]
        sheet.documentElement.replaceChild(sheet.importNode(cells.getElementsByTagName("sheetData")[0], True), old)
        # F열까지 포함한 원래 인쇄 영역을 한 페이지에 맞춘다.
        props = sheet.getElementsByTagName("sheetPr")
        if not props:
            node = sheet.createElement("sheetPr")
            sheet.documentElement.insertBefore(node, sheet.documentElement.firstChild)
            props = [node]
        if props:
            setup = props[0].getElementsByTagName("pageSetUpPr")
            if not setup:
                node = sheet.createElement("pageSetUpPr")
                props[0].appendChild(node)
                setup = [node]
            setup[0].setAttribute("fitToPage", "1")
        for setup in sheet.getElementsByTagName("pageSetup"):
            setup.setAttribute("fitToWidth", "1")
            setup.setAttribute("fitToHeight", "1")
        replacements = {
            "xl/worksheets/sheet1.xml": sheet.toxml(encoding="UTF-8"),
            "xl/styles.xml": changed.read("xl/styles.xml"),
        }
        # TODAY()를 발행일 값으로 바꾸었으므로 원본 수식 계산 체인도 제거한다.
        for path, tag, attr, value in (
            ("xl/_rels/workbook.xml.rels", "Relationship", "Type", "/calcChain"),
            ("[Content_Types].xml", "Override", "PartName", "/xl/calcChain.xml"),
        ):
            doc = minidom.parseString(original.read(path))
            for node in list(doc.getElementsByTagName(tag)):
                if node.getAttribute(attr).endswith(value):
                    node.parentNode.removeChild(node)
            replacements[path] = doc.toxml(encoding="UTF-8")
        for entry in original.infolist():
            if entry.filename != "xl/calcChain.xml":
                result.writestr(entry, replacements.get(entry.filename, original.read(entry.filename)))
    return output.getvalue()
