import xml.etree.ElementTree as xmltree
from PIL import Image
import pytesseract
import docx
from docx.shared import Cm
import pathlib

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# output = open('output.txt', 'w', encoding='utf-8')
word_output = docx.Document()


def get_text_from_photo(media_path, photo_name, paragraph):
    photo = Image.open(media_path + photo_name)
    text_on_photo = pytesseract.image_to_string(photo, lang="pol")
    # output.write(f"\n***DANE NA OBRAZKU {photo_name} ZCZYTANE PRZEZ PROGRAM***")
    # output.write(f"\n***POCZATEK TEKSTU NA OBRAZKU***\n{text_on_photo}\n***KONIEC TEKSTU NA OBRAZKU***\n")
    paragraph.add_run(f"\n***DANE NA OBRAZKU {photo_name} ZCZYTANE PRZEZ PROGRAM***").bold = True
    paragraph.add_run(f"\n***POCZATEK TEKSTU NA OBRAZKU***\n{text_on_photo}\n***KONIEC TEKSTU NA OBRAZKU***\n")
    word_output.add_heading(f"OBRAZEK {photo_name}", level=2)
    word_output.add_picture(media_path + photo_name, width=Cm(10))


def find_photos(file_path, relation_path, media_path):
    file = xmltree.parse(file_path)
    relations = xmltree.parse(relation_path)
    file_root = file.getroot()
    rel_root = relations.getroot()
    for blip in file_root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"):
        photo_id = blip.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed']
        for child in rel_root:
            if photo_id == child.attrib["Id"]:
                paragraph = word_output.add_paragraph()
                photo_name = child.attrib["Target"].split("/")[2]
                get_text_from_photo(media_path, photo_name, paragraph)


def find_text_in_xml(file_path):
    file = xmltree.parse(file_path)
    root = file.getroot()
    paragraph_iter = root.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}p")
    for p in paragraph_iter:
        paragraph = word_output.add_paragraph()
        for r in p.findall("{http://schemas.openxmlformats.org/drawingml/2006/main}r"):
            parse_text(r, paragraph)
        # output.write("\n")


def parse_text(text_data, paragraph):
    rPr = text_data.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
    t = text_data.find("{http://schemas.openxmlformats.org/drawingml/2006/main}t")
    if "b" in rPr.attrib.keys():
        if rPr.attrib["b"]:
            if t.text is not None:
                paragraph.add_run(t.text.strip() + " ").bold = True
                # output.write(t.text.strip() + "\n")
    else:
        if t.text is not None:
            paragraph.add_run(t.text.strip() + " ")
        # output.write(t.text.strip() + " ")


def count_slides(main_path):
    count = 0
    for path in pathlib.Path(main_path).iterdir():
        if path.is_file():
            count += 1
    return count


def main(main_path, rel_path, media_path):
    for i in range(1, count_slides(main_path) + 1):
        print(f"Przerabiam slajd {i}")
        # output.write(f"Slajd {i}\n")
        word_output.add_heading(f"Slajd {i}", level=1)
        find_text_in_xml(f"{main_path}slide{i}.xml")
        find_photos(f"{main_path}slide{i}.xml", f"{rel_path}slide{i}.xml.rels", media_path)
        # output.write("\n")
    word_output.save(f"{przedmiot}.docx")


przedmiot = input("Podaj przedmiot do przerobienia: ")
sciezka_media = f"UNZIPPED/{przedmiot}/ppt/media/"
sciezka_slajdy = f"UNZIPPED/{przedmiot}/ppt/slides/"
sciezka_relacje = f"UNZIPPED/{przedmiot}/ppt/slides/_rels/"
main(sciezka_slajdy, sciezka_relacje, sciezka_media)
