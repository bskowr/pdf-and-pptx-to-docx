import pathlib
from PIL import Image
import pytesseract
import docx


def count_files(main_path):
    count = 0
    for path in pathlib.Path(main_path).iterdir():
        if path.is_file():
            count += 1
    return count


def get_text_from_image(path):
    photo = Image.open(path)
    text_on_photo = pytesseract.image_to_string(photo, lang="pol")
    paragraph = word_output.add_paragraph()
    paragraph.add_run(f"\n***DANE NA SLAJDZIE ZCZYTANE PRZEZ PROGRAM***").bold = True
    paragraph.add_run(f"\n***POCZATEK TEKSTU NA SLAJDZIE***\n{text_on_photo}\n***KONIEC TEKSTU NA SLAJDZIE***\n")


def main(name):
    for i in range(1, count_files(f"UNZIPPED/{name}/") - 1):
        print(f"Przerabiam slajd {i}")
        word_output.add_heading(f"Slajd {i}", level=1)
        get_text_from_image(f"UNZIPPED/{name}/{name}{i:03}.png")
    word_output.save(f"{przedmiot}_output.docx")


if __name__ == '__main__':
    word_output = docx.Document()
    przedmiot = input("Podaj przedmiot do przerobienia: ")
    main(przedmiot)
