from pdf2image import convert_from_path, convert_from_bytes
from PIL import Image
import fitz
import os

import pyttsx3
import pytesseract
from googletrans import Translator

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files (x86)/Tesseract-OCR/tesseract.exe'


def get_path_list_from_folder(folder_path):
    path_list = []
    for path in os.listdir(folder_path):
        path_list.append(
            "{}\\{}\\{}".format(
                os.path.dirname(os.path.realpath(__file__)),
                folder_path,
                path
            )
        )
    return path_list


def get_path_list_from_complete_folder_path(folder_path):
    path_list = []
    for path in os.listdir(folder_path):
        path_list.append(
            "{}\\{}".format(
                folder_path,
                path
            )
        )
    return path_list


def make_dir(folder_path):
    try:
        os.mkdir(folder_path)
    except OSError:
        print("Creation of the directory %s failed" % folder_path)
    else:
        print("Successfully created the directory %s " % folder_path)


def pdf_to_images(pdf_path, image_output_folder):
    doc_name = pdf_path.split("\\")[-1].split("/")[-1]

    doc = fitz.open(pdf_path)
    page_number = 0
    while True:
        try:
            page = doc.loadPage(page_number)  # number of page
            pix = page.getPixmap()

            file_name = "{} - {}.png".format(
                doc_name,
                (3 - len(str(page_number))) * "0" + str(page_number)
            )

            output = "{}\\{}".format(image_output_folder, file_name)
            pix.writePNG(output)
            page_number += 1
        except ValueError:
            break


def images_to_txt_doc(txt_doc_name, file_name, image_output_folder):

    with open(txt_doc_name, mode='w') as file:

        for i, image_path in enumerate(get_path_list_from_complete_folder_path(image_output_folder)):
            print("\t - {}".format(image_path))

            if i != 0:
                file.write("\n\n-------------------------------------------\n\n")

            img = Image.open(image_path)
            result = pytesseract.image_to_string(img)

            file.write("{} - Page {}\n".format(file_name, i))
            file.write(result)


def main():

    base_image_output_folder = "PDF to Images"

    print("Converting PDFs to Images")
    for pdf_path in get_path_list_from_folder("PDFs to Load"):
        print("Loading - {}".format(pdf_path))

        file_name = pdf_path.split("\\")[-1].split("/")[-1]
        new_image_output_folder = "{}\\{}".format(base_image_output_folder, "{} Folder".format(file_name))
        make_dir(new_image_output_folder)

        pdf_to_images(pdf_path, new_image_output_folder)

    base_txt_output_folder = "Images to Txt"

    print("Converting Images to Txt")
    for image_output_folder in get_path_list_from_folder(base_image_output_folder):
        print("Loading - {}".format(image_output_folder))

        txt_file_name = ".".join(image_output_folder.split("\\")[-1].split("/")[-1].split(".")[:-1])
        txt_file_path = "{}\\{}.txt".format(
            base_txt_output_folder,
            txt_file_name,
            ".txt"
        )

        images_to_txt_doc(txt_file_path, txt_file_name, image_output_folder)


main()
