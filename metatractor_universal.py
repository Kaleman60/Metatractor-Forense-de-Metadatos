import os 
import json
from PIL import Image
from PIL.ExifTags import TAGS
from PyPDF2 import PdfReader
from docx import Document 
from openpyxl import load_workbook


def get_image_meta(path):
    meta = {}
    try: 
        with Image.open(path) as imag:
            exif = img._getexif()
            if exif:
                for tag, value in exif.items():
                    decoded = TAGS.get(tag, tag)
                    meta[decoded] = str(value)


    except : pass
    return meta


def get_pdf_meta(path):
    meta = {}
    try: 
        reader = PdfReader(path)
        info = reader.metadata
        for key in info: 
            meta [key.replace('/','')] = str(info[key])

    except: pass
    return meta



def get_office_meta(path):
    meta = {}
    try: 
        if path.endswith('.docx'):
            doc = Document(path)
            prop = doc.core_properties

        elif path.endswith('.xlsx'):
            wb = load_workbook(path)
            prop = wb.properties

        attrs = ['author', 'created', 'last_modified_by','last_printed','title','version']
        for attr in attrs:
            if hasattr(prop, attr):
                meta[attr] = str(getattr(prop, attr))

    except: pass
    return meta


def main():
    