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
            exif = img._get