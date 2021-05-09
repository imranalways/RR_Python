import requests
import io
import pytesseract
# import yaml
from PIL import Image
import re
import ConfigurationManager


def ReadText():
    # Reading text file:
    CIF_ID = "Empty"
    FilePath = ConfigurationManager.RobotData("ImageFilePath")
    img = Image.open(FilePath.strip() + "SQDE_code.PNG")
    width, height = img.size
    new_size = width * 6, height * 6
    img = img.resize(new_size, Image.LANCZOS)
    img = img.convert('L')
    img = img.point(lambda x: 0 if x < 155 else 255, '1')
    print(ConfigurationManager.RobotData("tesseract_path"))
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
    imagetext = pytesseract.image_to_string(img)
    line = imagetext.split()
    for word in line:
        match = re.search(r"(?:(?<!\d)\d{8}(?!\d))", word)
        if match:
            CIF_ID = str(match.group())
    return CIF_ID
