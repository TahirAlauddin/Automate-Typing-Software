import pyautogui
import docx


class TextAreaNotFocusedException(Exception):
    pass


class TypeWriter:

    status = 'run'

    def __init__(self) -> None:
        pass


    def start_typing(self, text):
        pyautogui.write(text)
                

def read_document(filepath):
    """It will return the text from docx file"""
    doc = docx.Document(filepath)
    text = ''
    for para in doc.paragraphs:
        text += para.text + '\n'

    return text
