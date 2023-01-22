from pynput.keyboard import Controller, Listener, Key
import docx
import sys


class TextAreaNotFocusedException(Exception):
    pass


class TypeWriter:

    status = 'run'

    def __init__(self) -> None:
        self.keyboard = Controller()
        # in a non-blocking fashion:
        listener = Listener(on_press=self.on_press)
        listener.start()

    def on_press(self, key):
        if key == Key.esc:
            self.status = 'stop'
            sys.exit()

    def start_typing(self, text):
        for char in text:
            if self.status == 'stop':
                break
            elif self.status == 'run':
                self.keyboard.type(char)
                

def read_document(filepath):
    """It will return the text from docx file"""
    file = open(filepath, mode='rb')
    doc = docx.Document(file)
    text = ''
    for para in doc.paragraphs:
        text += para.text + '\n'

    return text
