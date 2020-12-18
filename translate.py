import comtypes.client
from comtypes.client import PumpEvents, ShowEvents

from pyautogui import press, typewrite, hotkey

from time import sleep

word = ''


def init():
    global word
    word = comtypes.client.CreateObject('Word.Application')
    word.Documents.Open('./lectures/lect1.docx')
    word.Visible = 1


if __name__ == "__main__":
    init()
    hotkey('alt', 'tab');sleep(2);press('alt');press('r');sleep(1);press('l');sleep(1);press('t');sleep(5)