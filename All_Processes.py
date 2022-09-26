class process:
    import pyautogui
    import time
    import os
    import pyperclip
    import keyboard
    import cv2
    import numpy as np
    import win32api, win32con

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots')

    def click(x, y):
        self.win32api.SetCursorPos((x, y))
        self.win32api.mouse_event
        (self.win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        self.time.sleep(0.1)
        (self.win32con.MOUSEEVENTF_LEFTUP, 0,0)

    def search_JobRef(self, jobref):
        FindRef = self.pyautogui.locateOnScreen('FindRef.png', confidence=0.9)
        if FindRef != None:
            FindRef = self.pyautogui.center(FindRef)
            x, y = FindRef
            self.pyautogui.click(x, y)
        DragTO = self.pyautogui.locateOnScreen('FindRef dragleft.png', confidence=0.9)
        if DragTO != None:
            DragTO = self.pyautogui.center(DragTO)
            x, y = DragTO
            self.pyautogui.dragTo(x, y, button='left')
            self.pyautogui.press('backspace')  # deletes what is entered in box
            self.pyautogui.typewrite(str(jobref))
        SearchRef = self.pyautogui.locateOnScreen('SearchRef.png', confidence=0.9)
        if SearchRef != None:
            SearchRef = self.pyautogui.center(SearchRef)
            x, y = SearchRef
            self.pyautogui.click(x, y)
            self.time.sleep(3)

