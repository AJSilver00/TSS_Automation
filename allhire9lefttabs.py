class allhire9lefttabs:
    import pyautogui
    import time
    import os
    import pyperclip
    import keyboard
    import cv2
    import numpy as np
    import win32api, win32con

    def __init__(self):
        # location of all screenshots
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots')
        # 9 unselected tabs:
        self.summary_unselected_found = False
        self.allhire_unselected_found = False
        self.contacts_unselected_found = False
        self.echo_unselected_found = False
        self.purchaseorders_unselected_found = False
        self.stock_unselected_found = False
        self.eventsvenues_unselected_found = False
        self.finalaccount_unselected_found = False
        self.reports_unselected_found = False
        # same 9 tabs but now selected:
        self.summary_selected_found = False
        self.allhire_selected_found = False
        self.contacts_selected_found = False
        self.echo_selected_found = False
        self.purchaseorders_selected_found = False
        self.stock_selected_found = False
        self.eventsvenues_selected_found = False
        self.finalaccount_selected_found = False
        self.reports_selected_found = False

    def clicks(self, x, y):
        self.process.click(x, y)

# def clicksummarytab(self):
#     self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots\allhire9lefttabs')
#     while summary_unselected_found == False:
#         summaryunsel = self.pyautogui.locateOnScreen('summaryunsel.png', confidence=0.9)
#         if summaryunsel == pyscreeze.Box:
#             summarysel = self.pyautogui.center(summaryunsel)
#             x, y = summarysel
#             process.click(x,y)
#
#             self.summary_unselected_found = False
#             self.summary_selected_found = True
#     while summary_unselected_found == False:
#         summarysel = self.pyautogui.locateOnScreen('summarysel.png', confidence=0.9)
#         elif summarysel == pyscreeze.Box:
#
#             summarysel = self.pyautogui.center(summarysel)
#             x, y = summarysel
#             self.pyautogui.click(x, y)
#             self.summary_unselected_found = False
#             self.summary_selected_found = True


class findhirejob:
    import pyautogui
    import time
    import os
    import pyperclip
    import win32api, win32con
    import random

    def __init__(self):
        self.os.chdir(r'C:\Users\Aimee\Pictures\TSS Python Screenshots')
        # all unselected buttons:
        self.allhiretab_unselected_found = True
        self.searchtab_unselected_found = True
        self.firm_unselected_found = True
        self.enq_unselected_found = True
        self.cross_unselected_found = True
        self.os_unselected_found = True
        self.includearchived_unselected_found = True
        self.searchbar_unselected_found = True
        self.searchbutton_unselected_found = True
        self.listrecentjobs_unselected_found = True
        self.findbynameitem_unselected_found = True
        self.jobsbymonth_unselected_found = True
        self.searchsettings_unselected_found = True
        self.tickall_unselected_found = True
        self.tickname_unselected_found = True
        self.tickaddress_unselected_found = True
        self.tickreference_unselected_found = True
        self.tickdeladdress_unselected_found = True
        self.tickjobfileevent_unselected_found = True
        self.orderbyreference_unselected_found = True
        self.limitresultsto100_unselected_found = True
        # same buttons now selected:
        self.allhiretab_selected_found = True
        self.searchtab_selected_found = True
        self.firm_selected_found = True
        self.enq_selected_found = True
        self.cross_selected_found = True
        self.os_selected_found = True
        self.includearchived_selected_found = True
        self.searchbar_selected_found = True
        self.searchbutton_selected_found = True
        self.listrecentjobs_selected_found = True
        self.findbynameitem_selected_found = True
        self.jobsbymonth_selected_found = True
        self.searchsettings_selected_found = True
        self.tickall_selected_found = True
        self.tickname_selected_found = True
        self.tickaddress_selected_found = True
        self.tickreference_selected_found = True
        self.tickdeladdress_selected_found = True
        self.tickjobfileevent_selected_found = True
        self.orderbyreference_selected_found = True
        self.limitresultsto100_selected_found = True

