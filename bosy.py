import time
import win32gui
import win32ui
import win32con
import numpy as np
import cv2
import requests
import lxml.html
from lxml import html
from datetime import datetime
from datetime import timedelta
from pytesseract import *
pytesseract.tesseract_cmd = r'pathto\tesseract.exe'
import difflib
import math
import os
os.environ['OMP_THREAD_LIMIT'] = '1'
from tkinter import *
from collections import Counter

root = Tk()
root.resizable(False, False)

class WindowCapture:
    w=0
    h=0
    hwnd = None
    def __init__(self, window_name):
        self.hwnd = win32gui.FindWindow(None, window_name)
        window_rect=win32gui.GetWindowRect(self.hwnd)
        self.w=window_rect[2]-window_rect[0]
        self.h=window_rect[3]-window_rect[1]
        border=8
        titlebar=34
        skillbar=68
        self.w=self.w-(border*2)
        self.h=self.h-border-titlebar-skillbar
        self.cropped_x=border
        self.cropped_y=titlebar
        self.w=self.w
        self.h=self.h
        self.cropped_x=border
        self.cropped_y=titlebar
                
    def get_screenshot(self):
        wDC = win32gui.GetWindowDC(self.hwnd)
        dcObj=win32ui.CreateDCFromHandle(wDC)
        cDC=dcObj.CreateCompatibleDC()
        dataBitMap = win32ui.CreateBitmap()
        dataBitMap.CreateCompatibleBitmap(dcObj, self.w, self.h)
        cDC.SelectObject(dataBitMap)
        cDC.BitBlt((0,0),(self.w, self.h) , dcObj, (self.cropped_x,self_cropped_y), win32con.SRCCOPY)
        
        signedIntsArray = dataBitMap.GetBitmapBits(True)
        img = np.frombuffer(signedIntsArray, dtype='uint8')
        img.shape = (self.h,self.w,4)

        dcObj.DeleteDC()
        cDC.DeleteDC()
        win32gui.ReleaseDC(self.hwnd, wDC)
        win32gui.DeleteObject(dataBitMap.GetHandle())
        img = np.ascontiguousarray(img)
        return img

wincap = WindowCapture('nazwaSerwera')

channels = [1,2,3,4,5,6,7,8]
channele = []

bossName = [
    "nazwaBossa1",
    "nazwaBossa2",
    "nazwaBossa3",
    "nazwaBossa4",
    "nazwaBossa5",
    "nazwaBossa6",
    "nazwaBossa7",
    "nazwaBossa8",
    "nazwaBossa9",
    "nazwaBossa10",
    "nazwaBossa11",
    "nazwaBossa12",
    "nazwaBossa13",
    "nazwaBossa14",
    "nazwaBossa15"
]

bossTable = [
    ["nazwaBossa1", 0, 10, 3],
    ["nazwaBossa2", 0, 8, 2],
    ["nazwaBossa3", 0, 3, 2],
    ["nazwaBossa4", 0, 13, 2],
    ["nazwaBossa5", 0, 2, 2],
    ["nazwaBossa6", 0, 9, 2],
    ["nazwaBossa7", 0, 6, 2],
    ["nazwaBossa8", 0, 5, 2],
    ["nazwaBossa9", 0, 4, 2],
    ["nazwaBossa10", 0, 1, 2],
    ["nazwaBossa11", 0, 12, 4],
    ["nazwaBossa12", 0, 11, 2],
    ["nazwaBossa13", 1, 0, 1],
    ["nazwaBossa14", 1, 0, 1],
    ["nazwaBossa15", 1, 0, 1]
]

response = requests.get('http://forum.nazwaSerwera.pl')
tree = lxml.html.fromstring(response.text)
title_elem = tree.xpath('//*xpath do godziny na forum')[0]
spawn_time = title_elem.text_content()
cut_string = spawn_time.split(':')
reset_time = "22."+cut_string[1].strip()+":"+cut_string[2]+":00"
reset_time = datetime.strptime(reset_time, '%y.%d.%m %H:%M:%S')
tera = datetime.now()
difference = tera - reset_time 
diff = (divmod(difference.days * 86400 + difference.seconds, 60))
godzin_round = math.floor(diff[0]/60)

bossyPreview = []

def tableReset(i):
    for x in range(len(bossTable)):
        if bossTable[x][1] == i:
            channele = []
            for y in range(int(bossTable[x][3])):
                if bossTable[x][2] != 0:
                    channele += channels
                    channele.sort()
            bossTable[x] = [bossTable[x][0], bossTable[x][1], bossTable[x][2], bossTable[x][3], channele]
            if godzin_round % 4 == 3:
                if bossTable[x][2] != 0 and bossTable[x][1] != 5:
                    boss = [bossTable[x][0], 1,2,3,4,5,6,7,8, bossTable[x][0]]
                    bossyPreview.append(boss)
            else:
                if bossTable[x][2] != 0 and bossTable[x][1] != 3 and bossTable[x][1] != 5:
                    boss = [bossTable[x][0], 1,2,3,4,5,6,7,8, bossTable[x][0]]
                    bossyPreview.append(boss)

for vx in range(0, 7):
    tableReset(vx)

def first_list_bossy():
    for label in root.winfo_children():
        if type(label) == Label:
            label.destroy()
    for index in range(len(bossyPreview)):
        x = bossyPreview[index]
        num = 0
        for y in x:
            if y in bossName:
                background_clr = "grey"
                width_w = 20
            elif (y == 1 or y == 2 or y == 3 or y == 4 or y == 5 or y == 6 or y == 7 or y == 8):
                background_clr = "green"
                width_w = 4
            else:
                background_clr = "red"
                width_w = 4
            lookup_label = Label(text=y, font='bold', width=width_w, pady=1,
                                 borderwidth=1, relief="solid", background=background_clr)
            lookup_label.grid_propagate(False)
            lookup_label.grid(row=index, column=num)
            num += 1
    root.update()

first_list_bossy()

def list_bossy():
    nume = 0
    for index in range(len(bossyPreview)):
        x = bossyPreview[index]
        for y in x:
            if y in bossName:
                background_clr = "grey"
                width_w = 20
            elif (y == 1 or y == 2 or y == 3 or y == 4 or y == 5 or y == 6 or y == 7 or y == 8):
                background_clr = "green"
                width_w = 4
            else:
                background_clr = "red"
                width_w = 4
            if type(root.winfo_children()[nume]) == Label:
                root.winfo_children()[nume].config(text=y, background=background_clr)
            nume += 1
    root.update()
       
def deep_index(lst, w):
    return [i for (i, sub) in enumerate(lst) if w in sub]

prev_separated_string = []
new_separated_string = []

while True:
    img = wincap.get_screenshot()
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    (thresh, img) = cv2.threshold(img, 127, 255, cv2.THRESH_BINARY)
    img = cv2.bitwise_not(img)
    img = cv2.resize(img, None, fx=1.2, fy=1.2, interpolation=cv2.INTER_AREA)
    output = pytesseract.image_to_string(img, config='--psm 4').strip()
    separated_string = list(filter(bool, output.splitlines()))
    
    if len(separated_string) != 0:
        if prev_separated_string == []:
            for x in range(len(separated_string)):
                a = separated_string[x].split('-')
                b = separated_string[x].split(']')
                if len(a) > 1 and len(b) > 1:
                    channel = b[0][-1]
                    foundBoss = difflib.get_close_matches(a[1], bossName, 1, 0.25)
                    if foundBoss != []:
                        if channel.isnumeric() == True or channel == "S":
                            if channel == "S":
                                channel = "5"
                        else:
                            channel = "u"
                        bossTableIndex = bossName.index(foundBoss[0])
                        prev_separated_string.append([channel, bossTable[bossTableIndex][0]])
                        if channel.isnumeric() == True:
                            channel = int(channel)
                            if channel in bossTable[bossName.index(foundBoss[0])][4]:
                                bossTable[bossTableIndex][4] = (
                                    bossTable[bossTableIndex][4][:bossTable[bossTableIndex][4].index(int(channel))] +
                                    bossTable[bossTableIndex][4][bossTable[bossTableIndex][4].index(int(channel)) + 1:]
                                )
                                if bossTable[bossTableIndex][2] != 0:
                                    if bossTable[bossName.index(bossTable[bossTableIndex][0])][4].count(channel) == 0:
                                        bossNazwa = bossTable[bossTableIndex][0]
                                        listinlist = deep_index(bossyPreview, bossNazwa)
                                        if listinlist[0] != []:
                                            bossyPreview[listinlist[0]][channel] = ''
                                            if (bossyPreview[listinlist[0]][1] != 1 and
                                                bossyPreview[listinlist[0]][2] != 2 and
                                                bossyPreview[listinlist[0]][3] != 3 and
                                                bossyPreview[listinlist[0]][4] != 4 and
                                                bossyPreview[listinlist[0]][5] != 5 and
                                                bossyPreview[listinlist[0]][6] != 6 and 
                                                bossyPreview[listinlist[0]][7] != 7 and
                                                bossyPreview[listinlist[0]][8] != 8):
                                                del bossyPreview[listinlist[0]]
                                                first_list_bossy()
                                                break
                                            else:
                                                list_bossy()
                                                break
        if prev_separated_string != []:
            new_separated_string = []
            for x in range(len(separated_string)):
                a = separated_string[x].split('-')
                b = separated_string[x].split(']')
                if len(a) > 1 and len(b) > 1:
                    channel = b[0][-1]
                    foundBoss = difflib.get_close_matches(a[1], bossName, 1, 0.25)
                    if foundBoss != []:
                        if channel.isnumeric() == True or channel == "S":
                            if channel == "S":
                                channel = "5"
                        else:
                            channel = "u"
                        bossTableIndex = bossName.index(foundBoss[0])
                        new_separated_string.append([channel, bossTable[bossTableIndex][0]])

        if prev_separated_string != new_separated_string:
            indexDiff = []
            for n in range(len(new_separated_string)):
                for p in range(n, len(prev_separated_string)):
                    if (new_separated_string[n][1] == prev_separated_string[p][1] and
                       (new_separated_string[n][0] == prev_separated_string[p][0] or
                        new_separated_string[n][0] == "u" or prev_separated_string[p][0] == "u")):
                        if (new_separated_string[n][0] == "u" and prev_separated_string[p][0] != "u"):
                            prev_separated_string[p][0] = new_separated_string[n][0]
                        if (new_separated_string[n][0] != "u" and prev_separated_string[p][0] == "u"):
                            new_separated_string[n][0] = prev_separated_string[p][0]
                        indexDiff.append(n - p)
                        break
            if indexDiff != []:
                last = len(indexDiff) - 1 - indexDiff[::-1].index(Counter(indexDiff).most_common(1)[0][0])
                for nss in range(len(new_separated_string[last+1:])):
                    channel = new_separated_string[last+1:][nss][0]
                    bossNazwa = new_separated_string[last+1:][nss][1]
                    bossTableIndex2 = bossName.index(bossNazwa)
                    if channel.isnumeric() == True:
                        channel = int(channel)
                        if channel in bossTable[bossName.index(bossNazwa)][4]:
                            bossTable[bossTableIndex2][4] = (
                                bossTable[bossTableIndex2][4][:bossTable[bossTableIndex2][4].index(int(channel))] +
                                bossTable[bossTableIndex2][4][bossTable[bossTableIndex2][4].index(int(channel)) + 1:]
                            )
                            if bossTable[bossTableIndex2][2] != 0:
                                if bossTable[bossName.index(bossNazwa)][4].count(channel) == 0:
                                    listinlist = deep_index(bossyPreview, bossNazwa)
                                    if listinlist != []:
                                        bossyPreview[listinlist[0]][channel] = ''
                                        if (bossyPreview[listinlist[0]][1] != 1 and
                                            bossyPreview[listinlist[0]][2] != 2 and
                                            bossyPreview[listinlist[0]][3] != 3 and
                                            bossyPreview[listinlist[0]][4] != 4 and
                                            bossyPreview[listinlist[0]][5] != 5 and
                                            bossyPreview[listinlist[0]][6] != 6 and 
                                            bossyPreview[listinlist[0]][7] != 7 and
                                            bossyPreview[listinlist[0]][8] != 8):
                                            del bossyPreview[listinlist[0]]
                                            first_list_bossy()
                                            break
                                        else:
                                            list_bossy()
                                            break
            prev_separated_string = new_separated_string
