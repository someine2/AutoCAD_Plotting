import comtypes
import comtypes.client
import tkinter.filedialog
import pyautocad
import re
import os
import pyautogui as pgui
import time
import array
from comtypes.gen.AutoCAD import *


def checking_autocad():
    """ (None) -> object

    Checking Autocad status. Running or not. Returning comtypes.Pointer (object for working)

    """
    try:
        acad = comtypes.client.GetActiveObject("Autocad.Application")
        acad.Visible = True
    except:
        acad = comtypes.client.CreateObject("AutoCAD.Application")
        acad.Visible = True
    return acad


def select_poli(win):
    """ (object) -> list
    Gives to choose polilines and name from curent file in AutoCAD and returns list of Object for working with them
    """
    acad = pyautocad.Autocad(create_if_not_exists=True)
    win.minimize()
    win.maximize()
    win.minimize()
    time.sleep(1)
    pgui.alert("Виберіть полінії, що відповідає розмірам формату листа та назви і номери листів. "
               "Після вибору натисніть Enter")
    time.sleep(2)
    sel = acad.get_selection()
    time.sleep(5)
    return sel


def printing(x1, y1, x2, y2, file_name):
    """ (float, float, float, float, str) -> None
    Printing drawing and saving it with file_name
    """
    papsize = {"ISO_full_bleed_A0_(841.00_x_1189.00_mm)": ['841x1189', '1189x841'],
               "ISO_full_bleed_A1_(594.00_x_841.00_mm)": ['594x841', '841x594'],
               "ISO_full_bleed_A2_(420.00_x_594.00_mm)": ['594x420', '420x594'],
               "ISO_full_bleed_A3_(297.00_x_420.00_mm)": ['297x420', '420x297'],
               "ISO_full_bleed_A4_(210.00_x_297.00_mm)": ['210x297', '297x210']}
    #Calculating papsize
    a_lay = int(abs(x2 - x1))
    b_lay = int(abs(y2 - y1))
    pap = str(a_lay) + "x" + str(b_lay)
    papnam = "ISO_full_bleed_A3_(297.00_x_420.00_mm)"
    for nam in papsize.keys():
        if pap in papsize[nam]:
            papnam = nam
            print(papnam)
    #setting plot parameter and printing
    pl = adraw.plot
    ms = adraw.ActiveLayout
    lower_point = array.array('d', [x1, y1])
    print(lower_point)
    upper_point = array.array('d', [x2, y2])
    print(upper_point)
    ms.ConfigName = "Dwg To PDF.pc3"  # plot name
    time.sleep(2)
    ms.SetWindowToPlot(lower_point, upper_point)  # define the portion of ploting
    ms.PlotType = acWindow  # setting types of portion of plotting
    ms.CenterPlot = True  # центрировать
    ms.StyleSheet = "monochrome.ctb"  # стиль печати
    ms.StandardScale = acScaleToFit  # вписать
    ms.CanonicalMediaName = papnam  # розмір листа
    if a_lay > b_lay:
        ms.PlotRotation = ac90degrees  # ac0degrees - книжная, ac90degrees - альбомная
    elif a_lay < b_lay:
        ms.PlotRotation = ac0degrees
    pl.PlotToFile(file_name)


def getting_window_cor(select):
    """ (lst) -> lst
    Return list of turple with window coordinates
    """
    select_pol = []
    for item in select:
        if item.ObjectName == "AcDbPolyline":
            select_pol.append(item.coordinates)
    return select_pol


def getting_text_dic(select):
    """ (lst) -> dict
    Return dict with selected text and his Insertion Point
    """
    select_text = {}
    for item in select:
        if item.ObjectName == "AcDbMText" or item.ObjectName == 'AcDbText':
            name = item.TextString
            name = name.strip('\\pxqc;')
            select_text[name] = item.InsertionPoint
    return select_text


def object_selection_text(select_text, x1, y1, x2, y2):
    """(dict) ==> str
    Select from Autocad text with name of paper and his number for creating a file name. Return file name
    """
    pattern = r'(.*)([\{]?[\\]*[A-Z]{1}[\d]?[\.]?[\d]+;)(.*)([\}])'
    numlst = ''
    namlst = ''
    for text in select_text.keys():
        if (min(int(x1), int(x2)) <= int(select_text[text][0]) and int(select_text[text][0]) <= max(int(x1),
                                                                                                    int(x2))) and (
                min(int(y1), int(y2)) <= int(select_text[text][1]) and int(select_text[text][1]) <= max(int(y1),
                                                                                                        int(y2))):
            if text.isdigit() or len(text) == 2:
                numlst = str(text)
            else:
                if re.match(pattern, text) is not None:
                    namelst = re.sub(pattern, r'\3', str(text))
                    # print(namelst)
                else:
                    namelst = str(text)
                for char in namelst:
                    if char not in '\/:*?"<>|':
                        namlst += char
    return numlst + '_' + namlst + '_' + 'rev0'

def extract_selected_format(directory_path):
    """ (str) -> list
    Checking for dwg files in directory and return list of files.
    """
    files = os.listdir(directory_path)
    pat = r'.*.dwg'
    files1 = []
    for file in files:
        if re.fullmatch(pat, file) is not None:
            files1.append(file)
    files = files1
    files = [os.path.join(directory_path, file) for file in files]
    return files

acad = checking_autocad()
savedirectory = tkinter.filedialog.askdirectory(title="Виберіть папку для збереження надрукованих креслень")
adraw = acad.ActiveDocument
version = pgui.prompt(text="Введіть свою версію Autocad(значення замість ХХ). Autodesk Autocad 20XX",
                      title="Версія програми", default='18')
ver = 'Autodesk AutoCAD 20' + str(version)
win = pgui.getWindowsWithTitle(ver)[0]
select = select_poli(win)
select_window = getting_window_cor(select)
select_text = getting_text_dic(select)
for cor in select_window:
    x1 = cor[0]
    y1 = cor[1]
    x2 = cor[2]
    y2 = cor[5]
    name = object_selection_text(select_text, x1, y1, x2, y2)
    file_name = savedirectory + '/' + name + '.pdf'
    c = printing(x1, y1, x2, y2, file_name)
    time.sleep(2)
pgui.alert('Job finihed')