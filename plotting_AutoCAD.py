import win32com.client
import pythoncom
import tkinter.filedialog
import pyautocad
import re
import os
import pyautogui as pgui
import time


def checking_autocad():
    """ (None) -> object

    Checking Autocad status and creating it if not running. Running or not. Returning object for working

    """
    acad = win32com.client.Dispatch("AutoCAD.Application")
    # AutoCAD.Application.19 is ProgID
    acad.Visible = True
    return acad


def select_poli(win):
    """ (object) -> list
    Gives to choose polilines and name from curent file in AutoCAD and returns list of Object for working with them
    """
    acad = pyautocad.Autocad(create_if_not_exists=True)
    win.minimize()
    win.maximize()
    time.sleep(1)
    pgui.alert("Choose polilines, which are equivalent to size of paper inside which should be drawing "
               "and text with page number and name of drawing. After choosing press Enter")
    time.sleep(2)
    sel = acad.get_selection()
    time.sleep(1)
    return sel


def select_allPoli(doc):
    """ (object) -> list
    Selecting all elements for printing. Return list of element.
    """
    selection = doc.ActiveSelectionSet
    selection.Clear()
    filter_type = [0, 370]
    filter_type = vtInt(filter_type)
    filter_data = ["lwpolyline", 25]  # lineweilght *100 dxf code 370
    filter_data = vtVariant(filter_data)
    selection.Select(5, 0, 0, filter_type, filter_data)
    filter_type = [0, 8, 67]
    filter_type = vtInt(filter_type)
    filter_data = ["Text,MText", 'dimentions', 1]
    filter_data = vtVariant(filter_data)
    selection.Select(5, 0, 0, filter_type, filter_data)
    return selection


def printing(x1, y1, x2, y2, file_name, adraw):
    """ (float, float, float, float, str) -> None
    Printing drawing and saving it with file_name
    """
    papsize = {"ISO_full_bleed_A0_(841.00_x_1189.00_mm)": ['841x1189', '1189x841'],
               "ISO_full_bleed_A1_(594.00_x_841.00_mm)": ['594x841', '841x594'],
               "ISO_full_bleed_A2_(420.00_x_594.00_mm)": ['594x420', '420x594'],
               "ISO_full_bleed_A3_(297.00_x_420.00_mm)": ['297x420', '420x297'],
               "ISO_full_bleed_A4_(210.00_x_297.00_mm)": ['210x297', '297x210']}
    # Calculating papsize
    a_lay = int(abs(x2 - x1))
    b_lay = int(abs(y2 - y1))
    pap = str(a_lay) + "x" + str(b_lay)
    papnam = "ISO_full_bleed_A3_(297.00_x_420.00_mm)"
    for nam in papsize.keys():
        if pap in papsize[nam]:
            papnam = nam
    # setting plot parameter and printing
    pl = adraw.plot
    ms = adraw.ActiveLayout
    lower_point = vtpnt(x1, y1)
    upper_point = vtpnt(x2, y2)
    ms.ConfigName = "Dwg To PDF.pc3"  # plot name
    time.sleep(2)
    ms.SetWindowToPlot(lower_point, upper_point)  # define the portion of ploting
    ms.PlotType = 4  # acWindow  # setting types of portion of plotting
    ms.CenterPlot = True  # центрировать
    ms.StyleSheet = "monochrome.ctb"  # стиль печати
    ms.StandardScale = 0  # acScaleToFit  # вписать
    ms.CanonicalMediaName = papnam  # розмір листа
    if a_lay > b_lay:
        ms.PlotRotation = 1  # ac90degrees  # ac0degrees - книжная, ac90degrees - альбомная
    elif a_lay < b_lay:
        ms.PlotRotation = 0  # ac0degrees
    pl.PlotToFile(file_name)


def getting_window_cor(select, doc):
    """ (lst) -> lst
    Return list of turple with window coordinates
    """
    select_pol = {}
    for item in select:
        if item.ObjectName == "AcDbPolyline":
            owner_id = str(item.OwnerID)
            if owner_id not in select_pol.keys():
                select_pol[owner_id] = [item.coordinates]
            else:
                select_pol[owner_id].append(item.coordinates)
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
            if name.startswith('Розміри поз.'):
                pass
            else:
                select_text[name] = item.InsertionPoint
    return select_text


def in_poli(x1, y1, x2, y2):
    """ (float, float, float, float, int, int) -> bool

    Return True if text_point in papsize poliline.

    """
    max_x = max(int(x1), int(x2)) - 5
    # min_x = min(int(x1), int(x2))
    min_x = max_x - 185
    # max_y = max(int(y1), int(y2))
    min_y = min(int(y1), int(y2)) - 5
    max_y = min_y + 55
    in_papsize = [min_x, max_x, min_y, max_y]
    return in_papsize


def object_selection_text(select_text, x1, y1, x2, y2):
    """(dict) ==> str
    Select from Autocad text with name of paper and his number for creating a file name. Return file name
    """
    pattern = r'(.*)([\{]?[\\]*[A-Z]{1}[\d]?[\.]?[\d]+;)(.*)([\}])'
    numlst = ''
    namlst = ''
    point = in_poli(x1, y1, x2, y2)
    for text in select_text.keys():
        text_point_x = int(select_text[text][0])
        text_point_y = int(select_text[text][1])
        if text_point_x in range(point[0], point[1]) and text_point_y in range(point[2], point[3]):
            print(text)
            if text.isdigit() or len(text) == 2:
                numlst = str(text)
            else:
                """namelst = re.sub(r'\{?\\[^%s][^;]+;', '', str(text))
                namelst = re.sub(r'\}', '', s)
                for char in namelst:
                    if char not in '\/:*?"<>|':
                        namlst += char"""
                if re.match(pattern, text) != None:
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


def vtpnt(x, y=0):
    """Coordinate points are converted to floating point numbers"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))


def vtFloat(list):
    """The list is converted to a floating point number"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list)


def vtInt(list):
    """list is converted to integer """
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, list)


def vtVariant(list):
    """The list is converted to a variant """
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, list)


def layouts_dic(adraw):
    """ (object) -> dic
    Checking all layouts at drawing and returns dic with layouts objectID and his name
    """
    layouts = adraw.Layouts
    layouts_dic = {}
    for lay in layouts:
        name = lay.Name
        blockl = adraw.Layouts.Item(name).Block
        id = str(blockl.ObjectID)
        layouts_dic[id] = name
    return layouts_dic


def start_plot(acad, select_schlyach):
    """ (object, int) -> None
    Start plotting.
    """
    adraw = acad.ActiveDocument
    if select_schlyach == 1:
        select = select_allPoli(adraw)
    else:
        select = select_poli(win)
    time.sleep(2)
    select_window = getting_window_cor(select, adraw)
    print(select_window)
    select_text = getting_text_dic(select)
    print(select_text)
    layout = layouts_dic(adraw)
    print(layout)
    for layout_id in layout.keys():
        if layout_id in select_window.keys():
            adraw.ActiveLayout = adraw.Layouts.Item(layout[layout_id])
            time.sleep(1)
            for cor in select_window[layout_id]:
                x1 = cor[0]
                y1 = cor[1]
                x2 = cor[2]
                y2 = cor[5]
                name = object_selection_text(select_text, x1, y1, x2, y2)
                file_name = savedirectory + "/" + name + '.pdf'
                print(file_name)
                time.sleep(1)
                printing(x1, y1, x2, y2, file_name, adraw)


acad = checking_autocad()
savedirectory = tkinter.filedialog.askdirectory(title="Папка для збереження креслень")
print(savedirectory)
version = pgui.prompt(text="Версія AutoCAD. Autodesk Autocad 20XX",
                      title="Program version", default='18')
ver = 'Autodesk AutoCAD 20' + str(version)
win = pgui.getWindowsWithTitle(ver)[0]
shlyach = int(pgui.prompt(text="Для друку з вже відкритого файлу введіть 1; Для друку всіх dwg-файлів "
                               "в конкретній папці введіть 2 ", title="Ввід даних", default='1'))
select_schlyach = int(pgui.prompt(text="Для друку всіх аркушів на листах введіть 1; для друку лише вибраних вами "
                                       "листів введіть 2", title='Ввід даних', default='1'))
if shlyach == 2:
    directory = tkinter.filedialog.askdirectory(title="Виберіть папку з файлами для друку")
    files = extract_selected_format(directory)
    for file in files:
        open_file = os.system('"' + file + '"')
        time.sleep(5)
        process = start_plot(acad, select_schlyach)
        time.sleep(2)
        closing = acad.ActiveDocument.Close()
elif shlyach == 1:
    process = start_plot(acad, select_schlyach)
time.sleep(2)
pgui.alert('Job finihed')
