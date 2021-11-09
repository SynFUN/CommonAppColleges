# -*- coding = utf-8 -*-
# encoding: utf-8
# @Time : 2021.11.1 09:19
# @Author : Synthesis 杜品赫
# @File : main.pypip install lxml==3.7.0
# @Software : PyCharm
# https://github.com/SynFUN/Portable_CommonAppColleges

# pip install pywin32==223

import os
import re
import sys
import time
import tkinter
from tkinter import filedialog
from bs4 import BeautifulSoup
from datetime import datetime
import xlwings as xw


def fileChooser(title: str, filetypes: list = None) -> str:
    """
    :param title: 窗口名称 The name of the folder chooser window
    :param filetypes: [('File Type Describe', '*.filename_extension'), ('TXT File', '*.txt'), ('Audio File', '*.mp3')]
    :return: 选取的文件的路径 The selected file's path
    """
    if filetypes is None:
        filetypes = [('All Files', '*')]
    path = tkinter.filedialog.askopenfilename(title=title, filetypes=filetypes)
    return path


def fileSaver(title: str, filetypes: list, defaultextension: str, initialfile: str) -> str:
    """
    :param title: 窗口名称 The name of the folder chooser window
    :param filetypes: [('File Type Describe', '*.filename_extension'), ('TXT File', '*.txt'), ('Audio File', '*.mp3')]
    :param defaultextension: 补全扩展名 Default file extension
    :param initialfile: 补全文件名 Default file name
    :return: 保存的文件的路径 The saved file's path
    """
    path = tkinter.filedialog.asksaveasfilename(title=title, filetypes=filetypes, defaultextension=defaultextension,
                                                initialfile=initialfile)
    return path


if __name__ == '__main__':
    print(r"""                                                                                                       
                                                                                                       
  ,----..                    ____           ____                        ,---,                          
 /   /   \                 ,'  , `.       ,'  , `.                     '  .' \     ,-.----. ,-.----.   
|   :     :  ,---.      ,-+-,.' _ |    ,-+-,.' _ |  ,---.       ,---, /  ;    '.   \    /  \\    /  \  
.   |  ;. / '   ,'\  ,-+-. ;   , || ,-+-. ;   , || '   ,'\  ,-+-. /  :  :       \  |   :    |   :    | 
.   ; /--` /   /   |,--.'|'   |  ||,--.'|'   |  ||/   /   |,--.'|'   :  |   /\   \ |   | .\ |   | .\ : 
;   | ;   .   ; ,. |   |  ,', |  ||   |  ,', |  |.   ; ,. |   |  ,"' |  :  ' ;.   :.   : |: .   : |: | 
|   : |   '   | |: |   | /  | |--'|   | /  | |--''   | |: |   | /  | |  |  ;/  \   |   |  \ |   |  \ : 
.   | '___'   | .; |   : |  | ,   |   : |  | ,   '   | .; |   | |  | '  :  | \  \ ,|   : .  |   : .  | 
'   ; : .'|   :    |   : |  |/    |   : |  |/    |   :    |   | |  |/|  |  '  '--' :     |`-:     |`-' 
'   | '/  :\   \  /|   | |`-'     |   | |`-'      \   \  /|   | |--' |  :  :       :   : :  :   : :    
|   :    /  `----' |   ;/         |   ;/           `----' |   |/     |  | ,'       |   | :  |   | :    
 \   \ .'          '---'          '---'                   '---'      `--''         `---'.|  `---'.|    
  `---`                                                                              `---`    `---`                                                                                                                                             
  ,----..            ,--,    ,--,                                           
 /   /   \         ,--.'|  ,--.'|                                           
|   :     :  ,---. |  | :  |  | :                                           
.   |  ;. / '   ,'\:  : '  :  : '             ,----._,.          .--.--.    
.   ; /--` /   /   |  ' |  |  ' |     ,---.  /   /  ' /  ,---.  /  /    '   
;   | ;   .   ; ,. '  | |  '  | |    /     \|   :     | /     \|  :  /`./   
|   : |   '   | |: |  | :  |  | :   /    /  |   | .\  ./    /  |  :  ;_     
.   | '___'   | .; '  : |__'  : |__.    ' / .   ; ';  .    ' / |\  \    `.  
'   ; : .'|   :    |  | '.'|  | '.''   ;   /'   .   . '   ;   /| `----.   \ 
'   | '/  :\   \  /;  :    ;  :    '   |  / |`---`-'| '   |  / |/  /`--'  / 
|   :    /  `----' |  ,   /|  ,   /|   :    |.'__/\_: |   :    '--'.     /  
 \   \ .'           ---`-'  ---`-'  \   \  / |   :    :\   \  /  `--'---'   
  `---`                              `----'   \   \  /  `----'              
                                               `--`-'                       """)
    print("https://github.com/SynFUN/Portable_CommonAppColleges v1.0 2021.11.8\n")
    print("Initializing...")
    # 初始化Tk
    tkinter.Tk().withdraw()
    try:
        path = ""
        try:
            path += str(sys.argv[1])
        except Exception:
            pass
        if path == "":
            path = fileChooser("Choose Common App Application Requirements Page's HTML", [('HTML', '*.html')])
        with open(path, 'r', encoding='utf-8') as file:
            xlsx = fileSaver('Save Xlsx', [('Xlsx File', '*.xlsx')], ".xlsx", "Common App")
            print("Converting...")
            app = xw.App(visible=False, add_book=False)
            app.display_alerts = False
            app.screen_updating = False
            if not os.path.exists(xlsx):
                excelFile = app.books.add()
                excelFile.save(xlsx)
            wb = app.books.open(xlsx)
            sheet = wb.sheets.active
            html = file.read()
            html = re.sub(
                r'<tr _ngcontent-[a-z][a-z]*[a-z]*[a-z]*-c236="" id="reqRow[0-9][0-9]*[0-9]*[0-9]*" class="ng-star-inserted">',
                r'<tr id="college">', html, count=0, flags=0)
            soup1 = BeautifulSoup(html, 'lxml')
            all_college = []
            for item in soup1.find_all('tr', id='college'):
                soup2 = BeautifulSoup(str(item), 'lxml')
                for tag in soup2.find_all('th'):
                    # Debug的时候才发现下面正则里【[a-z][a-z]*[a-z]*[a-z]*】是会变的
                    s_tag = re.sub('<th _ngcontent-[a-z][a-z]*[a-z]*[a-z]*-c236="" scope="row">', '', str(tag)).replace(
                        '</th>', '')
                    college_name = s_tag
                count = -1
                # 这个for写的很屎 但是懒得改了 别学 建议用回调函数做这种事
                # This for loop is so suck but I don't want to change it anymore
                for tag in soup2.find_all('td'):
                    # Debug的时候才发现下面正则里【[a-z][a-z]*[a-z]*[a-z]*】是会变的
                    # 所以直接用正则暴力转换成第一版代码里的【nwq】了
                    s_tag = str(re.sub('<td _ngcontent-[a-z][a-z]*[a-z]*[a-z]*-c236="">', '<td _ngcontent-nwq-c236="">',
                                       str(tag)))
                    if count == 0:
                        if s_tag != '<td _ngcontent-nwq-c236=""></td>':
                            ed = s_tag.replace('<td _ngcontent-nwq-c236="">', '').replace('</td>', '')
                        else:
                            ed = "-1"
                    if count == 1:
                        if s_tag != '<td _ngcontent-nwq-c236=""></td>':
                            edII = s_tag.replace('<td _ngcontent-nwq-c236="">', '').replace('</td>', '')
                        else:
                            edII = "-1"
                    if count == 2:
                        if s_tag != '<td _ngcontent-nwq-c236=""></td>':
                            ea = s_tag.replace('<td _ngcontent-nwq-c236="">', '').replace('</td>', '')
                        else:
                            ea = "-1"
                    if count == 3:
                        if s_tag != '<td _ngcontent-nwq-c236=""></td>':
                            eaII = s_tag.replace('<td _ngcontent-nwq-c236="">', '').replace('</td>', '')
                        else:
                            eaII = "-1"
                    if count == 4:
                        if s_tag != '<td _ngcontent-nwq-c236=""></td>':
                            rea = s_tag.replace('<td _ngcontent-nwq-c236="">', '').replace('</td>', '')
                        else:
                            rea = "-1"
                    if count == 5:
                        if s_tag != '<td _ngcontent-nwq-c236=""></td>':
                            rd = s_tag.replace('<td _ngcontent-nwq-c236="">', '').replace('</td>', '')
                        else:
                            rd = "-1"
                    if count == 6:
                        break
                    count += 1
                college = [college_name, ed, edII, ea, eaII, rea, rd]
                all_college.append(college)
            width_count = 1
            college_count = 2
            sheet.range('A1').value = "College"
            sheet.range('B1').value = "ED"
            sheet.range('C1').value = "EDII"
            sheet.range('D1').value = "EA"
            sheet.range('E1').value = "EAII"
            sheet.range('F1').value = "REA"
            sheet.range('G1').value = "RD/Rolling"
            for i in all_college:
                if len(str(i[0])) > width_count: width_count = len(str(i[0]))
            for i in all_college:
                prt = str(i[0])
                sheet.range('A' + str(college_count)).value = str(i[0])
                sheet.range('A' + str(college_count)).column_width = int(width_count)
                if str(i[1]) != '-1':
                    prt += "\nED: " + datetime.strptime(str(i[1]), '%m/%d/%Y').strftime('%Y-%m-%d')
                    sheet.range('B' + str(college_count)).value = datetime.strptime(str(i[1]), '%m/%d/%Y').strftime(
                        '%Y-%m-%d')
                if str(i[2]) != '-1':
                    prt += "\nEDII: " + datetime.strptime(str(i[2]), '%m/%d/%Y').strftime('%Y-%m-%d')
                    sheet.range('C' + str(college_count)).value = datetime.strptime(str(i[2]), '%m/%d/%Y').strftime(
                        '%Y-%m-%d')
                if str(i[3]) != '-1':
                    prt += "\nEA: " + datetime.strptime(str(i[3]), '%m/%d/%Y').strftime('%Y-%m-%d')
                    sheet.range('D' + str(college_count)).value = datetime.strptime(str(i[3]), '%m/%d/%Y').strftime(
                        '%Y-%m-%d')
                if str(i[4]) != '-1':
                    prt += "\nEAII: " + datetime.strptime(str(i[4]), '%m/%d/%Y').strftime('%Y-%m-%d')
                    sheet.range('E' + str(college_count)).value = datetime.strptime(str(i[4]), '%m/%d/%Y').strftime(
                        '%Y-%m-%d')
                if str(i[5]) != '-1':
                    prt += "\nREA: " + datetime.strptime(str(i[5]), '%m/%d/%Y').strftime('%Y-%m-%d')
                    sheet.range('F' + str(college_count)).value = datetime.strptime(str(i[5]), '%m/%d/%Y').strftime(
                        '%Y-%m-%d')
                if str(i[6]) != '-1':
                    prt += "\nRD/Rolling: " + datetime.strptime(str(i[6]), '%m/%d/%Y').strftime('%Y-%m-%d')
                    sheet.range('G' + str(college_count)).value = datetime.strptime(str(i[6]), '%m/%d/%Y').strftime(
                        '%Y-%m-%d')
                for x in ['B', 'C', 'D', 'E', 'F', 'G']:
                    sheet.range(x + str(college_count)).column_width = 10
                college_count += 1
                print(prt + '\n')
        wb.save(xlsx)
        wb.close()
        app.quit()
        print("Convert done")
    except FileNotFoundError as te:
        with open(r'.\error.log', 'a', encoding='utf-8') as file:
            file.write(time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '\n' + str(
                te) + '\nOpen file canceled (e001)\n\n')
    except Exception as e:
        with open(r'.\error.log', 'a', encoding='utf-8') as file:
            file.write(
                time.strftime("%Y-%m-%d %H:%M:%S", time.localtime()) + '\n' + str(e) + '\nUnknown error (e002)\n\n')
