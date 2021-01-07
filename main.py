# -*- coding: utf-8 -*-

from docx import Document
from docx.shared import Pt

import os
from os import listdir
from os.path import isfile, join

import ctypes
from ctypes import wintypes

import math
import configparser


def get_font(fnt):
    class LOGFONT(ctypes.Structure):
        _fields_ = [
            ('lfHeight', wintypes.LONG),
            ('lfWidth', wintypes.LONG),
            ('lfEscapement', wintypes.LONG),
            ('lfOrientation', wintypes.LONG),
            ('lfWeight', wintypes.LONG),
            ('lfItalic', wintypes.BYTE),
            ('lfUnderline', wintypes.BYTE),
            ('lfStrikeOut', wintypes.BYTE),
            ('lfCharSet', wintypes.BYTE),
            ('lfOutPrecision', wintypes.BYTE),
            ('lfClipPrecision', wintypes.BYTE),
            ('lfQuality', wintypes.BYTE),
            ('lfPitchAndFamily', wintypes.BYTE),
            ('lfFaceName', ctypes.c_wchar * 32)
        ]

    fontenumproc = ctypes.WINFUNCTYPE(ctypes.c_int,
                                      ctypes.POINTER(LOGFONT), wintypes.LPVOID, wintypes.DWORD, wintypes.LPARAM)

    fontlist = []

    def font_enum(logfont, textmetricex, fonttype, param):
        string = logfont.contents.lfFaceName
        if not any(string in s for s in fontlist):
            fontlist.append(string)
        return True

    hdc = ctypes.windll.user32.GetDC(None)
    ctypes.windll.gdi32.EnumFontFamiliesExW(hdc, None, fontenumproc(font_enum), 0, 0)
    ctypes.windll.user32.ReleaseDC(None, hdc)
    return fnt.lower() in [x.lower() for x in fontlist]


output_font = ""
input_files = []
input_folder_name = ""


def check_errors():
    must_exit = False

    if not os.path.exists("settings.ini"):
        config['INPUT'] = {
            '# Input Folder Name': None,
            'InputFolder': 'inputs',
        }

        config['OUTPUT'] = {
            '# Output Folder Name': None,
            'OutputFolder': 'outputs',

            '\n# Output File\'s Base Name': None,
            'OutputFile': 'output',

            '\n# Output File\'s Font Name': None,
            'Font': 'Cousine',

            '\n# Output File\'s Font Size': None,
            'FontSize': 11,

            '\n# Output File\'s Line Spacing': None,
            'LineSpacing': 1.75,

            '\n# Output File\'s Character Count Per Line': None,
            'LineCharLength': 65,

            '\n# Output File\'s Minimum Underscore Length': None,
            'MinUnderscoreLength': 25,

            '\n# Output Files Use Right-To-Left Formatting': None,
            'RightToLeft': False,
            
            '\n# Character to Remove From Output': None,
            'CharsToRemove': ['.', ',', '/', '\\', "'", '"', '?', '<', '>', '!', '@', '(', ')', '-', '_', '[', ']', ':', ';'],
        }

        print("NO SETTINGS FILE FOUND!")
        print("CREATING ONE! PLEASE VERIFY THE CONFIGURATIONS!")

        with open('settings.ini', 'w') as configfile:
            config.write(configfile)

        must_exit = True

    config.read('settings.ini')

    global input_folder_name
    input_folder_name = str(config['INPUT']['inputfolder']) + '/'

    if not os.path.exists(input_folder_name):
        os.mkdir(input_folder_name)

    global input_files
    input_files = [f for f in listdir(input_folder_name) if isfile(join(input_folder_name, f))]

    if len(input_files) <= 0:
        print("NO INPUT TEXT FILE FOUND!")
        print("Creating one! Please verify the configurations!")

        f = open(str(input_folder_name + "input.txt"), "a")
        f.close()

        must_exit = True

    global output_font
    output_font = str(config['OUTPUT']['font'])

    if not get_font(output_font):
        print("FONT NOT INSTALLED!")
        print("Please install %s or change the font in the settings file!" % output_font)

        must_exit = True

    if must_exit:
        print('Restart the PROGRAM after fixing above issues.')
        print('Press ENTER to exit.')
        input()
        exit()


config = configparser.ConfigParser(allow_no_value=True)

check_errors()

output_folder_name = str(config['OUTPUT']['outputfolder']) + '/'
output_file_name = str(config['OUTPUT']['outputfile'])
output_file = str(output_folder_name + output_file_name)
output_fontsize = float(config['OUTPUT']['fontsize'])
output_linespacing = float(config['OUTPUT']['linespacing'])
output_linecharlength = int(config['OUTPUT']['linecharlength'])
output_minunderscorelength = int(config['OUTPUT']['minunderscorelength'])
output_righttoleft = str(config['OUTPUT']['righttoleft'])
output_charstoremove = list(config['OUTPUT']['charstoremove'])

document = Document()
paragraph = document.add_paragraph()

style = document.styles['Normal']
font = style.font

paragraph_format = paragraph.paragraph_format

paragraph_format.line_spacing = output_linespacing

font.name = output_font
font.size = Pt(output_fontsize)

output_files = []

for file in input_files:
    new_text = []

    input_file = str(input_folder_name + file)

    try:
        i = open(input_file, "r", encoding='utf-8')
        inpt = i.read()
        i.close()
    except UnicodeDecodeError:
        i = open(input_file, "r", encoding='windows-1255')
        inpt = i.read()
        i.close()

    for char in output_charstoremove:
        inpt = inpt.replace(char, '')

    inpt = inpt.splitlines()

    for t in inpt:
        if len(t) >= 1:
            count = 0
            t = t.rstrip()
            t = t.lstrip()

            words = t.split()
            extra_chars = len(t) - output_linecharlength
            t += " "

            if extra_chars >= 0:
                word_length = 0
                for word in range(len(words)):
                    word_length += len(words[-(word + 1)])
                    if extra_chars == word_length:
                        if output_righttoleft.lower() == "true":
                            t = t.rjust((math.ceil(len(t) / output_linecharlength) * output_linecharlength) - 1, "_")
                        else:
                            t = t.ljust((math.ceil(len(t) / output_linecharlength) * output_linecharlength) - 1, "_")
                        break
                    elif extra_chars == word_length + 1:
                        if output_righttoleft.lower() == "true":
                            t = t.rjust((math.ceil(len(t) / output_linecharlength) * output_linecharlength) + 1, "_")
                        else:
                            t = t.ljust((math.ceil(len(t) / output_linecharlength) * output_linecharlength) + 1, "_")
                        break
                    elif extra_chars < word_length + 1:
                        if output_righttoleft.lower() == "true":
                            t = t.rjust(((math.ceil(len(t) / output_linecharlength) * output_linecharlength) + 1) -
                                        word_length, "_")
                        else:
                            t = t.ljust(((math.ceil(len(t) / output_linecharlength) * output_linecharlength) + 1) -
                                        word_length, "_")
                        break
                    else:
                        pass

            else:
                if output_righttoleft.lower() == "true":
                    t = t.rjust(output_linecharlength, "_")
                else:
                    t = t.ljust(output_linecharlength, "_")
                if abs(extra_chars) < output_minunderscorelength:
                    if output_righttoleft.lower() == "true":
                        t = " " + t
                        t = "_" * output_linecharlength + t
                    else:
                        t += " "
                        t += "_" * output_linecharlength
            new_text.append(t)
        else:
            new_text.append('')

    output = '\n'.join(new_text)

    paragraph.text = output

    output_name = output_file
    increment = 0

    while os.path.exists("%s.docx" % output_name):
        increment += 1
        output_name = "%s-%s" % (output_file, increment)

    if not os.path.exists(output_folder_name):
        os.mkdir(output_folder_name)

    document_name = '%s.docx' % output_name
    document.save(document_name)
    output_files.append(document_name)

for output_file in output_files:
    os.system('start %s' % output_file)
