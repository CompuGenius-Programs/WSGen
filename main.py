# -*- coding: utf-8 -*-

from docx import Document
from docx.shared import Pt

import os
from os import listdir
from os.path import isfile, join

import ctypes
from ctypes import wintypes

import math


def get_font(font):
    class LOGFONT(ctypes.Structure): _fields_ = [
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
        ('lfFaceName', ctypes.c_wchar * 32)]

    FONTENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_int,
                                      ctypes.POINTER(LOGFONT), wintypes.LPVOID, wintypes.DWORD, wintypes.LPARAM)

    fontlist = []

    def font_enum(logfont, textmetricex, fonttype, param):
        str = logfont.contents.lfFaceName;
        if (any(str in s for s in fontlist) == False):
            fontlist.append(str)
        return True

    hdc = ctypes.windll.user32.GetDC(None)
    ctypes.windll.gdi32.EnumFontFamiliesExW(hdc, None, FONTENUMPROC(font_enum), 0, 0)
    ctypes.windll.user32.ReleaseDC(None, hdc)
    return font.lower() in [x.lower() for x in fontlist]


import configparser

config = configparser.ConfigParser(allow_no_value=True)

must_exit = False
no_input_file = False

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
    }

    print("NO SETTINGS FILE FOUND!")
    print("CREATING ONE! PLEASE VERIFY THE CONFIGURATIONS!")

    with open('settings.ini', 'w') as configfile:
        config.write(configfile)

    must_exit = True

config.read('settings.ini')

input_folder_name = str(config['INPUT']['inputfolder']) + '/'

if not os.path.exists(input_folder_name):
    os.mkdir(input_folder_name)

input_files = [f for f in listdir(input_folder_name) if isfile(join(input_folder_name, f))]

if len(input_files) <= 0:
    print("NO INPUT TEXT FILE FOUND!")
    print("CREATING ONE! PLEASE ADD TEXT TO IT!")

    f = open(str(input_folder_name + "input.txt"), "a")
    f.close()

    must_exit = True
    no_input_file = True

output_font = str(config['OUTPUT']['font'])

if not get_font(output_font):
    print("FONT NOT INSTALLED!")
    print("PLEASE INSTALL %s OR CHANGE THE FONT IN THE SETTINGS FILE!" % output_font)

    must_exit = True

if must_exit:
    print('Restart the PROGRAM after fixing above issues.')
    print('Press ENTER to exit.')
    input()
    exit()

output_folder_name = str(config['OUTPUT']['outputfolder']) + '/'
output_file_name = str(config['OUTPUT']['outputfile'])
output_file = str(output_folder_name + output_file_name)
output_fontsize = float(config['OUTPUT']['fontsize'])
output_linespacing = float(config['OUTPUT']['linespacing'])
output_linecharlength = int(config['OUTPUT']['linecharlength'])
output_minunderscorelength = int(config['OUTPUT']['minunderscorelength'])
output_righttoleft = str(config['OUTPUT']['righttoleft'])

document = Document()
paragraph = document.add_paragraph()

style = document.styles['Normal']
font = style.font

paragraph_format = paragraph.paragraph_format

paragraph_format.line_spacing = output_linespacing

font.name = output_font
font.size = Pt(output_fontsize)


# for file in input_files:
#     new_text = []
#
#     input_file = str(input_folder_name + file)
#
#     i = open(input_file, "r", encoding='utf-8')
#     inpt = i.read()
#     i.close()
#
#     inpt = inpt.splitlines()
#
#     for t in inpt:
#         s = set(t)
#         if len(s) >= 1:
#             t = t.rstrip()
#             t = t.lstrip()
#             if output_righttoleft.lower() == "true":
#                 t = ' ' + t
#                 t = t.rjust(output_linecharlength, "_")
#             else:
#                 t = t + ' '
#                 t = t.ljust(output_linecharlength, "_")
#             new_text.append(t)
#         else:
#             new_text.append('')
#
#     output = '\n'.join(new_text)
#
#     paragraph.text = output
#
#     output_name = output_file
#     increment = 0
#
#     while os.path.exists("%s.docx" % output_name):
#         increment += 1
#         output_name = "%s-%s" % (output_file, increment)
#
#     if not os.path.exists(output_folder_name):
#         os.mkdir(output_folder_name)
#
#     document.save('%s.docx' % output_name)


for file in input_files:
    new_text = []

    input_file = str(input_folder_name + file)

    i = open(input_file, "r", encoding='utf-8')
    inpt = i.read()
    i.close()

    inpt = inpt.splitlines()

    for t in inpt:
        if len(t) >= 1:
            count = 0
            t = t.rstrip()
            t = t.lstrip()
            # if output_righttoleft.lower() == "true":
            #     t = ' ' + t
            #     t = t.rjust(output_linecharlength, "_")
            # else:
            #     t = t + ' '
            #     t = t.ljust(output_linecharlength, "_")
            #
            # for char in t:
            #     if char == "_":
            #         count += 1
            #
            # if count <= 25:
            #     if output_righttoleft.lower() == "true":
            #         t = "_" * 25 + t
            #     else:
            #         t = t + "_" * 25

            words = t.split()
            extra_chars = len(t) - output_linecharlength
            t += " "

            if extra_chars >= 0:
                word_length = 0
                for word in range(len(words)):
                    word_length += len(words[-(word + 1)])
                    if extra_chars == word_length:
                        t = t.ljust((math.ceil(len(t) / output_linecharlength) * output_linecharlength) - 1, "_")
                        break
                    elif extra_chars == word_length + 1:
                        t = t.ljust((math.ceil(len(t) / output_linecharlength) * output_linecharlength) + 1, "_")
                        break
                    elif extra_chars < word_length + 1:
                        t = t.ljust(((math.ceil(len(t) / output_linecharlength) * output_linecharlength) + 1) -
                                    word_length, "_")
                        break
                    else:
                        pass

            else:
                t = t.ljust(output_linecharlength, "_")
                if abs(extra_chars) < output_minunderscorelength:
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

    document.save('%s.docx' % output_name)
