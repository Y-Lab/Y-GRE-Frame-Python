# -*- coding: utf-8 -*-

import os
import sys
import xlrd
import datetime
import hashlib
import base64
import json
from basic.common import getrootpath, makeDirsForFile, existFile
from basic.tools import date2dateStr, countDays
from basic.data import WEEKDAY2STR

reload(sys)
sys.setdefaultencoding('utf-8');


def main():
    rootpath = getrootpath()
    inputpath = os.path.join(rootpath, 'Data')
    outputpath = os.path.join(rootpath, 'Results')

    # name = '田哲予'
    name = raw_input('Please input a name (e.g.: 张三): ')
    print '--'

    YGREFrameFile = os.path.join(inputpath, 'Y-GRE-Frame-%s.xls' % name)
    YGREFrameFileX = os.path.join(inputpath, 'Y-GRE-Frame-%s.xlsx' % name)

    YGREFrameData = {
        'general': {
            'basics': {},
            'scores': {
                'gre': {},
                'toefl': {},
                'gaokao': {},
            },
            'application': {},
            'framework': {},
            'assessment': {},
        },
        'schedule': [],
        'remarks': [],
    }


    # Read Y-GRE Frame data from Y-GRE-Frame-Name.xls(x)
    try:
        excelFile = YGREFrameFile
        workbook = xlrd.open_workbook(excelFile)
    except Exception, e:
        excelFile = YGREFrameFileX
        workbook = xlrd.open_workbook(excelFile)
    print 'Import Y-GRE Frame data: %s' % excelFile


    # Read Y-GRE Frame data: General
    sheet = workbook.sheet_by_index(0)
    values = []
    for row in range(sheet.nrows):
        ctype = sheet.cell(row, 1).ctype
        value = sheet.cell(row, 1).value
        if ctype == 0:
            # empty
            value = 'N/A'
        elif ctype == 1:
            # string
            pass
        elif ctype == 2:
            # number
            if row == 45:
                value = str(value * 100)
            else:
                value = str(value)
            if value[-2:] == '.0':
                value = value[:-2]
            if row == 45:
                value = value + '%'
        elif ctype == 3:
            # date
            value = datetime.date(*xlrd.xldate_as_tuple(value, workbook.datemode)[:3])
            value = date2dateStr(value)
        else:
            print 'Warning: invalid ctype:', ctype
            value = 'Invalid'
        values.append(value)

    YGREFrameData['general'] = {
        'basics': {
            'name': values[0],
            'vb_class': values[1],
            'y_gre_class': values[2],
            'university': values[3],
            'department': values[4],
            'enrollment_year': values[5],
            'degree': values[6],
            'gpa': values[7],
            'gpa_full_marks': values[8],
        },
        'scores': {
            'gre': {
                'v_initial': values[9],
                'q_initial': values[10],
                'aw_initial': values[11],
                'v_admission': values[12],
                'q_admission': values[13],
                'aw_admission': values[14],
                'v_ppii_1': values[15],
                'q_ppii_1': values[16],
                'v_ppii_2': values[17],
                'q_ppii_2': values[18],
                'v_aim': values[19],
                'q_aim': values[20],
                'aw_aim': values[21],
            },
            'toefl': {
                'total': values[22],
                'reading': values[23],
                'listening': values[24],
                'speaking': values[25],
                'writing': values[26],
                'aim': values[27],
            },
            'gaokao': {
                'total': values[28],
                'full_marks': values[29],
                'math': values[30],
                'english': values[31],
            },
        },
        'application': {
            'country': values[32],
            'major': values[33],
            'degree': values[34],
            'aim': values[35],
            'agency': values[36],
        },
        'framework': {
            'g0_start': values[37],
            'g0_end': values[38],
            'g0_days': str(countDays(values[37], values[38])),
            'g1_start': values[39],
            'g1_end': values[40],
            'g1_days': str(countDays(values[39], values[40])),
            'g2_start': values[41],
            'g2_end': values[42],
            'g2_days': str(countDays(values[41], values[42])),
            'deadline': values[43],
            'required_time': values[44],
        },
        'assessment': {
            'applicability': values[45],
            'assessor': values[46],
            'first_executive_supervisor': values[47],
            'second_executive_supervisor': values[48],
        },
    }


    # Read Y-GRE Frame data: Schedule
    sheet = workbook.sheet_by_index(1)
    weekCount = 0
    weekDict = {
        'monday': {
            'tasks': [],
        },
        'tuesday': {
            'tasks': [],
        },
        'wednesday': {
            'tasks': [],
        },
        'thursday': {
            'tasks': [],
        },
        'friday': {
            'tasks': [],
        },
        'saturday': {
            'tasks': [],
        },
        'sunday': {
            'tasks': [],
        },
        'nota_bene': [],
    }
    YGREFrameData['schedule'].append(weekDict)

    # Week date prefix
    firstDay = datetime.date(*xlrd.xldate_as_tuple(sheet.cell(0,0).value, workbook.datemode)[:3])
    for delta in range(firstDay.weekday(), 0, -1):
        date = firstDay - datetime.timedelta(delta)
        weekday = date.weekday()
        dateStr = date2dateStr(date)
        YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['date'] = dateStr
        YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['display'] = False

    # Week date contents
    for row in range(sheet.nrows):
        values = sheet.row_values(row)
        date = datetime.date(*xlrd.xldate_as_tuple(values[0], workbook.datemode)[:3])
        weekday = date.weekday()
        dateStr = date2dateStr(date)
        YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['date'] = dateStr
        YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['display'] = True
        for value in values[1:]:
            if value != '':
                YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['tasks'].append(value)
        if weekday == 6:
            weekCount += 1
            weekDict = {
                'monday': {
                    'tasks': [],
                },
                'tuesday': {
                    'tasks': [],
                },
                'wednesday': {
                    'tasks': [],
                },
                'thursday': {
                    'tasks': [],
                },
                'friday': {
                    'tasks': [],
                },
                'saturday': {
                    'tasks': [],
                },
                'sunday': {
                    'tasks': [],
                },
                'nota_bene': [],
            }
            YGREFrameData['schedule'].append(weekDict)

    # Week date suffix
    if weekday == 6:
        YGREFrameData['schedule'].pop(-1)
    else:
        lastDay = date
        for delta in range(1, 7 - lastDay.weekday()):
            date = lastDay + datetime.timedelta(delta)
            weekday = date.weekday()
            dateStr = date2dateStr(date)
            YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['date'] = dateStr
            YGREFrameData['schedule'][weekCount][WEEKDAY2STR[weekday]]['display'] = False


    # Read Y-GRE Frame data: N.B.
    sheet = workbook.sheet_by_index(2)
    for row in range(sheet.nrows):
        values = sheet.row_values(row)
        for value in values[1:]:
            if value != '':
                YGREFrameData['schedule'][row]['nota_bene'].append(value)


    # Read Y-GRE Frame data: Remarks
    sheet = workbook.sheet_by_index(3)
    values = sheet.col_values(0)
    for value in values:
        YGREFrameData['remarks'].append(value)


    # Generate passkey
    password = YGREFrameData['general']['basics']['name']
    salt = YGREFrameData['general']['basics']['vb_class'] + YGREFrameData['general']['basics']['y_gre_class'] + YGREFrameData['general']['basics']['university']
    passkey = base64.urlsafe_b64encode(hashlib.pbkdf2_hmac('sha256', password, salt, 100000, dklen=24))
    print 'Passkey:', passkey

    # Write to passkey.js file
    jsonFilename = '%s.js' % passkey
    print 'Write to file: %s' % jsonFilename
    jsonFile = os.path.join(outputpath, jsonFilename)
    makeDirsForFile(jsonFile)
    templateContent = 'frame_data = %s;'
    with open(jsonFile, 'w') as f:
        f.write(templateContent % json.dumps(YGREFrameData))


if __name__ == '__main__':
    main()