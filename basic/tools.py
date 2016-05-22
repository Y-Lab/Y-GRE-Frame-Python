# -*- coding: utf-8 -*-

import datetime


def date2dateStr(date, seperator='-'):
    return str(date.year) + seperator + str(date.month) + seperator + str(date.day)


def dateStr2date(dateStr):
    dateStr = str(dateStr)
    try:
        yearStr, monthStr, dayStr = dateStr.split('/')
    except Exception, e:
        yearStr, monthStr, dayStr = dateStr.split('-')
    return datetime.date(int(yearStr), int(monthStr), int(dayStr))


def countDays(startDateStr, endDateStr):
    if startDateStr == 'N/A' or endDateStr == 'N/A':
        return 'N/A'
    else:
        startDate = dateStr2date(startDateStr)
        endDate = dateStr2date(endDateStr)
        timeDelta =  endDate - startDate
        return timeDelta.days + 1