import math
import sys
import subprocess
import os
import itertools
import random
import codecs
import time
import numpy as np
from random import seed
from io import StringIO
from random import randint
from subprocess import Popen, CREATE_NEW_CONSOLE
from datetime import datetime
from pathlib import Path
import shutil, os
import pandas as pd
import xlsxwriter
import openpyxl


def var(rtg1, rtg2, result, t):  # Function rating deviation

    k = 0  # Setting K value for classical/rapid/blitz
    if t == 0:
        k = 25
    elif t == 1:
        k = 15
    elif t == 2:
        k = 10
    else:
        return 0

    if rtg1 - rtg2 > 400:  # Capping rating dif at 400
        rtg1 = rtg2 + 400
    elif rtg2 - rtg1 > 400:
        rtg2 = rtg1 + 400

    Q1 = pow(10, (rtg1 / 400))
    Q2 = pow(10, (rtg2 / 400))
    E1 = Q1 / (Q1 + Q2)
    # E2 = Q2/(Q1+Q2)
    if ord(result) == 189:  # Actually ½ isn't a number for some reason
        result = 0.5
    if rtg1 == 0:
        return 0
    if rtg2 == 0:
        return 0
    return round(float(k * (float(result) - E1)), 2)  # Rating deviation


def findplayer(name, size):
    # xlsx_file1a = Path('fnxlist', 'rtgfnx.xlsx')
    # wb_obj1a = openpyxl.load_workbook(xlsx_file1a, data_only=True)
    # rtgfnx1a = wb_obj1a.active
    for i in range(9, size):
        if name in str(rtgfnx["B" + str(i)].value):
            # print([rtgfnx["A" + str(i)].value, rtgfnx["B" + str(i)].value, rtgfnx["J" + str(i)].value])
            return [rtgfnx["A" + str(i)].value, rtgfnx["B" + str(i)].value, rtgfnx["J" + str(i)].value]
    return 0


def findplayerinlist(name, player):
    for i in range(0, len(player)):
        if player[i] != 0:
            if player[i][1] == name:
                # print(name, player[i][2])
                return player[i][2]
    return 0


def ratingperformance(result, averagertg, perftable):
    if result == 0:
        return round(averagertg - 500, 2)
    if result == 9:
        return round(averagertg + 500, 2)
    if result == 4.5:
        return round(averagertg, 2)
    if result < 4.5:
        return round(averagertg + perftable[1][int(result / 0.5) - 1], 2)
    if result > 4.5:
        return round(averagertg + perftable[3][(int(result / 0.5)) - 9], 2)
    return 0


def round_off_rating(number):
    return round(number * 2) / 2


pm = np.zeros((4, 8), dtype=float)
np.set_printoptions(suppress=True)
pm[0][0] = 0.06
pm[0][1] = 0.11
pm[0][2] = 0.17
pm[0][3] = 0.22
pm[0][4] = 0.38
pm[0][5] = 0.33
pm[0][6] = 0.39
pm[0][7] = 0.44
pm[1][0] = -444
pm[1][1] = -351
pm[1][2] = -273
pm[1][3] = -220
pm[1][4] = -166
pm[1][5] = -125
pm[1][6] = -80
pm[1][7] = -43
pm[2][0] = 0.56
pm[2][1] = 0.61
pm[2][2] = 0.67
pm[2][3] = 0.72
pm[2][4] = 0.78
pm[2][5] = 0.83
pm[2][6] = 0.89
pm[2][7] = 0.94
pm[3][0] = 43
pm[3][1] = 80
pm[3][2] = 125
pm[3][3] = 166
pm[3][4] = 220
pm[3][5] = 273
pm[3][6] = 351
pm[3][7] = 444
# print(pm)

print("Starting")

tournamentstest = os.listdir('tournaments')
# print(tournamentstest)
for i in range(len(tournamentstest)):
    if ".xlsx" in tournamentstest[i]:
        # print("true")
        # Actually to change .xls to .xlsx use https://convertio.co/pt/xls-xlsx/ or something

        xlsx_file1 = Path('fnxlist', 'rtgfnx.xlsx')
        wb_obj1 = openpyxl.load_workbook(xlsx_file1, data_only=True)
        rtgfnx = wb_obj1.active
        row_countfnx = int(rtgfnx.max_row)
        column_countfnx = int(rtgfnx.max_column)
        maxfnx = 0
        temp11 = 0

        for ia in range(9, row_countfnx + 1):
            # print(i, rtgfnx["A" + str(i)].value)
            if rtgfnx["A" + str(ia)].value is None:
                temp11 = temp11 + 1
            if not rtgfnx["A" + str(ia)].value is None:
                temp10 = int(rtgfnx["A" + str(ia)].value)
                if maxfnx < temp10:
                    maxfnx = temp10
        row_countfnx = row_countfnx - temp11
        print('maxfnx:', maxfnx, row_countfnx)

        xlsx_file2 = Path('tournaments', tournamentstest[i])
        wb_obj2 = openpyxl.load_workbook(xlsx_file2, data_only=True)
        tournament = wb_obj2.active
        row_counttournament = int(tournament.max_row)
        column_counttournament = int(tournament.max_column)

        shutil.copy(Path('fnxlist', 'rtgfnx.xlsx'), 'backup')
        old_name = Path('backup', 'rtgfnx.xlsx')
        new_name = Path('backup', 'rtgfnx' + str(datetime.now().strftime("%Y_%m_%d_%I_%M_%S_%p")) + '.xlsx')
        os.rename(old_name, new_name)
        players = []

        xlsx_file3 = Path('fnxlist', 'perfrtgfnx.xlsx')
        wb_obj3 = openpyxl.load_workbook(xlsx_file3, data_only=True)
        perfrtgfnx = wb_obj3.active
        row_countperfrtgfnx = int(perfrtgfnx.max_row)
        column_countperfrtgfnx = int(perfrtgfnx.max_column)

        time.sleep(10)

        for j in range(1, row_counttournament):
            # time.sleep(0.5)
            if tournament["E" + str(j)].value == "Name:":
                # print('\n')
                temp12 = findplayer(str(tournament["G" + str(j)].value), row_countfnx)
                # print(temp12, tournament["G" + str(j)].value)
                if temp12 != 0:
                    # print("True", temp12)
                    players.append(temp12)
                    rating1 = temp12[2]
                else:
                    # print("False")
                    players.append(findplayer(str(tournament["G" + str(j)].value), row_countfnx))
                    rating1 = findplayerinlist(tournament["G" + str(j)].value, players)
                # print('Rating1:', rating1)
                for l in range(1, 1000):
                    if tournament["A" + str(j + l)].value == "Rd.":
                        # print('\n')
                        variation = 0
                        for m in range(1, 20):
                            if tournament["A" + str(j + l + m)].value is None:
                                break
                            # print(tournament["A" + str(j + l + m)].value, m)
                            rating2 = findplayerinlist(tournament["D" + str(j + l + m)].value, players)
                            if rating1 == 0:
                                perfc = 0
                                score = 0
                                n1 = 0
                                for n in range(1, 20):
                                    if tournament["A" + str(j + l + n)].value is None:
                                        break
                                    temp2 = findplayerinlist(tournament["D" + str(j + l + n)].value, players)
                                    temp1 = str(tournament["H" + str(j + l + n)].value)
                                    if ord(temp1) == 189:  # Actually ½ isn't a number for some reason
                                        temp1 = 0.5
                                    if temp2 != 0:
                                        n1 = n1 + 1
                                        perfc = perfc + temp2
                                        score = round(score + float(temp1), 1)
                                # print(perfc, score, n1)

                                # print(ratingperformance(score, perfc, pm))

                                # print('size', row_countperfrtgfnx, column_countperfrtgfnx)
                                test = False
                                for o in range(1, row_countperfrtgfnx):
                                    # print("Row Count Fnx Debug: ", row_countfnx)
                                    # print('debug', tournament["G" + str(j)].value, perfrtgfnx["B" + str(o)].value)
                                    if tournament["G" + str(j)].value == str(perfrtgfnx["B" + str(o)].value):
                                        temp5 = round(float(perfrtgfnx["J" + str(o)].value) + score, 2)
                                        # print("debug:", perfrtgfnx["J" + str(o)].value, score, temp5)
                                        if temp5 >= 2:
                                            # wb2 = openpyxl.load_workbook(xlsx_file1)
                                            # ws2 = wb2.active  # or wb.active
                                            rtgfnx['A' + str(row_countfnx + 1)] = str(perfrtgfnx["A" + str(o)].value)
                                            rtgfnx['B' + str(row_countfnx + 1)] = str(perfrtgfnx["B" + str(o)].value)
                                            rtgfnx['J' + str(row_countfnx + 1)] = ratingperformance(float(
                                                6 * int(perfrtgfnx["J" + str(o)].value) / (
                                                    int(perfrtgfnx["C" + str(o)].value))), round(
                                                int(perfrtgfnx["D" + str(o)].value) / (
                                                    int(perfrtgfnx["C" + str(o)].value)), 0), pm)
                                            rtgfnx['CS' + str(row_countfnx + 1)] = 1
                                            row_countfnx = row_countfnx + 1
                                            # wb2.save(xlsx_file1)
                                        perfrtgfnx['C' + str(o)] = perfrtgfnx["C" + str(o)].value + n1
                                        perfrtgfnx['D' + str(o)] = perfrtgfnx["D" + str(o)].value + perfc
                                        perfrtgfnx['J' + str(o)] = perfrtgfnx["J" + str(o)].value + score
                                        test = True
                                        break
                                if not test:
                                    # print('perfdata')
                                    # print(maxfnx)
                                    # worksheet.write('A3', 'teste')
                                    # workbook.close()
                                    # worksheet.write('D' + str(row_countperfrtgfnx), 'test1')
                                    # print('size', row_countperfrtgfnx, column_countperfrtgfnx)
                                    perfrtgfnx['A' + str(row_countperfrtgfnx + 1)] = maxfnx + 1
                                    maxfnx = maxfnx + 1
                                    perfrtgfnx['B' + str(row_countperfrtgfnx + 1)] = tournament["G" + str(j)].value
                                    temp4 = perfrtgfnx['C' + str(row_countperfrtgfnx + 1)].value
                                    if type(temp4) == int:
                                        perfrtgfnx('C' + str(row_countperfrtgfnx + 1), n1 + int(temp4))
                                    else:
                                        perfrtgfnx['C' + str(row_countperfrtgfnx + 1)] = n1
                                    perfrtgfnx['D' + str(row_countperfrtgfnx + 1)] = perfc
                                    perfrtgfnx['J' + str(row_countperfrtgfnx + 1)] = score
                                    row_countperfrtgfnx = row_countperfrtgfnx + 1
                                # wb2.save(xlsx_file1)
                                # wb.save(xlsx_file3)
                                break
                            else:
                                # print("debug:", m, rating1, rating2)
                                # print(rating1, rating2)
                                variation = round(
                                    variation + var(rating1, rating2, tournament["H" + str(j + l + m)].value, 0), 2)
                                # wb2 = openpyxl.load_workbook(xlsx_file1)
                                # ws2 = wb2.active  # or wb.active

                            # wb2.save(xlsx_file1)
                        temp14 = tournament["G" + str(j)].value
                        # print(temp14, variation)
                        if rating1 != 0:
                            d = 1
                            for ia in range(9, row_countfnx):
                                if temp14 in str(rtgfnx["B" + str(ia)].value):
                                    d = ia
                            print(d, row_countfnx, temp14, rtgfnx["J" + str(d)].value, variation)
                            rtgfnx['J' + str(d)] = round(int(rtgfnx["J" + str(d)].value) + variation, 0)
                            rtgfnx['CS' + str(d)] = int(rtgfnx["CS" + str(d)].value) + 1
                        break
        wb_obj1.save(xlsx_file1)
        wb_obj3.save(xlsx_file3)
        # wb2.save(xlsx_file1)
        print(players)

        # shutil.move(xlsx_file, Path('savedtournaments', tournamentstest[i]))
print("end")
