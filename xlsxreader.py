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
from datetime import datetime
from subprocess import Popen, CREATE_NEW_CONSOLE
from datetime import date
import openpyxl
from pathlib import Path
import shutil, os
import pandas as pd


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
    if rtg2 == 0:
        return 0
    if rtg1 == 0:
        return 0
    return float(k * (float(result) - E1))  # Rating deviation


def findplayer(name, size):
    for i in range(1, size):
        if name in str(rtgfnx["B" + str(i)].value):
            return [rtgfnx["A" + str(i)].value, rtgfnx["B" + str(i)].value, rtgfnx["J" + str(i)].value]

    return 0


def findplayerinlist(name, player):
    for i in range(0, len(player)):
        if player[i] != 0:
            if player[i][1] == name:
                print(name, player[i][2])
                return player[i][2]
    return 0


print("Starting")

xlsx_file = Path('fnxlist', 'rtgfnx.xlsx')
wb_obj = openpyxl.load_workbook(xlsx_file, data_only=True)
rtgfnx = wb_obj.active
row_countfnx = int(rtgfnx.max_row)
column_countfnx = int(rtgfnx.max_column)

tournamentstest = os.listdir('tournaments')
# print(tournamentstest)
for i in range(len(tournamentstest)):
    if ".xls" in tournamentstest[i]:
        # print("true")
        # Actually to change .xls to .xlsx use https://convertio.co/pt/xls-xlsx/ or something
        xlsx_file = Path('tournaments', tournamentstest[i])
        wb_obj = openpyxl.load_workbook(xlsx_file, data_only=True)
        tournament = wb_obj.active
        row_counttournament = int(tournament.max_row)
        column_counttournament = int(tournament.max_column)

        today = date.today()
        # shutil.copy(Path('fnxlist', 'rtgfnx.xlsx'), 'backup')
        # old_name = Path('backup', 'rtgfnx.xlsx')
        # new_name = Path('backup', 'rtgfnx'+today.strftime("%d%m%Y%H%M")+'.xlsx')
        # os.rename(old_name, new_name)
        players = []
        for j in range(1, row_counttournament):
            if tournament["E" + str(j)].value == "Name:":
                print('\n')
                print(tournament["G" + str(j)].value)
                players.append(findplayer(str(tournament["G" + str(j)].value), row_countfnx))
                for l in range(1, 1000):
                    if tournament["A" + str(j + l)].value == "Rd.":
                        variation = 0
                        for m in range(1, 20):
                            if tournament["A" + str(j + l + m)].value is None:
                                break
                            # print(tournament["A" + str(j + l + m)].value, m)
                            rating1 = findplayerinlist(tournament["G" + str(j)].value, players)
                            rating2 = findplayerinlist(tournament["D" + str(j + l + m)].value, players)
                            variation = variation + var(rating1, rating2, tournament["H" + str(j + l + m)].value, 0)
                        print(variation)
                        break
        print(players)

        # shutil.move(xlsx_file, Path('savedtournaments', tournamentstest[i]))
print("end")
