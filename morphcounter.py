#!/usr/bin/python
#-*- coding: utf-8 -*-

import re

import xlrd, xlwt
from nltk import FreqDist
from pymorphy2 import MorphAnalyzer

morph = MorphAnalyzer()
n = 1

def count_morpheme(sentence):
    tokenized_sentence = re.split(r'[\s+\t\n\.\|\:\/\,\?\!\"()]+', sentence.lower())

    final_list = []
    for w in tokenized_sentence[:-1]:
        word = morph.parse(w)[0].tag.POS
        if word is not None:
            final_list.append(word)

    fdist = FreqDist(final_list)

    # add output
    data = ''
    for key in fdist.keys():
        data += key + ": " + str(fdist[key]) + ';'

    num = "TOTAL: " + str(len(final_list))
    write_data(sentence, data, num)

def write_data(sent, morph_data, number):
    global n
    ws.write(n, 0, sent); ws.write(n, 1, morph_data); ws.write(n, 2, number)
    n += 1

# write the file
wb = xlwt.Workbook()
ws = wb.add_sheet('Statistics')
ws.write(0, 0, 'Sentence')
ws.write(0, 1, 'POS stat')
ws.write(0, 2, '# of words')

# read the file
rb = xlrd.open_workbook('input.xls',formatting_info=True) #input.xls - file with data
sheet = rb.sheet_by_index(0)
for rownum in range(sheet.nrows)[1:]:
    row = sheet.row_values(rownum)
    for sent in row:
        count_morpheme(sent)

wb.save('output.xls') #output.xls - file for result
