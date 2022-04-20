from posixpath import split
import xlrd
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import re

data = xlrd.open_workbook("safety.xlsx")
sheetsnumber = data.nsheets


def word(item, file):
    doc1 = Document()
    doc1.styles['Normal'].font.name = u'宋体'
    doc1.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc1.styles['Normal'].font.size = Pt(9)
    doc1.add_paragraph(''.join(item))
    doc1.save(file)


def joinlist(item, fuck):
    letters = ['B', 'C', 'D', 'E', 'F', 'G', 'H']
    item[0] = str(item[0]).replace("\n", "") + "\n"
    item[1] = ("A." + str(item[1]).replace("\n", "") if (item[1])
               and ("".join(str(item[1]).split()) != "") else "")
    if fuck == 3:
        item[fuck - 1] = " 答案：" + str(item[fuck - 1]).replace("\n", "") + "\n"
        return ''.join(item)
    if fuck > 3:
        for i in range(fuck - 3):
            item[2+i] = (" " + str(letters[i]) + "." + str(item[2+i]).replace("\n", "")
                         if (item[2+i]) and ("".join(str(item[2+i]).split()) != "") else "")
        item[fuck - 1] = " 答案：" + str(item[fuck - 1]).replace("\n", "") + "\n"
        return ''.join(item)


def process(items):
    questionindex = []
    question = []
    answer = []
    choice = []
    modelregexone = re.compile(r'题目')
    modelregextwo = re.compile(r'答案')
    modelregexthree = re.compile(r'A|B|C|D|E|F|G|H', re.I)
    for item in items:
        if modelregexone.search(item):
            question.append(items.index(item))
        if modelregextwo.search(item):
            answer.append(items.index(item))
        if modelregexthree.search(item):
            choice.append(items.index(item))
    if len(question) != 0:
        questionindex.append(question[0])
    if len(choice) != 0:
        questionindex.extend(choice)
    if len(answer) != 0:
        questionindex.append(answer[0])
    if len(questionindex) != 0:
        return questionindex, len(questionindex)
    return False, False


total = [0 for i in range(sheetsnumber)]
table = [0 for i in range(sheetsnumber)]
nrows = [0 for i in range(sheetsnumber)]
ncols = [0 for i in range(sheetsnumber)]
sum = [0 for i in range(sheetsnumber)]
for i in range(sheetsnumber):
    table[i] = data.sheets()[i]
    nrows[i] = table[i].nrows
    ncols[i] = table[i].ncols
    if nrows[i] != 0:
        temp, fuck = process(table[i].row_values(
            0, start_colx=0, end_colx=ncols[i]))
        if temp:
            sum[i] = []
            for j in range(nrows[i]-1):
                everyrow = []
                index = "(" + str(j+1) + ")"
                for k in temp:
                    everyrow.append(table[i].cell_value(j+1, k))
                everyrowlist = index + joinlist(everyrow, fuck)
                sum[i].append(everyrowlist)
            total[i] = ''.join(sum[i])
            title = str(table[i].name) + "Python.docx"
            word(total[i], title)
