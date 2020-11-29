from flask import Flask, render_template, Markup, request, send_file
import pandas as pd
import regex as re
import xlsxwriter as xl

app = Flask(__name__)

descriptions = None
projectName = ''
xldata = None
start = 0
end = 0

@app.route('/')
def hello_world():
    return render_template('searchtool.html')

@app.route('/download')
def downloadFile ():
    #For windows you need to use drive name [ex: F:/Example.pdf]
    path = "static/results.xlsx"
    return send_file(path, as_attachment=True)

@app.route('/search', methods=['POST', 'GET'])
def search():
    if request.method == 'POST':
        data = request.form.to_dict()
        query = data['query']

        download = "<a class='dwna' href='/download' target='blank'><button class='dwn'>Download As Excel</button></a>"

        resString = ''
        countString = ''

        if query != '':
            resString, countString = lookUpQuery(query)

            if resString == '':
                resString = "<p class='bline bor_1'></p>"

            if countString == '':
                countString = "<p class='count'>No Matches Found</p>"
                download = ''

        return render_template('searchtool.html', count=Markup(countString), results=Markup(resString), download=Markup(download))

def readData(sheetname):
    global xldata, descriptions, start, end, projectName

    xl_file = pd.ExcelFile(sheetname)
    dfs = {sheet_name: xl_file.parse(sheet_name)
              for sheet_name in xl_file.sheet_names}

    sheetNames = list(dfs.keys())
    titleSheet = None
    xldata = None

    if sheetname == 'static/data1.xlsx':
        titleSheet = dfs[sheetNames[0]]
        xldata = dfs[sheetNames[1]]
    elif sheetname == 'static/data2.xls':
        titleSheet = dfs[sheetNames[1]]
        xldata = dfs[sheetNames[2]]
    elif sheetname == 'static/data3.xlsx':
        xldata = dfs[sheetNames[1]]

    projectName = ''
    if sheetname == 'static/data3.xlsx':
        projectName = 'Residence at Kundli'
    else:
        N = titleSheet.shape[0]
        for i in range(N):
            if 'name' in str(titleSheet.iloc[i][0]).lower():
                projectName = str(titleSheet.iloc[i][0]).split(':-')[1].strip()

    if sheetname == 'static/data1.xlsx':
        start = 8
    elif sheetname == 'static/data2.xls':
        start = 15
    elif sheetname == 'static/data3.xlsx':
        start = 4

    end = xldata.shape[0]-1

    descriptions = xldata.iloc[:, 1]


def markUpLine(line, query):
    global projectName
    res = ''
    resXL = []

    lineNaN = line.isnull()

    res += "<p class='c1 bor_1 cen pad_5'>"
    if not lineNaN[0]:
        if type(line[0]) == float:
            res += "%.1f" % line[0]
            resXL.append("%.1f" % line[0])
        else:
            res += str(line[0])
            resXL.append(str(line[0]))
    else:
        resXL.append("")
    res += "</p>\n"

    lineNaN = line.isnull()

    res += "<p class='c2 bor_1 pad_5'>"
    if not lineNaN[1]:
        desc = str(line[1])
        resXL.append(desc)
        desc = desc.replace(query.lower(), "<i class=highlight>" + query + "</i>")
        desc = desc.replace(query.capitalize(), "<i class=highlight>" + query.capitalize() + "</i>")
        res += desc
    else:
        resXL.append("")
    res += "</p>\n"

    for i in range(2, 5):
        res += "<p class='c" + str(i+1) + " bor_1 cen pad_5'>"
        s = '' if lineNaN[i] else str(line[i])
        res += s
        resXL.append(s)
        res += "</p>\n"

    for i in range(4):
        res += "<p class='c" + str(i+6) + " bor_1 cen pad_5'></p>\n"
        resXL.append("")

    return (res, resXL)

def create_xlsx(name):
    return xl.Workbook(name)

def add_worksheet(workbook):
    return workbook.add_worksheet()

def writeLineXL(worksheet, i, b9, rsXL):
    cols = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']

    for idx in range(len(rsXL)):
        worksheet.write(cols[idx]+str(i), rsXL[idx], b9)

    return i+1

def lookUpQuery(query):
    global xldata, descriptions, start, end, projectName
    sheetnames = ["static/data1.xlsx", "static/data2.xls", "static/data3.xlsx"]

    resString = ''
    finCount = 0
    f = open('static/b.txt', 'w')
    workbook = create_xlsx('static/results.xlsx')
    worksheet = add_worksheet(workbook)
    worksheet.set_column(1, 1, 50)
    worksheet.set_default_row(25)

    b9 = workbook.add_format({'border': 1, 'font_size': 9, 'text_wrap': True})

    worksheet.write('A1', 'Item No.', b9)
    worksheet.write('B1', 'Description', b9)
    worksheet.write('C1', 'Unit', b9)
    worksheet.write('D1', 'Quantity', b9)
    worksheet.write('E1', 'Unit Rate', b9)
    worksheet.write('F1', 'Location', b9)
    worksheet.write('G1', 'Client', b9)
    worksheet.write('H1', 'Start Year', b9)
    worksheet.write('I1', 'End Year', b9)
    si = 2

    for sheetname in sheetnames:
        readData(sheetname)
        resultIdxs = []
        for i in range(start, end):
            if query in str(descriptions[i]).lower():
                resultIdxs.append(i)

        res, count = getSections(resultIdxs, xldata)
        finCount += count

        resIdxs = []
        for idx in res:
            if idx is not None:
                line = xldata.iloc[idx]
                if 'total :-' in str(line[1]).lower():
                    continue
                if line.isnull()[1]:
                    continue

            if not idx:
                resIdxs.append(None)
            elif idx not in resIdxs:
                resIdxs.append(idx)

        borderline = "<p class='proj'>"+projectName+"</p>\n"

        for idx in resIdxs:
            if idx is None:
                resString += borderline
                worksheet.merge_range('A'+str(si)+':I'+str(si), projectName, b9)
                si += 1
            else:
                line = xldata.iloc[idx]
                rs, rsXL = markUpLine(line, query)
                resString += rs
                f.write(str(i) + '\n')
                si = writeLineXL(worksheet, si, b9, rsXL)

    if finCount == 0:
        countString = ''
    else:
        countString = "<p class='count'>"+str(finCount)+" Matches Found</p>"

    workbook.close()

    return (resString, countString)

def checkItemNo(line):
    lineNaN = line.isnull()
    if lineNaN[0]:
        return False

    lString = str(line[0])
    if re.search(r'^\d+.+\d+.+\d+$', lString) is not None:
        return True

    if re.search(r'^\d+.+\d+$', lString) is not None:
        return True

    if re.search(r'^\d+$', lString) is not None:
        return True

    return False

def checkLine(line):
    lString = str(line[0])
    if re.search(r'^\d+.+\d+$', lString) is not None:
        return True

    if re.search(r'^\d+$', lString) is not None:
        return True

    return False

def getSections(resultIdxs, data):
    res = []
    count = 0

    for idx in resultIdxs:
        if idx in res:
            continue

        if checkLine(data.iloc[idx]):
            continue

        res.append(None)
        count += 1

        start = idx
        while not checkItemNo(data.iloc[start]):
            start -= 1

        end = idx+1
        if end < data.shape[0]-1:
            while not checkItemNo(data.iloc[end]):
                end += 1
                if end >= data.shape[0]-1:
                    break

        res += [x for x in range(start, end)]
        if end == idx:
            res.append(idx)

    return (res, count)
