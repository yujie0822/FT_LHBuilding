# -*- coding: utf-8 -*-
import xlrd
import xlsxwriter
import sys
import datetime
import time
stdout = sys.stdout
stdin = sys.stdin
stderr = sys.stderr
reload( sys )
sys.setdefaultencoding('utf-8')
sys.stdout = stdout
sys.stdin = stdin
sys.stderr = stderr

now=datetime.datetime.now()
time = now.strftime('%d%H%M%S')

infoPath = 'D:\Work\Main\InfoBook.xlsx'
rawDataPath = 'D:\Work\Main\RawData.xlsx'
outputPath = 'D:\Work\Main\Output{}.xlsx'.format(time)

"""
-------------------以上为配置部分，编码默认utf-8----------------



-------------------------自定义函数-----------------------------

**************************************************************

                       炒鸡华丽的分割线
**************************************************************

"""


"""

findRowNum: 搜索函数，在list l中搜索a的值，返回对应下标，无值返回False


"""
# value:a,col:l,return a row number
def findRowNum(a,l):
    for x in range(len(l)):
        if a == l[x]:
            return x
    return -1

"""

myTrim: 将unicode转为utf-8 string，
数字转换成不带小数点的utf-8 string形式，不支持带小数点的float型

"""

#unicode转字符串+去小数点
def myTrim(l):
    for x in range(len(l)):
        if type(l[x]) == float:
            if l[x]%1.0 == 0.0:
                l[x] = str(int(l[x]))
            else:
                print '第{}行{}错误'.format(x,l[0])
                l[x] = 'ERROR'
        l[x] = l[x].encode('utf-8').strip()
        if((l[x].find("　") != -1) or (l[x].find(" ") != -1)):
            print "第{}行{}存在空格".format(x,l[0])
            print l[x]

"""
由于excel中的数字不区分int与float,此函数为excel中整型数字去小数点处理函数

"""
#float转int
def myFloatToInt(l):
    l[0] = l[0].encode('utf-8').strip()
    for x in range(len(l)):
        if type(l[x]) == float:
            if l[x]%1.0 == 0.0:
                l[x] = int(l[x])
            else:
                print '第{}行{}异常'.format(x,l[0])
                l[x] = 'ERROR'

"""
excel写入函数
sheet为写入的excel sheet, colNum为列号，colList为写入的List数据

"""

#Excel插入单列
def insertCol(sheet,colNum,colList):
    for x in range(len(colList)):
        sheet.write(x+1,colNum,colList[x])

"""
-------------------------主函数---------------------------------------
**************************************************************

                       炒鸡华丽的分割线
**************************************************************

"""

def main():
    infoBook = xlrd.open_workbook(infoPath)
    rawDataBook = xlrd.open_workbook(rawDataPath)
#Sheet
#产品线对应表
    cpxSheet = infoBook.sheet_by_index(0)
#raw data
    rawSheet = rawDataBook.sheet_by_index(0)
#料号
    lcSheet = infoBook.sheet_by_index(1)
#产品类别sheet
    lbSheet = infoBook.sheet_by_index(2)
#产品子类别sheet
    zlbSheet = infoBook.sheet_by_index(3)
#BU Sheet
    buSheet = infoBook.sheet_by_index(4)
#品牌 Sheet
    brandSheet = infoBook.sheet_by_index(5)

#产品线名称
    cpxInfo = cpxSheet.col_values(2)
    myTrim(cpxInfo)
#产品线代码
    cpxCode = cpxSheet.col_values(0)
    myTrim(cpxCode)
#供应商代码
    gysCode = cpxSheet.col_values(3)
#产品组
    cpzCode =cpxSheet.col_values(4)
    myTrim(cpzCode)
#Segment
    segmentCode = cpxSheet.col_values(6)
    myTrim(segmentCode)
#Licence代码
    lcCode = lcSheet.col_values(0)
    myTrim(lcCode)
#Licence内容
    lcInfo = lcSheet.col_values(1)
    myTrim(lcInfo)
#产品类别
    lbInfo = lbSheet.col_values(1)
    myTrim(lbInfo)
#产品类别代码
    lbCode = lbSheet.col_values(0)
    myTrim(lbCode)
#子类别代码
    zlbCode = zlbSheet.col_values(0)
    myTrim(zlbCode)
#子类别产品线
    zlbCpx = zlbSheet.col_values(2)
    myTrim(zlbCpx)
#子类别名
    zlbInfo = zlbSheet.col_values(3)
    myTrim(zlbInfo)
#BU代码
    buCode = buSheet.col_values(0)
    myTrim(buCode)
#BU内容
    buInfo = buSheet.col_values(2)
    myTrim(buInfo)
#BU天数
    buDay = buSheet.col_values(3)
#品牌
    brandInfo = brandSheet.col_values(1)

#输入数据
    headList = ['产品线','料号','长物料','PN','产品子类别','MPQ',\
    'MOQ','L/T（天）','License','是否NCNR料','品牌',\
    '产品类别','项目号','BU','中文品名','性能及功能描述',\
    '是否进关','是否定制件','应用及领域','尺寸','应用领域',\
    '最惠国税率','普通关税','进关品名','监管条件','HS  CODE',\
    '净重（千克/颗）','一层包装','12nc']

    rawinputList = [[] for x in headList]
    outputList1 = [[] for x in range(25)]
    outputList2 = [[] for x in range(49)]


    inputCol = 0
    for x in range(rawSheet.ncols):
        temp = rawSheet.col_values(x)
        index = findRowNum(temp[0],headList)
        if (index != -1) :
            rawinputList[index] = temp
        elif (temp[0].find('PN') != -1):
            rawinputList[3] = temp
        elif (temp[0].find('HS') != -1):
            rawinputList[25] = temp
        else:
            print "{}列未录入".format(temp[0])

    for x in range(len(rawinputList)):
        if x in [2,3,9,12,13,19,25,26,28]:
            continue
        if (len(rawinputList[x]) == 0):
            print "无--"+headList[x]+"--列"


    for x in range(len(rawinputList)):
        if (rawinputList[x] == []):
            continue
        if (rawinputList[x][0] == 'MPQ') or (rawinputList[x][0] == 'MOQ') or \
        (rawinputList[x][0] == 'L/T（天）') or (rawinputList[x][0] == '一层包装'):
            myFloatToInt(rawinputList[x])
        elif(rawinputList[x][0] == '项目号') or (rawinputList[x][0] == '最惠国税率') or (rawinputList[x][0] == '普通关税')\
         or (rawinputList[x][0].find('HS') != -1) or (rawinputList[x][0] == '净重（千克/颗）') or (rawinputList[x][0] == '12nc'):
            continue
        else:
            myTrim(rawinputList[x])

    cpxList = []
    #PN
    outputList2[8] = rawinputList[3][1:]
    #MPQ,MOQ,L/T
    outputList1[4] = rawinputList[5][1:]
    outputList1[5] = rawinputList[6][1:]
    outputList1[6] = rawinputList[7][1:]
    #一层包装
    outputList2[29] = rawinputList[27][1:]
    for x in range(len(outputList2[29])):
        if outputList2[29][x]  == '':
            print "行{}一层包装为空".format(x+2)
    #12n
    outputList2[9] = rawinputList[28][1:]
    #品牌
    outputList1[9] = rawinputList[10][1:]
#检测品牌是否第一次出现
    for x in range(len(outputList1[9])):
        sTemp = outputList1[9][x]
        if findRowNum(outputList1[9][x],brandInfo) == -1:
            if findRowNum(sTemp.upper(),brandInfo) != -1:
                outputList1[9][x] = sTemp.upper()
            else:
                print "行{}的品牌{}为新品牌".format(x+2,outputList1[9][x])
    #最惠国税率
    outputList1[14] =rawinputList[21][1:]
    #进关品名
    outputList1[15] = rawinputList[23][1:]
    #重量
    outputList1[16] = ["" for x in rawinputList[26][1:]]
    for x in range(len(rawinputList[26][1:])):
        if(type(rawinputList[26][x+1]) == float):
            outputList1[16][x]=1000000000*rawinputList[26][x+1]
        else:
            if(rawinputList[26][x+1].encode('utf-8').strip()!=""):
                print "行{}净重异常".format(x+2)

    #HSCODE
    outputList2[32] = rawinputList[25][1:]
    #长物料
    outputList2[38] = rawinputList[2][1:]
    #附表料号
    outputList2[2] = rawinputList[15][1:]
    #附表普通关税
    outputList2[6] = rawinputList[22][1:]
    #附表应用及领域
    outputList2[3] = rawinputList[18][1:]
    #附表中文品名
    outputList2[4] = rawinputList[14][1:]
    #附表尺寸
    outputList2[5] = rawinputList[19][1:]
#产品线List处理
#转大写
    rawinputList[0] = [s.upper() for s in rawinputList[0]]
    cpxInfo = [s.upper() for s in cpxInfo]
#Chilsin转Chilisin
    for x in range(1,len(rawinputList[0])):
        if rawinputList[0][x] == 'CHILSIN':
            rawinputList[0][x] = 'CHILISIN'
#处理
    for x in range(1,len(rawinputList[0])):
        temp = findRowNum(rawinputList[0][x],cpxInfo)
        if temp != -1:
            cpxList += [temp]
        elif rawinputList[0][x] == 'KYOCERA':
            if rawinputList[4][x].upper() == 'CONNECTOR':
                cpxList += [7]
            elif rawinputList[4][x].upper() == 'CRYSTAL' or rawinputList[4][x].upper() == 'TCXO' \
                 or rawinputList[4][x].upper() == 'SAW FILTER' or rawinputList[4][x].upper() == 'RF MODULE'\
                 or rawinputList[4][x].upper() == 'RESONATOR' or rawinputList[4][x].upper() == 'SAW DUPLEXER'\
                 or rawinputList[4][x].upper() == 'TOOL' or rawinputList[4][x].upper() == 'ETALON FILTER':
                cpxList += [6]
            else :
                cpxList += [999]
        elif findRowNum(rawinputList[0][x],cpxCode) != -1:
            cpxList += [findRowNum(rawinputList[0][x],cpxCode)]
        else:
            cpxList +=[999]

#产品线代码处理
    for x in cpxList:
        if x!=999:
            temp = cpxCode[x]
        else:
            temp = 999
        outputList1[0] += [temp]


#子类别搜索处理
    for x in range(1,len(rawinputList[4])):
        cpx = outputList1[0][x-1]
        if cpx == 999:
            outputList1[3] += ['']
            print '行{}无产品线'.format(x+1)
        else :
            found = False
            zlb = rawinputList[4][x].upper()
            for y in range(1,len(zlbCpx)):
                if (zlbCpx[y].upper() == cpx) and (zlbInfo[y].upper() == zlb):
                    outputList1[3] += [zlbCode[y]]
                    found = True
                    break
            if (not found):
                outputList1[3] += ['']
                print '行{}子类别未找到'.format(x+1)


#料号List处理 不支持float形料号输入
    for x in range(1,len(rawinputList[1])):
        if(len(rawinputList[1][x]) > 25):
            print "行{}料号超长,{}".format((x+1),(len(rawinputList[1])))
        outputList1[2] += [rawinputList[1][x].upper()]


#Licence
    for x in range(1,len(rawinputList[8])):
        temp = findRowNum(rawinputList[8][x],lcInfo)
        if temp != -1:
            outputList1[7] += [lcCode[temp]]
        elif rawinputList[8][x] == '':
            outputList1[7] += ['A02']
        else:
            outputList1[7] += ['Error']
            print '行{}License 未找到!'.format(x+1)

#NCNR
    for x in range(1,len(rawinputList[9])):
        if rawinputList[9][x].find('NCNR') != -1:
            outputList1[8] += ['NCN']
        elif rawinputList[9][x] =='STANDARD':
            outputList1[8] += ['STA']
        elif rawinputList[9][x] =='NCN':
            outputList1[8] += ['NCN']
        elif rawinputList[9][x] =='STA':
            outputList1[8] += ['STA']
        elif rawinputList[9][x] == 'Standard':
            outputList1[8] += ['STA']
        else :
            outputList1[8] += [rawinputList[9][x]]
            print "行{}NCNR Error".format(x+1)

#产品类别
    for x in range(1,len(rawinputList[11])):
        temp = findRowNum(rawinputList[11][x],lbInfo)
        if temp != -1:
            outputList1[10] += [lbCode[temp]]
        else:
            outputList1[10] += ['Error']
            print '行{}产品类别未找到!'.format(x+1)
#BU
    outputList1[11] = ['']*(len(rawinputList[13])-1)
    for x in range(1,len(rawinputList[13])):
        if rawinputList[13][x] != '':
            day = outputList1[6][x-1]/7.0
            found = False
            for y in range(len(buInfo)):
                if ((rawinputList[13][x] == buInfo[y]) and (buDay[y] == day)) :
                    outputList1[11][x-1] = buCode[y]
                    found = True
                    break
            if not found:
                print "行{}BU未找到".format(x+1)


#进关
    outputList1[12] = ['' for x in range(len(outputList1[0]))]
    for x in range(1,(len(outputList1[12])+1)):
        if (len(rawinputList[17])<x):
            if ((rawinputList[16][x] == '') or (rawinputList[16][x] == 'Y') or (rawinputList[16][x] == '是')\
             or (rawinputList[16][x] == '能进关') or (rawinputList[16][x] == '要进关') or (rawinputList[16][x] == '参与进关')\
              or (rawinputList[16][x] == '进关')):
                outputList1[12][x-1] = ''
            elif ((rawinputList[16][x] == '不进关') or (rawinputList[16][x] == 'N') or (rawinputList[16][x] == '不能进关')\
             or (rawinputList[16][x] == '不参与进关') or (rawinputList[16][x] == '不参与') or (rawinputList[16][x] == '否')\
             or (rawinputList[16][x] == '不涉及进关') or (rawinputList[16][x] == '不涉及')):
                outputList1[12][x-1] = 'N'
            else:
                print "行{}是否进关未找到".format(x+1)
        elif ((rawinputList[17][x] == '是') or (rawinputList[17][x] == 'Y') or (rawinputList[17][x] == '定制件')):
            outputList1[12][x-1] = 'C'
        elif ((rawinputList[17][x] == '否') or (rawinputList[17][x] == 'N') or (rawinputList[17][x] == '非定制件') or (rawinputList[17][x] == '')):
            if ((rawinputList[16][x] == '') or (rawinputList[16][x] == 'Y') or (rawinputList[16][x] == '是')\
             or (rawinputList[16][x] == '能进关') or (rawinputList[16][x] == '要进关') or (rawinputList[16][x] == '参与进关')\
              or (rawinputList[16][x] == '进关')):
                outputList1[12][x-1] = ''
            elif ((rawinputList[16][x] == '不进关') or (rawinputList[16][x] == 'N') or (rawinputList[16][x] == '不能进关')\
             or (rawinputList[16][x] == '不参与进关') or (rawinputList[16][x] == '不参与') or (rawinputList[16][x] == '否')\
             or (rawinputList[16][x] == '不涉及进关') or (rawinputList[16][x] == '不涉及')):
                outputList1[12][x-1] = 'N'
            else:
                print "行{}是否进关未找到".format(x+1)
        else:
            print "行{}是否进关未找到".format(x+1)

#应用领域

    for x in range(1,len(rawinputList[20])):
        if rawinputList[20][x].find('通讯') != -1:
            outputList1[13] += ['A01']
        elif rawinputList[20][x].find('宽带') != -1:
            outputList1[13] += ['A02']
        elif rawinputList[20][x].find('消费') != -1:
            outputList1[13] +=['A03']
        elif rawinputList[20][x].find('工业') != -1:
            outputList1[13] +=['A04']
        else :
            outputList1[13] +=['']
            print "行{}应用领域未找到".format(x+1)

#计量单位
    for x in outputList1[0]:
        if x == 'AVX' or x == 'NIC' or x == 'CHI':
            outputList1[17] += ['KP']
        elif x == 999:
            outputList1[17] += ['']
        else :
            outputList1[17] += ['PC']
#供应商 产品组 Segment Allocation默认N
    for x in cpxList:
        if x == 999:
            outputList1[18] += ['']
            outputList1[21] += ['']
            outputList1[24] += ['']
        else:
            outputList1[18] += [gysCode[x]]
            outputList1[21] += [cpzCode[x]]
            outputList1[24] += [segmentCode[x]]
        outputList1[22] += ['N']
#MPQ标记
    for x in outputList1[21]:
        if x == 'PAS':
            outputList1[20] += ['E']
        elif x == 'ACT':
            outputList1[20] += ['W']
        else :
            outputList1[20] += ['']

#监管条件
    for x in rawinputList[24][1:]:
        if x ==['']:
            outputList2[7] += ['']
        elif x.find('无') != -1:
            outputList2[7] += ['A4']
        else :
            outputList2[7] += [x]

#项目号
    if len(rawinputList[12]) == 0:
        outputList2[26] = ['000000' for x in outputList1[2]]
    else:
        rawinputList[12] = rawinputList[12][1:]
        for x in range(len(outputList1[2])):
            if x > len(rawinputList[12]):
                outputList2[26] += ['000000']
            elif rawinputList[12][x] == ('') or rawinputList[12][x] == (u''):
                outputList2[26] += ['000000']
            else:
                outputList2[26] += [rawinputList[12][x]]

#创建日期
    dayStr = now.strftime('%Y-%m-%d')
    outputList2[48] = [dayStr for x in outputList1[2]]


    outputList1[1] = outputList1[0]
    outputList2[1] = outputList1[2]


    workbook = xlsxwriter.Workbook(outputPath)
    worksheet = workbook.add_worksheet('上传')
    worksheet2 = workbook.add_worksheet('附加')
    orangeFormat = workbook.add_format()
    orangeFormat.set_bg_color('orange')

    blueFormat = workbook.add_format()
    blueFormat.set_bg_color('#66ccff')

    yellowFormat = workbook.add_format()
    yellowFormat.set_bg_color('yellow')

    worksheet.write('A1','总账级',orangeFormat)
    worksheet.write('B1','产品线',orangeFormat)
    worksheet.write('C1','料号',orangeFormat)
    worksheet.write('D1','子类别',orangeFormat)
    worksheet.write('E1','MPQ',orangeFormat)
    worksheet.write('F1','MOQ',orangeFormat)
    worksheet.write('G1','L/T',orangeFormat)
    worksheet.write('H1','License',orangeFormat)
    worksheet.write('I1','NCNR',orangeFormat)
    worksheet.write('J1','品牌',orangeFormat)
    worksheet.write('K1','产品类别',orangeFormat)
    worksheet.write('L1','BU')
    worksheet.write('M1','是否进关')
    worksheet.write('N1','应用领域',orangeFormat)
    worksheet.write('O1','最惠国税率',orangeFormat)
    worksheet.write('P1','进关品名',orangeFormat)
    worksheet.write('Q1','净重')
    worksheet.write('R1','计量单位',orangeFormat)
    worksheet.write('S1','默认供应商',orangeFormat)
    worksheet.write('T1','是否可以更新净重')
    worksheet.write('U1','是否控制MPQ',orangeFormat)
    worksheet.write('V1','产品组',orangeFormat)
    worksheet.write('W1','Allocation标记',orangeFormat)
    worksheet.write('X1','QSO标记')
    worksheet.write('Y1','Segment',orangeFormat)

    worksheet2.write('A1','短项目号')
    worksheet2.write('B1','料号')
    worksheet2.write('C1','性能及功能描述',blueFormat)
    worksheet2.write('D1','应用及领域',blueFormat)
    worksheet2.write('E1','中文品名',blueFormat)
    worksheet2.write('F1','尺寸',blueFormat)
    worksheet2.write('G1','普通关税',yellowFormat)
    worksheet2.write('H1','监管条件',yellowFormat)



#Format
    worksheet.set_column(2,2,18)
    worksheet.set_column(14,14,10.38)
    worksheet.set_column(15,15,27)
    worksheet.set_column(18,18,11)
    worksheet.set_column(19,19,17)
    worksheet.set_column(19,19,17)
    worksheet.set_column(22,22,15.25)

    worksheet2.set_column(0,0,10)
    worksheet2.set_column(1,1,18)
    worksheet2.set_column(2,2,17)
    worksheet2.set_column(3,4,10)

    for x in range(len(outputList1)):
        insertCol(worksheet,x,outputList1[x])
    for x in range(len(outputList2)):
        insertCol(worksheet2,x,outputList2[x])


    print "Running Successfully!"

if __name__=="__main__":
    main()
