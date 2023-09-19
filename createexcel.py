import openpyxl as pyxl
import xml.etree.ElementTree as et
from lookupFunctions import search
DCC='{https://ptb.de/dcc}'
WB=pyxl.Workbook()
WB.create_sheet('Table')
root=et.parse('DFM-T220000.xml')
xmlfile="DFM-T220000.xml"

attributes=[['scope'],['dataCategory'],['measurand'],['unit'],['humanHeading']]
tab=root.find(DCC+'measurementResults').find(DCC+'measurementResult').find(DCC+'table')

acols=[]
for col in tab.findall(DCC+'column'):
    attributes[0].append(col.attrib['scope'])
    attributes[1].append(col.attrib['dataCategory'])
    attributes[2].append(col.attrib['measurand'])
    attributes[3].append(col.find(DCC+'unit').text)
    attributes[4].append(col.find(DCC+'name').find(DCC+'content').text)
    col=search(root, tab.attrib,col.attrib,col.find(DCC+'unit').text)[0]
    cols.append(acol)

#n=len(attributes[0])
#tabAttrib={'tableId':'TempCal', 'settingRef':'setting6', 'itemRef':'itemID1 itemID2'}
#cols=[]
#keys= colAttrib=[a[0] for a in attributes][0:3]
#for i in range(1,n):
    #colAttribValues=[a[i] for a in attributes][0:3]
    #colAttrib=dict(zip(keys,colAttribValues))
    #unit=[a[i] for a in attributes][3]
    #col=search(root,tab.attrib,colAttrib,unit)[0]
    #cols.append(col)

ws=WB['Table']
for row in attributes:
    ws.append(row)
for n in range(0,len(cols[0])):
    r=['']+[c[n] for c in cols]
    ws.append(r) 
WB.save("test.xlsx")
        
