import openpyxl as pyxl
import xml.etree.ElementTree as et
from DCChelpfunctions import search

"""
TODO : make main function with input /outpu arguments
"""
if __name__=="__main__":
    import sys
    args=sys.argv[1:]
    print(len(args))
    if len(args)==0:
        xmlfile="DFM-T220000.xml"
        print("0 arg")
    else:
        xmlfile=args[0]
    if len(args)==2:
        WB=pyxl.load_workbook(args[1])
        print("2 arg")
    else:    
        WB=pyxl.Workbook()
    WB.create_sheet('Table')
   
    DCC='{https://dfm.dk}'

    root=et.parse(xmlfile)
    
    
    attributes=[['scope'],['dataCategory'],['measurand'],['unit'],['metaDataCategory'],['humanHeading']]
    tab=root.find(DCC+'measurementResults').find(DCC+'measurementResult').find(DCC+'table')
    
    cols=[]
    for col in tab.findall(DCC+'column'):
        attributes[0].append(col.attrib['scope'])
        attributes[1].append(col.attrib['dataCategory'])
        attributes[2].append(col.attrib['measurand'])
        attributes[3].append(col.find(DCC+'unit').text)
        attributes[4].append(col.attrib['metaDataCategory'])
        attributes[5].append(col.find(DCC+'name').find(DCC+'content').text)
        col=search(root, tab.attrib,col.attrib,col.find(DCC+'unit').text)[0]
        if col=='-':
            col=['']*len(cols[0])
        cols.append(col)
    
    
    ws=WB['Table']
    ws.append(["DCCTable"])
    for item in tab.attrib.items():
        ws.append([item[0],item[1]])
    
    ws.append(['numRows',len(cols[0])])
    ws.append(['numColumns',len(cols)])
    
    for row in attributes:
        ws.append(row)
    for n in range(0,len(cols[0])):
        r=['']+[c[n] for c in cols]
        ws.append(r) 
    WB.save("test.xlsx")
         
