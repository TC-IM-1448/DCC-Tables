import openpyxl as pyxl
import xml.etree.ElementTree as et
from DCChelpfunctions import search
from excel2dcc import printelement

LANG='da'

def write_row(write_sheet, row_num: int, starting_column: str or int, write_values: list):
    if isinstance(starting_column, str):
        starting_column = ord(starting_column.lower()) - 96
    for i, value in enumerate(write_values):
        write_sheet.cell(row_num, starting_column + i, value)

def statements(sheetname, statementselement):
    #WB.create_sheet(sheetname)
    line=1
    ws=WB[sheetname]
    columnheadings=['category','id', 'heading','body']
    write_row(ws,line,2,columnheadings)
    line+=1
    headingstr=DCC+"heading[@lang='"+LANG+"']"
    bodystr=DCC+"body[@lang='"+LANG+"']"
    for statement in statementselement.findall(DCC+'statement'):
        [category,ID,heading,body]=["","","",""]
        category=statement.find(DCC+'category').text
        ID=statement.attrib['statementId']
        heading=statement.find(headingstr).text
        body=statement.find(bodystr).text
        write_row(ws,line,2,[category,ID,heading,body])
        line+=1



if __name__=="__main__":
    #################
    #first argument is dcc xml file
    #second argument is excel template to use
    import sys
    args=sys.argv[1:]
    print(len(args))
    if len(args)==0:
        xmlfile="Examples/DFM-T220000.xml"
    else:
        xmlfile=args[0]
    if len(args)==2:
        WB=pyxl.load_workbook(args[1])
    else:    
        WB=pyxl.Workbook()
        WB.create_sheet('Table')
        WB.create_sheet('statements')
   
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
        col=search(root, tab.attrib,col.attrib,col.find(DCC+'unit').text)[0][0][2].text.split()
        if col=='-':
            col=['']*len(cols[0])
        cols.append(col)
    
    
    ws=WB['Table']
    line=1
    write_row(ws,line,1,["DCCTable"])
    line+=1
    for item in tab.attrib.items():
        write_row(ws,line,1,[item[0],item[1]])
        line+=1
    
    write_row(ws,line,1,['numRows',len(cols[0])])
    line+=1
    write_row(ws,line,1,['numColumns',len(cols)])
    line+=3
    
    for row in attributes:
        write_row(ws,line,1,row)
        line+=1
        #ws.append(row)
    for n in range(0,len(cols[0])):
        r=['']+[c[n] for c in cols]
        write_row(ws,line,1,r)
        line+=1
        #ws.append(r) 

    stat=root.find(DCC+'administrativeData').find(DCC+'statements')
    statements('statements',stat)
    WB.save("view_content.xlsx")
         
