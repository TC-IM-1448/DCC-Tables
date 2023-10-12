import os
import openpyxl as pyxl
import xml.etree.ElementTree as etree
import DCChelpfunctions as dcchf
from DCChelpfunctions import search

LANG='da'
DCC='{https://dfm.dk}'




def lookupFromMappingFile(mapFileName:str, dccFileName:str):
    # filename = 'LookupList.csv'
    root = etree.parse(dccFileName)

    values = []
    outs = []
    query_types = []

    wb = pyxl.load_workbook(mapFileName)
    sheet = wb['Mapping']
    colA = list(sheet.columns)[1]
    n_rows = sheet.max_row
    n_cols = sheet.max_column
    cols = sheet.columns
    rows = sheet.rows

    for i in range(2, n_rows-1):
        cellA = sheet.cell(row=i, column=1).value
        if cellA == "--END--": break
        
        queryType = str(sheet.cell(row=i, column=3).value)
        print(queryType, end = "   ")

        if queryType == 'xpath': 
            cellC = sheet.cell(row=i, column=4).value
            xpath = cellC
            s = xpath.split("/dcc:digitalCalibrationCertificate")[1]
            ss = s.replace("dcc:", DCC)
            # print(ss, end= "    ")
            elm = root.find(ss)
            # print(elm)
            elm = elm.text
            print(elm)
            cell = sheet.cell(row=i, column=n_cols+1, value=elm)
        elif queryType == 'data':
            tableId = sheet.cell(row=i, column=5).value
            itemRef = sheet.cell(row=i, column=6).value
            settingRef = sheet.cell(row=i, column=7).value
            scope = sheet.cell(row=i, column=8).value
            dataCategory = sheet.cell(row=i, column=9).value
            measurand = sheet.cell(row=i, column=10).value
            metaDataCategory = sheet.cell(row=i, column=11).value
            unit = sheet.cell(row=i, column=12).value
            customerTag = sheet.cell(row=i, column=13).value
            data = search(root,
                          {'tableId': tableId, 
                            'itemRef': itemRef, 
                            'settingRef': settingRef, 
                            }, 
                           {'scope': scope, 
                            'dataCategory': dataCategory, 
                            'measurand': measurand, 
                            'metaDataCategory': metaDataCategory
                            }, 
                            unit = unit, 
                            customerTag = customerTag
                            )[0]
            print(data)
            cell = sheet.cell(row=i, column=n_cols+1, value=data)
        else: 
            cell = sheet.cell(row=i, column=n_cols+1, value="FAILED")


    outFileName = mapFileName.rsplit(".",maxsplit=1)[0]+'_QueryResult.xlsx'
    wb.save(outFileName)


    # with open(filename,"r") as f:
    #     print(f.readline())
    #     lines = f.readlines()
    #     for l in lines:
    #         temp = l.split(',')
    #         args = temp[:-1]
    #         unit = temp[-2]
    #         print(args)
    #         tbl, col = LookupColumn(dccFile, *args)
    #         dcccol = xml2dccColumn(col, unit)
    #         values.append(dcccol.columnData)
    #         outs.append(''.join(l[:-1])+','+','.join(dcccol.columnData)+'\n')

    # with open(filename[:-4]+'Out.csv', "w") as f:
    #     f.write('TableID, itemID, scope, category, measurand, unit, value\n')
    #     f.writelines(outs)


if __name__=="__main__":
    #################
    #first argument is dcc xml file
    #second argument is excel template to use

    # import sys
    # args=sys.argv[1:]
    # print(len(args))
    # if len(args)==0:
    #     xmlfile="DFM-T220000.xml"
    # else:
    #     xmlfile=args[0]
    # if len(args)==2:
    #     WB=pyxl.load_workbook(args[1])
    # else:    
    #     mapfile = 'Examples/Mapping_Novo_temperatur_Certifikat.xlsx'
   
    # DCC='{https://dfm.dk}'

    lookupFromMappingFile('Examples'+os.sep+'Mapping_Novo_temperatur_Certifikat.xlsx', 'Examples'+os.sep+'Stip-230063-V1.xml')
         
