import os
import sys
import shutil
import openpyxl as pyxl
import xml.etree.ElementTree as et
from xml.dom import minidom
#import pydcc_tables as pydcc
#from pydcc_tables import DccTableColumn, DccTabel
# from pydcc_tables import DccTabel, DccTableColumn

from DCChelpfunctions import DccTableColumn, DccTabel
import DCChelpfunctions as DCCh


def read_statements_from_Excel(root, workbookName="DCC-Table_example3.xlsx",sheetName="Statements"):
    wb = pyxl.load_workbook(workbookName, data_only=True)
    ws = wb[sheetName]
    linetypes=ws['A']
    columntypes=ws['1']
    statements=[]
    lineno=0
    for linetype in linetypes:
        lineno+=1
        if linetype.value=="statement":
           statement={}
           for (name, content) in zip(columntypes, ws[lineno]):
               statement[name.value]=content.value
           statements.append(statement)
    adm=root.find(DCC+"administrativeData")
    statementselement=et.SubElement(adm,DCC+"statements")
    for statement in statements:
        statementelement=et.SubElement(statementselement,DCC+"statement", attrib={'id':statement['id']})
        et.SubElement(statementelement, DCC+"description", attrib={'lang':'en'}).text=statement['description']
        et.SubElement(statementelement, DCC+"description", attrib={'lang':'da'}).text=statement['description da']
        DCCh.add_name(statementelement,lang="en",text=statement['name en'])
        DCCh.add_name(statementelement,lang="da",text=statement['name da'])

    return root
           


def read_item_from_Excel(workbookName="DCC-Table_example3.xlsx",sheetName="Items"):
    wb = pyxl.load_workbook(workbookName, data_only=True)

    ws = wb[sheetName]
    item={}
    item['id']=ws['A2'].value
    item['custromerId']=ws['B2'].value
    item['equipmentClass']=ws['C2'].value
    item['description']=ws['D2'].value
    item['swRef']=ws['E2'].value
    item['manufacturer']=ws['F2'].value
    item['productName']=ws['G2'].value
    item['productNumber']=ws['E2'].value
    item['serialNumber']=ws['F2'].value
    return item

def read_admin_from_Excel(root, workbookName="DCC-Table_example3.xlsx",sheetName="AdministrativeData"):
    wb = pyxl.load_workbook(workbookName, data_only=True)
    ws = wb[sheetName]
    DFM_names=ws['B']
    values=ws['C']
    xpaths=ws['D']
    for i, path in enumerate(xpaths):
        try:
            levels=path.value.split("/dcc:")[2:]
        except:
            levels=[]
        element=root
        for level in levels:
            subelement=element.find(DCC+level)
            if type(subelement)!=type(None):
                element=subelement
                #print("1")
                #print(element)
            else:
                element=et.SubElement(element,DCC+level)
                #print("2")
                #print(element)
        element.text=values[i].value

    administrativeData=root.find(DCC+'administrativeData')
    accreditation=et.SubElement(administrativeData,DCC+'accreditation', attrib={'accrId':'accdfm'})
    et.SubElement(accreditation,DCC+'accreditationLabId').text="255"
    et.SubElement(accreditation,DCC+'accreditationBody').text="DANAK"
    et.SubElement(accreditation,DCC+'accreditationCountry').text="DK"
    et.SubElement(accreditation,DCC+'accreditationApplicability').text="2"

    return root

    
def read_tables_from_Excel(workbookName="DCC-Table_example3.xlsx",sheetName="Table2"):
    """ Function that finds all the tables in a given sheet """

    wb = pyxl.load_workbook(workbookName, data_only=True)
    ws = wb[sheetName]
    columns = []
    attrib={}
    attribnames=ws['A'][1:5]
    attribvalues=ws['B'][1:5]
    for (name, value) in zip(attribnames,attribvalues):
        if type(name.value) != type(None) and type(value.value) != type(None):
            attrib[name.value]=value.value
    statementRef = ws["B5"].value
    numRows = ws["B7"].value
    numColumns = ws["B8"].value

    nRows = int(numRows)+5
    nCols = int(numColumns)

    cell = ws["B9"]

    content = [[cell.offset(r,c).value for r in range(nRows)] for c in range(nCols)]
    # content = transpose_2d_list(content)

    for c in content:
        col = DccTableColumn(   scopeType=c[0],
                                columnType=c[1],
                                measurandType=c[2],
                                unit=c[3],
                                humanHeading = c[4],
                                columnData= list(map(str, c[5:])))
        columns.append(col)

    #tbl = DccTabel(tableID, itemID, settingRef, numRows, numColumns, columns)
    wb.close()
    #Create empty table with table attributes
    xmltable1=et.Element(DCC+"table",attrib=attrib)

    #Fill the table with data from table object
    columns=columns
    for col in columns:
        attributes={'scope':col.scopeType, 'dataCategory':col.columnType, 'measurand':col.measurandType}
        xmlcol=et.Element(DCC+'column',attrib=attributes)
        if type(col.unit)!=type(None):
          et.SubElement(xmlcol,DCC+'unit').text=' '.join([col.unit])
        DCCh.add_name(xmlcol,lang="en",text=col.humanHeading)
        #xmllist=realListXMLList(value=col.columnData,unit=[col.unit])
        if attributes['dataCategory']=='Conformity':
            et.SubElement(xmlcol,DCC+"conformityXMLList").text=' '.join(col.columnData)
        elif attributes['dataCategory']=='customerTag':
            et.SubElement(xmlcol,DCC+"stringXMLList").text=' '.join(col.columnData)
        elif attributes['dataCategory']=='Exception':
            et.SubElement(xmlcol,DCC+"exceptionXMLList").text=' '.join(col.columnData)
        elif attributes['dataCategory']=='accreditationApplies':
            et.SubElement(xmlcol,DCC+"accreditationAppliesXMLList",attrib={'accrRef':'accdfm'}).text=' '.join(col.columnData)
        else:
            et.SubElement(xmlcol,DCC+"valueXMLList").text=' '.join(col.columnData)
        xmltable1.append(xmlcol)
    return xmltable1
#from docx import Document

colAttrDefs = ("scopeType", "columnType", "measurandType", "unit", "humanHeading")
tblAttrDefs =("tableID", "itemID", "numRows", "numColumns")

#################### Make a minimal DCC and fill in the administrative data ##############################
DCC='{https://ptb.de/dcc}'
SI='{https://ptb.de/si}'
et.register_namespace("si", SI.strip('{}'))
et.register_namespace("dcc", DCC.strip('{}'))
LANG='en'

def add_item_data(root,inputItem):
    """
    Parameters
    ----------
    root : etree element
        DESCRIPTION.

    Returns
    root element updated with items section
    -------
    None.
    """

    ItemID=inputItem['id']
    Manufacturer=inputItem['manufacturer']
    Model=inputItem['productName']
    customerID=inputItem['custromerId']
    SerialNo=inputItem['serialNumber']
    Description=inputItem['description']
    #item['equipmentClass']=ws['C2']
    #item['swRef']=ws['E2']
    #item['productNumber']=ws['E2']

    #Make an item XML-element
    item=DCCh.item(ID=ItemID, manufacturer=Manufacturer,model=Model)
    DCCh.add_name(item, 'en', Description)
    DCCh.add_identification(item,customerID,issuer='customer', name_dk="MålerID", name_en="SensorID")
    DCCh.add_identification(item,SerialNo,issuer='manufacturer',name_dk="Serienummer",name_en="Serial No.")

    Items=root[0][2]
    Items.append(item)
    return root

def insertTable2Xml(root, xmltable):
    #Create empty table with table attributes
    #xmltable1=et.Element(DCC+"table",attrib={'itemRef':tab1.itemID,'tableId':tab1.tableID, 'statementRef':tab1.statementRef})

    #Fill the table with data from table object
    #columns=tab1.columns
    #for col in columns:
        #attributes={'scope':col.scopeType, 'dataCategory':col.columnType, 'measurand':col.measurandType}
        #xmlcol=et.Element(DCC+'column',attrib=attributes)
        #if type(col.unit)!=type(None):
          #et.SubElement(xmlcol,DCC+'unit').text=' '.join([col.unit])
        #DCCh.add_name(xmlcol,lang="en",text=col.humanHeading)
        ##xmllist=realListXMLList(value=col.columnData,unit=[col.unit])
        #if attributes['dataCategory']=='Conformity':
            #et.SubElement(xmlcol,DCC+"conformityXMLList").text=' '.join(col.columnData)
        #elif attributes['dataCategory']=='customerTag':
            #et.SubElement(xmlcol,DCC+"stringXMLList").text=' '.join(col.columnData)
        #elif attributes['dataCategory']=='accreditationApplies':
            #et.SubElement(xmlcol,DCC+"accreditationAppliesXMLList",attrib={'accrRef':'accdfm'}).text=' '.join(col.columnData)
        #else:
            #et.SubElement(xmlcol,DCC+"valueXMLList").text=' '.join(col.columnData)
        #xmltable1.append(xmlcol)

    #Append table to results section of the xml
    xmlresults=root[1][0][1]
    xmlresults.append(xmltable)

    ################# END add calibration data #####################################


def printelement(element):
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

if __name__ == "__main__":
    from importlib import reload
    reload(DCCh)
    examplefile="DCC-Table_example3.xlsx"
    #examplefile="DCC-mass_example.xlsx"

    root=DCCh.minimal_DCC()
    inputItem=read_item_from_Excel(workbookName=examplefile,sheetName="Items")
    #inputAdm=read_admin_from_Excel(workbookName=examplefile,sheetName="AdministrativeData")
    root=read_admin_from_Excel(root, workbookName=examplefile,sheetName="AdministrativeData")
    root = add_item_data(root, inputItem)
    ######################### Add table with calibration data to the xml ##########################
    tbl = read_tables_from_Excel(workbookName=examplefile,sheetName="Table2")

    print(inputItem)
    insertTable2Xml(root,tbl)

    #Print the tbl and column 5
    #columns = tbl.columns
    #columns[5].print()
    #for i in range(tbl.numColumns):
        #print(columns[i].columnData)

    ############### Output to xml-file ####################################

    #FIXME: with the namespace representation ns:elementname the printelement function does not work
    xmlstr=minidom.parseString(et.tostring(root)).toprettyxml(indent="   ")
    with open('certificate2.xml','wb') as f:
        f.write(xmlstr.encode('utf-8'))
    ############### END Output to xml-file ####################################

    print(DCCh.validate("certificate2.xml", "dcc.xsd"))
