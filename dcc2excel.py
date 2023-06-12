import shutil
import openpyxl as pyxl
from xml.dom import minidom
import xml.etree.ElementTree as et

from DCChelpfunctions import DccTabel, DccTableColumn

DCC='{https://ptb.de/dcc}'
SI='{https://ptb.de/si}'
LANG='en'

colAttrDefs = ("scopeType", "columnType", "measurandType", "unit", "humanHeading")
tblAttrDefs =("tableID", "itemID", "numRows", "numColumns")


def getRoot(xml):
    ## return the root element from an xml file
    et.register_namespace("si", SI.strip('{}'))
    et.register_namespace("dcc", DCC.strip('{}'))
    parser=et.XMLParser(encoding='utf-8')
    tree=et.parse(xml,parser)
    root=tree.getroot()
    return root


def getTableFromXML(xmlfile='example.xml', searchattributes={'attr1':'value1', 'attr2':'value2'}):
   #INPUT xml file
   #INPUT attribute dictionary
   #OUTPUT xml-element of type dcc:table

   #Openthe xlm document and store it in a root xml elemnent
   root=getRoot(xmlfile)
   #Find all the measurement results in the measurementResults section
   measResults=root.find(DCC+'measurementResults').findall(DCC+'measurementResult')
   xmltable=None
   count=0
   #Search all measurement results for a table with the required attributes
   for measResult in measResults:
       results=measResult.find(DCC+"results")
       searchResults=results.findall(DCC+"table")
       for table in searchResults:
           if table.attrib==searchattributes and count==0:
               xmltable=table
               count+=1

   if count==0:
       print('Warning: DCC contains no tables with the required attributes.')
   if count>1:
       print('Warning: DCC contains ' + str(count) + ' tables with the required attributes.\n Returning only the first instance')
   return xmltable


def getColumnFromTable(table,searchattributes, searchunit=""):
    #INPUT: xml-element of type dcc:table
    #INPUT: attribute dictionary
    #INPUT: searchunit as string.
    #OUTPUT: xml-element of type dcc:column
    for col in table.findall(DCC+'column'):
        unit=""
        if type(col.find(SI+'unit')) !=type(None):
            unit=col.find(SI+'unit').text
        if col.attrib==searchattributes and searchunit==unit:
            return col
    print("No column found with the required attributes")
    return 0

def printelement(element):
    #INPUT xml-element
    #Print out the whole structure of an element to screen
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

def LookupColumn(xmlFile, tableID, itemID, scope, dataCategory, measurand, unit):
    """ """
    tbl = getTableFromXML(xmlfile=xmlFile, searchattributes={'itemId':itemID, 'refId':tableID})
    print(tbl)
    col = getColumnFromTable(table=tbl, searchattributes={'dataCategory':dataCategory,
                                                          'measurand':measurand,
                                                          'scope':scope}, searchunit = unit)
    return tbl, col




def write_DCC_table_to_excel_sheet(dccTbl: DccTabel, workbookName = ""):
    tbl = dccTbl
    shutil.copy("DCC-Table_empty_template.xlsx", workbookName)
    wb = pyxl.load_workbook(workbookName)
    ws = wb["TableTemplate"]
    newws = wb.copy_worksheet(ws)
    newws.title = dccTbl.tableID
    ws = wb[dccTbl.tableID]
    wb.active = wb[dccTbl.tableID]
    print(wb.sheetnames)

    ws["B2"] = tbl.tableID
    ws["B3"] = tbl.itemID
    ws["B4"] = tbl.numRows
    ws["B5"] = tbl.numColumns

    cell = ws["B6"]
    columns = tbl.columns

    for c in range(tbl.numColumns):
        for r in range(len(colAttrDefs)):
            cell.offset(r, c).value = getattr(columns[c], colAttrDefs[r])
        for r in range(tbl.numRows):
            cell.offset(r+5, c).value = columns[c].columnData[r]

    wb.save(workbookName)


def xml2dccColumn(col, unit: str):
    """


    Parameters
    ----------
    col : TYPE
        DESCRIPTION.
    unit : str
        DESCRIPTION.

    Returns
    -------
    None.

    """
    dcccol=DccTableColumn( scopeType=col.attrib['scope'], columnType=col.attrib['dataCategory'],
                          measurandType=col.attrib['measurand'], unit=unit,
                          humanHeading=col.find(DCC+'name').find(DCC+'content').text,
                          columnData=col.find(SI+'ValueXMLList').text.split())
    return dcccol


def xml2dcctable(xmltable):
    dcccolumns=[]
    for col in xmltable.findall(DCC+'column'):
        unit=""
        if type(col.find(SI+'unit')) !=type(None):
            unit=col.find(SI+'unit').text
        # dcccol=DccTableColumn( scopeType=col.attrib['scope'], columnType=col.attrib['dataCategory'], measurandType=col.attrib['measurand'], unit=unit, humanHeading=col.find(DCC+'name').find(DCC+'content').text, columnData=col.find(SI+'ValueXMLList').text.split())
        dcccol = xml2dccColumn(col, unit)
        dcccolumns.append(dcccol)
    length=len(col.find(SI+'ValueXMLList').text.split())
    dcctbl=DccTabel(xmltable.attrib['refId'],xmltable.attrib['itemId'],length,len(dcccolumns),dcccolumns)
    return dcctbl

if __name__ == "__main__":
    xmlFile='mass_certificate.xml'

    from validator import validate
    validate("certificate2.xml", "dcc.xsd")


    xmlFile='certificate2.xml'
    tableAttrib={'itemId': 'item_ID1', 'refId': 'NN_temperature1'}
    columnAttrib={'dataCategory': 'Value', 'measurand': 'massConventional', 'scope': 'itemBias'}
    columnUnit="\mili\gram"

    #Lookup functions
    xmlTable=getTableFromXML(xmlfile=xmlFile,searchattributes=tableAttrib)
    col=getColumnFromTable(table=xmlTable, searchattributes=columnAttrib, searchunit=columnUnit)

    tbl, col = LookupColumn('mass_certificate.xml', 'NN_temperature1', 'item_ID1', 'itemBias', 'Value', 'massConventional', '\mili\gram')

    #Print result
    if col:
        printelement(col)

    dcccol = xml2dccColumn(col, columnUnit)
    dcccol.print()
    dcctbl = xml2dcctable(tbl)

    write_DCC_table_to_excel_sheet(dcctbl, "DCC-Table_example_output.xlsx")





