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

#from docx import Document

colAttrDefs = ("scopeType", "columnType", "measurandType", "unit", "humanHeading")
tblAttrDefs =("tableID", "itemID", "numRows", "numColumns")

#################### Make a minimal DCC and fill in the administrative data ##############################
DCC='{https://ptb.de/dcc}'
SI='{https://ptb.de/si}'
et.register_namespace("si", SI.strip('{}'))
et.register_namespace("dcc", DCC.strip('{}'))
LANG='en'

def add_administrative_data(root):
    """
    Temporary function for adding administrative data

    Parameters
    ----------
    root : etree element
        DESCRIPTION.

    Returns
    -------
    None.

    """
    administrativeData=root.find(DCC+'administrativeData')
    coreData=administrativeData.find(DCC+'coreData')
    coreData.find(DCC+'uniqueIdentifier').text='T2304'
    DCCh.add_identification(coreData,value="NN42",issuer='calibrationLaboratory', name_dk="kundenr",name_en="customer ID")
    DCCh.add_identification(coreData,value="jpx2340988",issuer="customer",name_dk="PO",name_en="PO")
    coreData.find(DCC+'receiptDate').text="2022-08-13"
    coreData.find(DCC+'beginPerformanceDate').text="2022-08-14"
    coreData.find(DCC+'endPerformanceDate').text="2022-08-15"

    lab=administrativeData.find(DCC+'calibrationLaboratory')
    contact=et.SubElement(lab, DCC+'contact')
    DCCh.fill_address(contact,name="DFM", eMail="srk@dfm.dk", phone="+45 2545 9040", city="Hørsholm",  postCode="2970", street="Kogle Allé", streetNo="5", further="www.dfm.dk")

    respPersons=administrativeData.find(DCC+'respPersons')
    DCCh.add_respPerson(respPersons,name="Erling Målermand", mainSigner=True)
    DCCh.add_respPerson(respPersons,name="Simon  Hansen", mainSigner=False)

    customer=administrativeData.find(DCC+'customer')
    DCCh.fill_address(customer,name="NN", eMail="pqrt@nn.com", phone="+45 6160 7019", city="Søborg", postCode="2860", street="Svanevej", streetNo="12", further="kundenummer: 1234")

    accreditation=et.SubElement(administrativeData,DCC+'accreditation', attrib={'accrId':'accdfm'})
    et.SubElement(accreditation,DCC+'accreditationLabId').text="255"
    et.SubElement(accreditation,DCC+'accreditationBody').text="DANAK"
    et.SubElement(accreditation,DCC+'accreditationCountry').text="DK"
    et.SubElement(accreditation,DCC+'accreditationApplicability').text="2"


    ################ User input for item ##########################
    ItemID="itemID1"
    Manufacturer='Mettler-Toledo'
    Model='Platinum Super'
    customerID="NN66"
    SerialNo="2341-LKJQ-1324LKLJJAAFLKK33"

    #Make an item XML-element
    item=DCCh.item(ID=ItemID, manufacturer=Manufacturer,model=Model)
    DCCh.add_name(item, 'en', 'Set of 7 weights')
    DCCh.add_identification(item,customerID,issuer='customer', name_dk="MålerID", name_en="SensorID")
    DCCh.add_identification(item,SerialNo,issuer='manufacturer',name_dk="Serienummer",name_en="Serial No.")

    Items=root[0][2]
    Items.append(item)
    return root
    ######################### END of administrative data ####################################################

def insertTable2Xml(root, tab1):
    #Create empty table with table attributes
    xmltable1=et.Element(DCC+"table",attrib={'itemRef':tab1.itemID,'tableId':tab1.tableID})

    #Fill the table with data from table object
    columns=tab1.columns
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
        elif attributes['dataCategory']=='accreditationApplies':
            et.SubElement(xmlcol,DCC+"stringXMLList").text=' '.join(col.columnData)
            #NOTE: should be accreditationAppliesXMLList. (the type needs fix in the dcc.xsd-schema)
        else:
            et.SubElement(xmlcol,DCC+"valueXMLList").text=' '.join(col.columnData)
        xmltable1.append(xmlcol)

    #Append table to results section of the xml
    xmlresults=root[1][0][1]
    xmlresults.append(xmltable1)

    ################# END add calibration data #####################################


def printelement(element):
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

def read_tables_from_Excel(workbookName="DCC-Table_example3.xlsx",sheetName="Table2"):
    """ Function that finds all the tables in a given sheet """

    wb = pyxl.load_workbook(workbookName, data_only=True)

    ws = wb[sheetName]


    columns = []

    tableID = ws["B2"].value
    itemID = ws["B3"].value
    numRows = ws["B4"].value
    numColumns = ws["B5"].value

    nRows = int(numRows)+5
    nCols = int(numColumns)

    cell = ws["B6"]

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

    tbl = DccTabel(tableID, itemID, numRows, numColumns, columns)
    wb.close()
    return tbl

if __name__ == "__main__":
    from importlib import reload
    reload(DCCh)

    root=DCCh.minimal_DCC()
    root = add_administrative_data(root)
    ######################### Add table with calibration data to the xml ##########################
    tbl = read_tables_from_Excel(workbookName="DCC-Table_example3.xlsx",sheetName="Table2")
    insertTable2Xml(root,tbl)

    #Print the tbl and column 5
    columns = tbl.columns
    columns[5].print()
    for i in range(tbl.numColumns):
        print(columns[i].columnData)

    ############### Output to xml-file ####################################

    #FIXME: with the namespace reprecentation ns:elementname the printelement function does not work
    xmlstr=minidom.parseString(et.tostring(root)).toprettyxml(indent="   ")
    with open('certificate2.xml','wb') as f:
        f.write(xmlstr.encode('utf-8'))
    ############### END Output to xml-file ####################################

    print(DCCh.validate("certificate2.xml", "dcc.xsd"))
