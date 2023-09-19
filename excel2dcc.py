import openpyxl as pyxl
import xml.etree.ElementTree as et
from xml.dom import minidom
import DCChelpfunctions as DCCh
#Used from DCChelpfunctions :
#item, add_name, add_identification, minimal, validate, DCC_tablecolumn,

DCC='{https://ptb.de/dcc}'
et.register_namespace("dcc", DCC.strip('{}'))
LANG='en'

def dictionaries_from_table(ws, rowtype):
    """
    Input: openpyxl worksheet object with column headers and row headers.
    Input: rowheader to look for (string)
    for each row of the given type make a dictionary with keys defined by the column headings and values defined in the cells of that row.
    Return the dictionaries as a list
    """
    columnheaders=ws['1']
    rowheaders=ws['A']
    dictionaries=[]
    rowno=0
    for rowheader in rowheaders:
        rowno+=1
        if rowheader.value==rowtype:
            dictionary={}
            for (name, content) in zip(columnheaders, ws[rowno]):
                dictionary[name.value]=content.value
            dictionaries.append(dictionary)
    return(dictionaries)

def read_statements_from_Excel(root, ws):
    adm=root.find(DCC+"administrativeData")
    statements=dictionaries_from_table(ws,'statement')
    statementselement=et.SubElement(adm,DCC+"statements")
    for statement in statements:
        statementelement=et.SubElement(statementselement,DCC+"statement", attrib={'statementId':statement['id']})
        et.SubElement(statementelement, DCC+"description", attrib={'lang':'en'}).text=statement['description']
        et.SubElement(statementelement, DCC+"description", attrib={'lang':'da'}).text=statement['description da']
        DCCh.add_name(statementelement,lang="en",text=statement['name en'])
        DCCh.add_name(statementelement,lang="da",text=statement['name da'])
    return root

def read_settings_from_Excel(root, ws):
    settings=dictionaries_from_table(ws,'setting')
    adm=root.find(DCC+"administrativeData")
    settingselement=et.SubElement(adm,DCC+"settings")
    for setting in settings:
        settingelement=et.SubElement(settingselement,DCC+"setting", attrib={'settingId':setting['id']})
        et.SubElement(settingelement, DCC+"description", attrib={'lang':'en'}).text=setting['description']
        et.SubElement(settingelement, DCC+"description", attrib={'lang':'da'}).text=setting['description da']
        DCCh.add_name(settingelement,lang="en",text=setting['name en'])
        DCCh.add_name(settingelement,lang="da",text=setting['name da'])
        if type(setting['value'])!=type(None):
           et.SubElement(settingelement, DCC+"value").text=str(setting['value'])
        if type(setting['unit'])!=type(None):
           et.SubElement(settingelement, DCC+"unit").text=setting['unit']
    return root

def read_item_from_Excel(root, ws):
    """
    Parameters
    ----------
    root : etree element
    ws : openpyxl worksheet object
        DESCRIPTION.

    Returns
    root element updated with items section
    -------
    None.
    """
    items=dictionaries_from_table(ws,'item')

    #Make an item XML-element
    Itemslist=root[0][2]
    for item in items:
       itemelement=DCCh.item(ID=item['id'], manufacturer=item['manufacturer'],model=item['productName'])
       DCCh.add_name(itemelement, 'en', item['description'])
       DCCh.add_identification(itemelement,item['customerId'],issuer='customer', name_dk="MålerID", name_en="SensorID")
       DCCh.add_identification(itemelement,item['serialNumber'],issuer='manufacturer',name_dk="Serienummer",name_en="Serial No.")
       Itemslist.append(itemelement)
    return root

def read_admin_from_Excel(root, ws):
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
            else:
                element=et.SubElement(element,DCC+level)
        element.text=values[i].value

    administrativeData=root.find(DCC+'administrativeData')
    accreditation=et.SubElement(administrativeData,DCC+'accreditation', attrib={'accrId':'accdfm'})
    et.SubElement(accreditation,DCC+'accreditationLabId').text="255"
    et.SubElement(accreditation,DCC+'accreditationBody').text="DANAK"
    et.SubElement(accreditation,DCC+'accreditationCountry').text="DK"
    et.SubElement(accreditation,DCC+'accreditationApplicability').text="2"

    return root

def read_tables_from_Excel(root, ws):
    return 0

def read_table_from_Excel(root, ws):

    """ TODO: Add function that finds all the tables in a given sheet """

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

    for c in content:
        col = DCCh.DccTableColumn(   scopeType=c[0],
                                columnType=c[1],
                                measurandType=c[2],
                                unit=c[3],
                                humanHeading = c[4],
                                columnData= list(map(str, c[5:])))
        columns.append(col)

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
    measurementResults=root.find(DCC+"measurementResults")
    measurementResult=et.SubElement(measurementResults,DCC+'measurementResult', attrib={'resId':'result1'})
    measurementResult.append(xmltable1)

    return root


def printelement(element):
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

if __name__ == "__main__":
    from importlib import reload
    reload(DCCh)

    workbookName="CalLab-DCC-writer.xlsx"
    # outputxml   ="certificate3.xml"
    schema      ="dcc.xsd"

    #load workbook
    wb = pyxl.load_workbook(workbookName, data_only=True)
    #Create root element with minimal content
    root = DCCh.minimal_DCC()

    #Update root element with content from worksheets in the workbook
    root = read_item_from_Excel(      root, ws=wb["Items"])
    root = read_admin_from_Excel(     root, ws=wb["AdministrativeData"])
    root = read_statements_from_Excel(root, ws=wb["Statements"])
    root = read_settings_from_Excel(root, ws=wb["Settings"])
    root = read_table_from_Excel(    root, ws=wb["Table2"])
    wb.close()

    ############### Output to xml-file ####################################
    outputxml = root.find(DCC+"administrativeData").find(DCC+"coreData").find(DCC+"uniqueIdentifier").text+".xml"
    xmlstr=minidom.parseString(et.tostring(root)).toprettyxml(indent="   ")
    with open(outputxml,'wb') as f:
        f.write(xmlstr.encode('utf-8'))

    #Validate the ouptut file against the schema and output the result.
    print(DCCh.validate(outputxml, schema))
