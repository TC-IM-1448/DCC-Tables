import xml.etree.ElementTree as et
from xml.dom import minidom
import openpyxl as pyxl
from dcc2excel import xml2dcctable, xml2dccColumn

DCC='{https://ptb.de/dcc}'
SI='{https://ptb.de/si}'

def getRoot(xml):
    ## return the root element from an xml file
    et.register_namespace("si", SI.strip('{}'))
    et.register_namespace("dcc", DCC.strip('{}'))
    parser=et.XMLParser(encoding='utf-8')
    tree=et.parse(xml,parser)
    root=tree.getroot()
    return root

def getResultFromRoot(root, resId):
    #root: DCC xml-root element
    #resId: attribute value as string 
    #Returns: xml-result element
    for result in root.find(DCC+'measurementResults'):
        if result.attrib['resId']==resId:
            return result
    raise ValueError("No result found with the required id")
    return None

def getTableFromResult(result, itemRef="", settingRef=""):
   #INPUT xml-result element 
   #INPUT itemId and settingId list of strings
   #OUTPUT xml-element of type dcc:table

   xmltable=None
   count=0
   #Search all measurement results for a table with the required attributes
   searchResults=result.findall(DCC+"table")
   for table in searchResults:
       if all(Id in table.attrib['itemRef'] for Id in itemRef.split()) and all(Id in table.attrib['settingRef'] for Id in settingRef.split()) and count==0:
           xmltable=table
           count+=1

   if count==0:
       raise ValueError('Warning: DCC contains no tables with the required Id.')
   if count>1:
       raise ValueError('Warning: DCC contains ' + str(count) + ' tables with the required Id.\n Returning only the first instance')
   return xmltable


def getColumnFromTable(table,searchattributes, searchunit=""):
    #INPUT: xml-element of type dcc:table
    #INPUT: attribute dictionary
    #INPUT: searchunit as string.
    #OUTPUT: xml-element of type dcc:column
    for col in table.findall(DCC+'column'):
        unit=""
        if type(col.find(DCC+'unit')) !=type(None):
            unit=col.find(DCC+'unit').text
            print(unit)
        if col.attrib==searchattributes and searchunit==unit:
            print('found column')
            return col
    raise ValueError("No column found with the required attributes")
    return None

def getRowFromColumn(column, table, customerTag):

    try:
        tagcol=getColumnFromTable(table,{'scope':'dataInfo','dataCategory':'customerTag','measurand':'metaData'},'nan')
    except:
        raise RuntimeError("The table does not contain a customerTag column")

    """Iterate through the tags to find the row number of the specified tag"""
    tags=tagcol[2].text.split()
    found=False
    for i, tag in enumerate(tags):
        if tag==customerTag:
            found=True
            break
    if found:    
       searchValue=column[2].text.split()[i]
    else: 
       raise Exception("The requested customer tag was not found")
    return searchValue

def printelement(element):
    #INPUT xml-element
    #Print out the whole structure of an element to screen
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

"""
def getTableFromXML(xmlfile='example.xml', tableId='string'):
   #INPUT xml file
   #INPUT tableId string
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
           if table.attrib['tableId']==tableId and count==0:
               xmltable=table
               count+=1

   if count==0:
       print('Warning: DCC contains no tables with the required Id.')
   if count>1:
       print('Warning: DCC contains ' + str(count) + ' tables with the required Id.\n Returning only the first instance')
   return xmltable
"""
"""
def LookupColumn(xmlFile, tableID, itemID, scope, dataCategory, measurand, unit):
    """ """
    tbl = getTableFromXML(xmlfile=xmlFile, searchattributes={'itemId':itemID, 'refId':tableID})
    print(tbl)
    col = getColumnFromTable(table=tbl, searchattributes={'dataCategory':dataCategory,
                                                          'measurand':measurand,
                                                          'scope':scope}, searchunit = unit)
    return tbl, col

"""

"""
def lookupFromLookupListInFile(filename:str, dccFile:str):
    # filename = 'LookupList.csv'
    values = []
    outs = []
    with open(filename,"r") as f:
        print(f.readline())
        lines = f.readlines()
        for l in lines:
            temp = l.split(',')
            args = temp[:-1]
            unit = temp[-2]
            print(args)
            tbl, col = LookupColumn(dccFile, *args)
            dcccol = xml2dccColumn(col, unit)
            values.append(dcccol.columnData)
            outs.append(''.join(l[:-1])+','+','.join(dcccol.columnData)+'\n')

    with open(filename[:-4]+'Out.csv', "w") as f:
        f.write('TableID, itemID, scope, category, measurand, unit, value\n')
        f.writelines(outs)

"""

"""
def find_child(parent,child_type, search_key, search_value):
    #find children of a specified type, with a specified key (i.e. refType) set to 
    #a specified value
    for child in parent.iterfind(child_type):
        for key in child.attrib.keys():
            if key==search_key and child.attrib[key]==search_value:
                return child

"""


if __name__=="__main__" and 0 :
     #Sample userinput:
    xmlFile='mass_certificate.xml'
    tableAttrib={'itemId': 'item_ID1', 'refId': 'NN_temperature1'}
    columnAttrib={'dataCategory': 'Value', 'measurand': 'massConventional', 'scope': 'itemBias'}
    columnUnit="\mili\gram"

    #Lookup functions
    massTable=getTableFromXML(xmlfile=xmlFile,searchattributes=tableAttrib)
    col=getColumnFromTable(table=massTable, searchattributes=columnAttrib, searchunit=columnUnit)

    # tbl, col = LookupColumn('mass_certificate.xml', 'NN_temperature1', 'item_ID1', 'itemBias', 'Value', 'massConventional', '\mili\gram')
    # tbl, col = LookupColumn('certificate2.xml', 'TemperatureCalibration', 'item_ID1', 'reference', 'Value', 'temperatureAbsolute', '\degreecelcius')

    #Print result
    # if col:
        # printelement(col)

    # dcccol = xml2dccColumn(col, columnUnit)
    # dcccol.print()
    # dcctbl = xml2dcctable(tbl)

    # lookupFromLookupListInFile('massLookupList.csv','mass_certificate.xml')
    lookupFromLookupListInFile('cert2LookupList.csv','certificate2.xml')







