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




if __name__=="__main__":
     #Sample userinput:
    xmlFile='mass_certificate.xml'
    tableAttrib={'itemId': 'item_ID1', 'refId': 'NN_temperature1'}
    columnAttrib={'dataCategory': 'Value', 'measurand': 'massConventional', 'scope': 'itemBias'}
    columnUnit="\mili\gram"

    #Lookup functions
    massTable=getTableFromXML(xmlfile=xmlFile,searchattributes=tableAttrib)
    col=getColumnFromTable(table=massTable, searchattributes=columnAttrib, searchunit=columnUnit)

    tbl, col = LookupColumn('mass_certificate.xml', 'NN_temperature1', 'item_ID1', 'itemBias', 'Value', 'massConventional', '\mili\gram')
    tbl, col = LookupColumn('certificate2.xml', 'NN_temperature1', 'item_ID1', 'reference', 'Value', 'temperatureAbsolute', '\degreecelcius')

    #Print result
    if col:
        printelement(col)

    dcccol = xml2dccColumn(col, columnUnit)
    dcccol.print()
    dcctbl = xml2dcctable(tbl)

    lookupFromLookupListInFile('massLookupList.csv','mass_certificate.xml')
    lookupFromLookupListInFile('cert2LookupList.csv','certificate2.xml')







