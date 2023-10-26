from lxml import etree
from urllib.request import urlopen
import xml.etree.ElementTree as et
from xml.dom import minidom
import openpyxl as pyxl
import openpyxl as pyxl

DCC='{https://dfm.dk}'
SI='{https://ptb.de/si}'
LANG='en'
et.register_namespace("si", SI.strip('{}'))
et.register_namespace("dcc", DCC.strip('{}'))

class DccTableColumn():
    """ """
    scopeType = ""
    columnType = ""
    measurandType = ""
    unit = ""
    metaDataCategory=""
    humanHeading = {}
    columnData = []

    def __init__(self,
                scopeType="",
                columnType="",
                measurandType="",
                unit="",
                metaDataCategory="",
                humanHeading = {},
                columnData=[]):
        """ """
        self.columnType = columnType
        self.scopeType = scopeType
        self.measurandType = measurandType
        self.unit = unit
        self.metaDataCategory = metaDataCategory
        self.humanHeading = {}
        for i, hh in enumerate(humanHeading):
            self.humanHeading['lang'+str(i+1)] = hh
        self.columnData = columnData

    def get_attributes(self):
        attr = [attr for attr in dir(self) if not callable(getattr(self, attr)) and not attr.startswith("__")]
        # local_attributes = {k: v for k, v in vars(self).items() if k in locals()}
        return attr

    def print(self):
        attr = self.get_attributes()
        str = ""
        for a in attr:
            print(a, ": \t", getattr(self, a))

class DccTabel():
    """ """
    tableID = ""
    itemID = ""
    numRows = ""
    numColumns = ""
    columns = []

    def __init__(self, tableID="", itemID="",
                 numRows = "", numColumns = "",columns=""):
        self.tableID = tableID
        self.itemID = itemID
        self.numRows = numRows
        self.numColumns = numColumns
        self.columns = columns

def validate(xml_path: str, xsd_path: str) -> bool:
    if xsd_path[0:5]=="https":
    #Note: etree.parse can not handle https, so we have to open the url with urlopen
       with urlopen(xsd_path) as xsd_file:
          xmlschema_doc = etree.parse(xsd_file)
    else:
       xmlschema_doc = etree.parse(xsd_path)

    xmlschema = etree.XMLSchema(xmlschema_doc)

    xml_doc = etree.parse(xml_path)
    result = xmlschema.validate(xml_doc)
    print(result)

    return xmlschema.error_log.filter_from_errors()

def transpose_2d_list(matrix):
    return [list(row) for row in zip(*matrix)]

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
        #if result.attrib['resId']==resId:
            return result
    raise ValueError("No result found with the required id")
    return None

def getTableFromResult(result, tableAttrib):
   #INPUT xml-result element 
   #INPUT itemId and settingId list of strings
   #OUTPUT xml-element of type dcc:table

   xmltable=None
   count=0
   #Search all measurement results for a table with the required attributes
   searchResults=result.findall(DCC+"table")
   for table in searchResults:
       if all(Id in table.attrib['itemRef'] for Id in tableAttrib['itemRef'].split()) and all(Id in table.attrib['settingRef'] for Id in tableAttrib['settingRef'].split()) and table.attrib['tableId']==tableAttrib['tableId'] and count==0:
           xmltable=table
           count+=1
   if count==0:
       raise ValueError('Warning: DCC contains no tables with the required combination of setting and item Ids.')
   if count>1:
       raise ValueError('Warning: DCC contains ' + str(count) + ' tables with the required Id.\n Returning only the first instance')
   return xmltable

def match_attributes(att,searchatt, unit, searchunit):
    for key in att.keys():
        if att[key]!='-' and searchatt[key]!='*' and att[key]!=searchatt[key]:
            return False
    if unit!='-' and searchunit!='*' and unit!=searchunit:
        return False
    return True


def getColumnFromTable(table,searchattributes, searchunit=""):
    #INPUT: xml-element of type dcc:table
    #INPUT: attribute dictionary
    #INPUT: searchunit as string.
    #OUTPUT: xml-element of type dcc:column
    cols=[]
    for col in table.findall(DCC+'column'):
        unit=""
        if type(col.find(DCC+'unit')) !=type(None):
            unit=col.find(DCC+'unit').text
        #if col.attrib==searchattributes and searchunit==unit:
        if match_attributes(col.attrib, searchattributes,unit,searchunit):
            cols.append(col)
            #return col
    if len(cols)==0: 
        raise ValueError("No column found with the required attributes")
        return None
    return cols

def getRowFromColumn(column, table, customerTag):

    try:
        tagcol=getColumnFromTable(table,{'scope':'*','dataCategory':'*','measurand':'*','metaDataCategory':'customerTag'},'*')
    except:
        raise RuntimeError("The table does not contain a customerTag column")

    """Iterate through the tags to find the row number of the specified tag"""
    tags=tagcol[0][-1].text.split()
    found=False
    for i, tag in enumerate(tags):
        if tag==customerTag:
            found=True
            break
    if found:    
       searchValue=column[-1].text.split()[i]
       return searchValue
    else: 
       raise Exception("The requested customer tag was not found")
       return None

def search(root, tableAttrib, colAttrib, unit, customerTag=None, lang="en"):
   """
   INPUT: 
   root: etree root element 
   tableAttributes itemRef, settingRef and tableId as dictionary of string values
   coAttributes scope, dataCategory and measurand  as dictionary of string values
   unit as string
   customerTag (optional)  as string
   OUTPUT:
   search result as string (or list of strings if customerTag is not specified)
   warnings as strings 
   """

   searchValue="-"
   warning="-"
   usertagwarning="-"
   colwarning="-"
   cols=[]

   try:
       """Find the right result using resId"""
       res=getResultFromRoot(root, resId="")
       try:
           """Find the right table using itemRef and settingRef"""
           tab=getTableFromResult(res, tableAttrib)
           try:
               """Find the rigt column using attributes and unit"""
               cols=getColumnFromTable(tab,colAttrib,unit)
               try:
                   if type(customerTag)!=type(None):
                       searchValue=[]
                       for col in cols:
                           searchValue.append(getRowFromColumn(col,tab,customerTag))
                   else:
                       #searchValue=col[2].text.split()
                       searchValue=cols
               except Exception as e:
                   usertagwarning=e.args[0]
           except Exception as e:
               colwarning=e.args[0]
       except Exception as e:
           warning=e.args[0]
   except Exception as e:
       warning=e.args[0]

   for col in cols:
       print("")
       selectstring="heading[@lang='"+lang+"']"
       heading=col.find(DCC+selectstring)
       if type(heading) != type(None):
           print(heading.text)
       #print(col.attrib)
       print("Unit:" +col.find(DCC+'unit').text)
       if type(customerTag)!=type(None):
            print(getRowFromColumn(col,tab,customerTag))
       else:
            print(col[-1].text.split())

   #return [searchValue, usertagwarning, colwarning, warning]
   return searchValue

def get_statement(root, ID, lang="en"):
    statements=root.find(DCC+"administrativeData").find(DCC+"statements")
    returnstatement=None
    for statement in statements:
        if ID==statement.attrib['statementId'] or ID=='*':
            for heading in statement.findall(DCC+"heading"):
                if heading.attrib['lang']==lang:
                    print("---Header-----")
                    print(heading.text)
            for body in statement.findall(DCC+"body"):
                if body.attrib['lang']==lang:
                    print("---Body-----")
                    print(body.text)
                    print("-----------------------------------------------------------------")
        returnstatement = statement
    return returnstatement

def get_item(root, ID,lang='en'):
    items=root.find(DCC+"administrativeData").find(DCC+"items")
    returnitem=[]
    for item in items:
        if ID==item.attrib['equipId'] or ID=='*':
            returnitem.append(item)
            print('------------'+item.attrib['equipId']+'------------')
            for heading in item.findall(DCC+"heading"):
                if heading.attrib['lang']==lang:
                    print(heading.text)
            for identification in item.findall(DCC+"identification"):
                for heading in identification.findall(DCC+"heading"):
                    if heading.attrib['lang']==lang:
                        print("------")
                        print(heading.text)
                print(identification.find(DCC+"value").text)
    return returnitem

def get_table(root, ID='*', lang='en'):
    returntable=[]
    tables=root.find(DCC+'measurementResults').find(DCC+'measurementResult').findall(DCC+'table')
    for table in tables:
        if ID==table.attrib['tableId'] or ID=='*':
            returntable.append(table)
            print('----------------table-------------')
            print('tableId: '+table.attrib['tableId'])
            print('itemRef: '+table.attrib['itemRef'])
            print('settingRef: '+table.attrib['settingRef'])
    return returntable

def get_setting(root, ID='*', lang='en'):
    returnsetting=[]
    settings=root.find(DCC+"administrativeData").find(DCC+"settings")
    for setting in settings:
        if ID==setting.attrib['settingId'] or ID=='*':
            returnsetting.append(setting)
            print('---------------'+setting.attrib['settingId']+'-------------')
            for heading in setting.findall(DCC+"heading"):
                if heading.attrib['lang']==lang:
                    print(heading.text)
            for body in setting.findall(DCC+"body"):
                if body.attrib['lang']==lang:
                    print(body.text)
            print('value: '+setting.find(DCC+'value').text)
    return returnsetting


def printelement(element):
    #INPUT xml-element
    #Print out the whole structure of an element to screen
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

if __name__ == "__main__":
    validate( "certificate2.xml", "dcc.xsd")
