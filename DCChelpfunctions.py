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


def add_name(element,lang="",text="",append=0):
    #A human readable name may be given to various elements such as quantitie, result, item, identification
    #printelement(element)
    name=element.find(DCC+'name')
    if type(name)==type(None):
        name=et.Element(DCC+'name')
        if append:
            element.append(name)
        else:
            element.insert(0,name)
    if lang!="":
        #et.SubElement(name,DCC+'content',attrib=attrib).text=text
        et.SubElement(name,DCC+'content', attrib={'lang':lang}).text=text
    else:
        #et.SubElement(name,DCC+'content').text=text
        et.SubElement(name,DCC+'content').text=text
    return name

def fill_address(element, name, attPerson="",eMail="", phone="", fax="", city="", country="", postCode="", street="", streetNo="", further=""):
    add_name(element,text=name)
    if attPerson!="" and type(attPerson)!=type(None):
        et.SubElement(element, DCC+'attPerson').text=attPerson
    if eMail!="" and type(eMail)!=type(None):
        et.SubElement(element, DCC+'eMail').text=eMail
    if phone!="" and type(phone)!=type(None):
        et.SubElement(element, DCC+'phone').text=phone
    if fax!="" and type(fax)!=type(None):
        et.SubElement(element, DCC+'fax').text=fax
    location=et.SubElement(element,DCC+'location')
    if city!="" and type(city)!=type(None):
        et.SubElement(location, DCC+'city').text=city
    if country!="" and type(country)!=type(None):
        et.SubElement(location, DCC+'country').text=country
    if postCode!="" and type(postCode)!=type(None):
        et.SubElement(location, DCC+'postCode').text=postCode
    if street!="" and type(street)!=type(None):
        et.SubElement(location, DCC+'street').text=street
    if streetNo!="" and type(streetNo)!=type(None):
        et.SubElement(location, DCC+'streetNo').text=streetNo
    if further!="" and type(further)!=type(None):
        f=et.SubElement(location, DCC+'further')
        et.SubElement(f,DCC+'content').text=further

def add_respPerson(element, name, mainSigner=""):
    respPerson=et.SubElement(element, DCC+'respPerson')
    person = et.SubElement(respPerson, DCC+'person')
    add_name(person,text=name)
    allowedvalues=[0,1,True,False]
    if mainSigner in allowedvalues:
        if mainSigner:
            et.SubElement(respPerson,DCC+'mainSigner').text='true'
        else:
            et.SubElement(respPerson,DCC+'mainSigner').text='false'
    else:
        print("allowed values for mainSigner are ")
        print(allowedvalues)
        print(mainSigner)

def add_identification(item_element,value,issuer, name_dk="",name_en=""):
    #A number of identifications can be added to an item
    allowed_issuers= ['customer', 'manufacturer', 'calibrationLaboratory', 'laboratory', 'other']
    if issuer not in allowed_issuers:
        raise ValueError('Issuer must be one of: '+str(allowed_issuers))
    identification=et.Element(DCC+'identification')
    et.SubElement(identification,DCC+'issuer').text=issuer
    et.SubElement(identification,DCC+'value').text=value
    add_name(identification,text=name_dk,append=0)
    #add_name(identification,'en',name_en,append=1)
    item_element.find(DCC+'identifications').append(identification)
    return identification

def item(ID, category, manufacturer,model):
    attributes={'itemId':ID}
    item=et.Element(DCC+'item',attrib=attributes)
    manu=et.SubElement(item,DCC+'manufacturer')
    manuname=et.SubElement(manu,DCC+'name')
    et.SubElement(manuname,DCC+'content').text=manufacturer
    et.SubElement(item,DCC+'category').text=category
    et.SubElement(item,DCC+'model').text=model
    et.SubElement(item,DCC+'identifications')
    return item

def minimal_DCC():
    version="1.0.0"
    # xsilocation="dcc.xsd" #
    #xsilocation="https://ptb.de/dcc dcc.xsd"
    xsilocation= DCC.strip('{}') + " dcc.xsd"
    xsi="http://www.w3.org/2001/XMLSchema-instance"
    #et.register_namespace("si", SI)
    #et.register_namespace("dcc", DCC)
    #root=et.Element(DCC+'digitalCalibrationCertificate', attrib={"schemaVersion":version, "xmlns:dcc":DCC, "xmlns:si":SI, "xmlns:xsi":xsi, "xsi:schemaLocation":xsilocation})
    root=et.Element(DCC+'digitalCalibrationCertificate',
            attrib={"schemaVersion":version,   "xmlns:xsi":xsi, "xsi:schemaLocation":xsilocation})
    administrativeData=et.SubElement(root,DCC+'administrativeData')
    ################## Software ##############################
    dccSoftware=et.SubElement(administrativeData, DCC+'dccSoftware')
    software=et.SubElement(dccSoftware, DCC+'software')
    add_name(software,text='DCCfunctions.py')
    et.SubElement(software,DCC+'release').text='0.0'
    #################### coreData ###############################
    performanceLocation="laboratory"
    coreData=et.SubElement(administrativeData,  DCC+'coreData')
    et.SubElement(coreData,DCC+'countryCodeISO3166_1').text='DK'
    et.SubElement(coreData,DCC+'usedLangCodeISO639_1').text='da'
    et.SubElement(coreData,DCC+'usedLangCodeISO639_1').text='en'
    et.SubElement(coreData,DCC+'mandatoryLangCodeISO639_1').text=LANG
    et.SubElement(coreData,DCC+'uniqueIdentifier')
    #et.SubElement(coreData,DCC+'identifications')
    #et.SubElement(coreData,DCC+'receiptDate')
    #et.SubElement(coreData,DCC+'beginPerformanceDate')
    #et.SubElement(coreData,DCC+'endPerformanceDate')
    #et.SubElement(coreData,DCC+'performanceLocation').text=performanceLocation
    ###################### calibrationLaboratory ############
    et.SubElement(administrativeData,  DCC+'calibrationLaboratory')
    ###################### customer ########### ############
    et.SubElement(administrativeData,  DCC+'customer')
    ############### responsible persons ##################
    et.SubElement(administrativeData, DCC+'respPersons')
    ####################### items #############################
    et.SubElement(administrativeData,  DCC+'items')

    ################## measurementResults ##########################
    measurementResults=et.SubElement(root,DCC+'measurementResults')
    #measurementResult=et.SubElement(measurementResults,DCC+'measurementResult')
    #results=et.SubElement(measurementResult,DCC+'results')
    #add_name(measurementResult,text='Measurement results')
    return root

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

"""
def get_item(root, ID="*", issuer="*", lang='en'):
    items=root.find(DCC+"administrativeData").find(DCC+"items")
    returnitem=[]
    if issuer!="*":
        selectstring="identification[@issuer='"+issuer+"']"
    else:
        selectstring="identification"
    for item in items:
        identification=item.find(DCC+selectstring)
        idvalue=identification.find(DCC+"value").text
        if ID==idvalue or ID=='*':
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
"""

           

def printelement(element):
    #INPUT xml-element
    #Print out the whole structure of an element to screen
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return



if __name__ == "__main__":
    validate( "certificate2.xml", "dcc.xsd")
