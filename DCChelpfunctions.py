from lxml import etree
from urllib.request import urlopen
import xml.etree.ElementTree as et

DCC1='{https://ptb.de/dcc}'
DCC='{https://ptb.de/dcc}'
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
    humanHeading = ""
    columnData = []

    def __init__(self,
                scopeType="",
                columnType="",
                measurandType="",
                unit="",
                humanHeading = "",
                columnData=[]):
        """ """
        self.columnType = columnType
        self.scopeType = scopeType
        self.measurandType = measurandType
        self.unit = unit
        self.humanHeading = humanHeading
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
    #DCC='{https://ptb.de/dcc}'
    #SI='{https://ptb.de/si}'
    #et.register_namespace("si", SI.strip('{}'))
    #et.register_namespace("dcc", DCC.strip('{}'))
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

def item(ID, manufacturer,model):
    attributes={'itemId':ID}
    item=et.Element(DCC+'item',attrib=attributes)
    manu=et.SubElement(item,DCC+'manufacturer')
    manuname=et.SubElement(manu,DCC+'name')
    et.SubElement(manuname,DCC+'content').text=manufacturer
    et.SubElement(item,DCC+'model').text=model
    et.SubElement(item,DCC+'identifications')
    return item

def minimal_DCC():
    version="3.2.0"
    # xsilocation="dcc.xsd" #
    xsilocation="https://ptb.de/dcc dcc.xsd"
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
    ####################### items #############################
    et.SubElement(administrativeData,  DCC+'items')
    ###################### calibrationLaboratory ############
    et.SubElement(administrativeData,  DCC+'calibrationLaboratory')
    ###################### customer ########### ############
    et.SubElement(administrativeData,  DCC+'customer')
    ############### responsible persons ##################
    et.SubElement(administrativeData, DCC+'respPersons')

    ################## measurementResults ##########################
    measurementResults=et.SubElement(root,DCC+'measurementResults')
    measurementResult=et.SubElement(measurementResults,DCC+'measurementResult')
    results=et.SubElement(measurementResult,DCC+'results')
    add_name(measurementResult,text='Measurement results')
    return root


if __name__ == "__main__":
    validate( "certificate2.xml", "dcc.xsd")
