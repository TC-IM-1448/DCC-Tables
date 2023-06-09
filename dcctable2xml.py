import xml.etree.ElementTree as et
import numpy as np
from xml.dom import minidom
import pydcc_tables as pydcc
from pydcc_tables import DccTableColumn, DccTabel
import sys
sys.path.append(r'I:\MS\4006-03 AI metrologi\Software\DCCfunctions\develop')
from DCCfunctions import add_name

import numpy as np
import os

DCCf_installdir=r'I:/MS/4006-03 AI metrologi/Software/DCCfunctions/develop/'
import sys
sys.path.append(DCCf_installdir)
import DCCfunctions as DCCf
import DCCadministrative as DCCa

import xml.etree.ElementTree as et
from xml.dom import minidom
from docx import Document

#################### Make a minimal DCC and fill out the administrative data ##############################
DCC='{https://ptb.de/dcc}'
SI='{https://ptb.de/si}'
et.register_namespace("si", SI.strip('{}'))
et.register_namespace("dcc", DCC.strip('{}'))

#%%
from importlib import reload
reload(DCCa)
reload(DCCf)

#%%

root=DCCa.minimal_DCC()


administrativeData=root.find(DCC+'administrativeData')
coreData=administrativeData.find(DCC+'coreData')
coreData.find(DCC+'uniqueIdentifier').text='T2304'
DCCf.add_identification(coreData,value="NN42",issuer='calibrationLaboratory', name_dk="kundenr",name_en="customer ID")
DCCf.add_identification(coreData,value="jpx2340988",issuer="customer",name_dk="PO",name_en="PO")
coreData.find(DCC+'receiptDate').text="2022-08-13"
coreData.find(DCC+'beginPerformanceDate').text="2022-08-14"
coreData.find(DCC+'endPerformanceDate').text="2022-08-15"

lab=administrativeData.find(DCC+'calibrationLaboratory')
contact=et.SubElement(lab, DCC+'contact')
DCCa.fill_address(contact,name="DFM", eMail="srk@dfm.dk", phone="+45 2545 9040", city="Hørsholm",  postCode="2970", street="Kogle Allé", streetNo="5", further="www.dfm.dk")

respPersons=administrativeData.find(DCC+'respPersons')
DCCa.add_respPerson(respPersons,name="Erling Målermand", mainSigner=True)
DCCa.add_respPerson(respPersons,name="Simon  Hansen", mainSigner=False)

customer=administrativeData.find(DCC+'customer')
DCCa.fill_address(customer,name="NN", eMail="pqrt@nn.com", phone="+45 6160 7019", city="Søborg", postCode="2860", street="Svanevej", streetNo="12", further="kundenummer: 1234")


################ User input for item ##########################
ItemID="item_ID1"
Manufacturer='Mettler-Toledo'
Model='Platinum Super'
customerID="NN66"
SerialNo="2341-LKJQ-1324LKLJJAAFLKK33"

#Make an item XML-element
item=DCCf.item(ID=ItemID, manufacturer=Manufacturer,model=Model)
DCCf.add_name(item, 'en', 'Set of 7 weights')
DCCf.add_identification(item,customerID,issuer='customer', name_dk="MålerID", name_en="SensorID")
DCCf.add_identification(item,SerialNo,issuer='manufacturer',name_dk="Serienummer",name_en="Serial No.")

Items=root[0][2]
Items.append(item)

######################### END of administrative data ####################################################

def realListXMLList(value,unit,label=None,uncertainty=None,coverageFactor=['2'],coverageProbability=['0.95']):
# realListXMLList is a data_element that allows for arays of data.
# All inputs are expected to be lists of strings.
# All lists should have either the same length as 'value' or length==1.
   element=et.Element(SI+'realListXMLList')
   if type(label) != type(None):
      et.SubElement(element,SI+'labelXMLList').text=' '.join(label)
   et.SubElement(element,SI+'valueXMLList').text=' '.join(value)
   et.SubElement(element,SI+'unitXMLList').text=' '.join(unit)
   if type(uncertainty)!=type(None):
      expandedUnc=et.Element(SI+'expandedUncXMLList')
      et.SubElement(expandedUnc,SI+'uncertaintyXMLList').text=' '.join(uncertainty)
      et.SubElement(expandedUnc,SI+'coverageFactorXMLList').text=' '.join(coverageFactor)
      et.SubElement(expandedUnc,SI+'coverageProbabilityXMLList').text=' '.join(coverageProbability)
      element.append(expandedUnc)
   return element

LANG='en'

from excel2dcctables import read_tables_from_Excel
tab1 = read_tables_from_Excel(workbookName="DCC-mass_example.xlsx",sheetName="Table2")

#xmlresults=et.Element(DCC+'results')
xmlresults=root[1][0][1]
xmltable1=et.Element(DCC+"table",attrib={'itemId':tab1.itemID,'refId':tab1.tableID})

columns=tab1.columns
for col in columns:
    xmlcol=et.Element(DCC+'column',attrib={'scope':col.scopeType, 'dataCategory':col.columnType, 'measurand':col.measurandType})
    if type(col.unit)!=type(None):
      et.SubElement(xmlcol,SI+'unit').text=' '.join([col.unit])
    add_name(xmlcol,lang="en",text=col.humanHeading)
    #xmllist=realListXMLList(value=col.columnData,unit=[col.unit])
    et.SubElement(xmlcol,SI+"ValueXMLList").text=' '.join(col.columnData)
    #xmlcol.append(xmllist)
    xmltable1.append(xmlcol)

xmlresults.append(xmltable1)

def printelement(element):
    #DCC='https://ptb.de/dcc'
    #SI='https://ptb.de/si'
    #version="3.1.2"
    #xsilocation="https://ptb.de/dcc https://ptb.de/dcc/v"+version+"/dcc.xsd"
    #xsi="http://www.w3.org/2001/XMLSchema-instance"
    ##et.register_namespace("si", SI)
    ##et.register_namespace("dcc", DCC)
    #root=et.Element(DCC+'digitalCalibrationCertificate', attrib={"schemaVersion":version, "xmlns:dcc":DCC, "xmlns:si":SI, "xmlns:xsi":xsi, "xsi:schemaLocation":xsilocation})
    #DCC='{https://ptb.de/dcc}'
    #SI='{https://ptb.de/si}'
    #root.append(element)
    #xmlstring=minidom.parseString(et.tostring(root)).toprettyxml(indent="   ")
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

#DCCf.printelement(coreData)
#NOTE: with the namespace reprecentation ns:elementname the printelement function does not work
xmlstr=minidom.parseString(et.tostring(root)).toprettyxml(indent="   ")
with open('mass_certificate.xml','wb') as f:
    f.write(xmlstr.encode('utf-8'))


def xml2dcctable(xmltable):
    dcccolumns=[]
    for col in xmltable.findall(DCC+'column'):
        unit=""
        if type(col.find(SI+'unit')) !=type(None):
            unit=col.find(SI+'unit').text
        dcccol=DccTableColumn( scopeType=col.attrib['scope'], columnType=col.attrib['dataCategory'], measurandType=col.attrib['measurand'], unit=unit, humanHeading=col.find(DCC+'name').find(DCC+'content').text, columnData=col.find(SI+'ValueXMLList').text.split())
        dcccolumns.append(dcccol)
    length=len(col.find(SI+'ValueXMLList').text.split())
    dcctbl=DccTabel(xmltable.attrib['refId'],xmltable.attrib['itemId'],length,len(dcccolumns),dcccolumns)
    return dcctbl



