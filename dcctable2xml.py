import xml.etree.ElementTree as et
import numpy as np
from xml.dom import minidom
import pydcc_tables as pydcc
import sys
sys.path.append(r'I:\MS\4006-03 AI metrologi\Software\DCCfunctions\develop')
from DCCfunctions import add_name


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

DCC='{https://ptb.de/dcc}'
SI='{https://ptb.de/si}'
LANG='en'
et.register_namespace("si", SI.strip('{}'))
et.register_namespace("dcc", DCC.strip('{}'))


#columns=[]
#col=pydcc.DccTableColumn(scopeType="itemBias", columnType="Value",measurandType="temperature",unit="\kelvin",humanHeading="Temperature in kelvin", columnData=["0.1", "15.3","25.1", "0.1"])
#columns.append(col)
#tab1=pydcc.DccTabel(itemID="nn11", tableID="table1",columns=columns)

tab1=tbl



xmlresults=et.Element(DCC+'results')
xmltable1=et.Element(DCC+"table",attrib={'itmeId':tab1.itemID,'refId':tab1.tableID})

columns=tab1.columns
for col in columns:
    xmlcol=et.Element(DCC+'column',attrib={'scopeType':col.scopeType, 'colType':col.columnType, 'measurandType':col.measurandType})
    if type(col.unit)!=type(None):
      et.SubElement(xmlcol,SI+'unit').text=' '.join([col.unit])
    add_name(xmlcol,lang="EN",text=col.humanHeading)
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

