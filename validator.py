from lxml import etree
from urllib.request import urlopen
import sys

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

if __name__ == "__main__":
    validate( "mass_certificate.xml", "dcc.xsd")
#    if not sys.argv[0] == "":
#        validate(sys.argv[0], "dcc.xsd")