import io
from lxml import etree as et
from urllib.request import urlopen
from xml.dom import minidom
import openpyxl as pyxl
import openpyxl as pyxl


# xsd_ns = {'xs':"http://www.w3.org/2001/XMLSchema"}
# DCC='{https://dfm.dk}'
# SI='{https://ptb.de/si}'
LANG='en'
# et.register_namespace("si", SI.strip('{}'))
# et.register_namespace("dcc", DCC.strip('{}'))

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
          xmlschema_doc = et.parse(xsd_file)
    else:
       xmlschema_doc = et.parse(xsd_path)

    xmlschema = et.XMLSchema(xmlschema_doc)

    xml_doc = et.parse(xml_path)
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
#%%
def load_xml(xml_path: str) -> (et._ElementTree, et.Element):
    ## return the root element from an xml file
    # et.register_namespace("si", SI.strip('{}'))
    # et.register_namespace("dcc", DCC.strip('{}'))
    parser=et.XMLParser(encoding='utf-8')
    tree=et.parse(xml_path,parser)
    root=tree.getroot()
    for k,v in root.nsmap.items():
        et.register_namespace(k,v)
    return tree, root

#%%
def get_table(root, ID='*', lang='en', show=False):
    ns = root.nsmap
    returntable=[]
    tables=root.find('dcc:measurementResults',ns).findall('dcc:table',ns)
    for table in tables:
        if ID==table.attrib['tableId'] or ID=='*':
            returntable.append(table)
            if show:
                print('----------------table-------------')
                for k,v in table.attrib.items():
                    print(f"{k}: {v}")
    
    count = len(returntable)
    if count==0:
        raise ValueError('Warning: DCC contains no tables with the required combination of setting and item Ids.')
    if count>1:
        raise ValueError('Warning: DCC contains ' + str(count) + ' tables with the required Id.\n Returning only the first instance')
    return returntable
    
    
def getTableFromResult(tableAttrib):
    #INPUT itemId and settingId list of strings
    #OUTPUT xml-element of type dcc:table

    xmltable=None
    count=0
    #Search all measurement results for a table with the required attributes
    ns = root.nsmap
    returntable=[]

    searchResults=root.find('dcc:measurementResults',ns).findall('dcc:table',ns)
    for table in searchResults:
        if tableAttrib == table.attrib:
        # if match_attributes(table.attrib, tableAttrib):
            returntable.append(table)    
    count = len(returntable)
    if count==0:
        raise ValueError('Warning: DCC contains no tables with the required combination of setting and item Ids.')
    if count>1:
        raise ValueError('Warning: DCC contains ' + str(count) + ' tables with the required Id.\n Returning only the first instance')
    return returntable[0]

# dd = dict(measuringSystemRef="ms1", tableId="MS120")
# print_node(getTableFromResult(dd))


#%%
def match_attributes(att,searchatt, unit="-", searchunit='*'):
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
    ns = table.nsmap
    cols=[]
    for col in table.findall('dcc:column',ns):
        unit=""
        if type(col.find('dcc:unit',ns)) !=type(None):
            unit=col.find('dcc:unit',ns).text
        #if col.attrib==searchattributes and searchunit==unit:
        if match_attributes(col.attrib, searchattributes,unit,searchunit):
            cols.append(col)
            #return col
    if len(cols)==0: 
        raise ValueError("No column found with the required attributes")
        return None
    return cols
       

def getColumnValues(column: et.Element) -> list: 
    return col.find("dcc:valueXMLList",column.nsmap).text.split()

# dtbl = dict(measuringSystemRef="ms1", tableId="MS120")
# tbl = getTableFromResult(dtbl)
# dcol = dict(dataCategory="Value", measurand="Measure.Volume", metaDataCategory="Data", scope="reference")
# col = getColumnFromTable(tbl,dcol,searchunit="*")[0]
# print_node(col)
# getColumnValues(col)

#%%
def getRowFromColumn(column, customerTag):
    table = column.getparent()
    try:
        tagcol=getColumnFromTable(table,{'scope':'*','dataCategory':'*','measurand':'*','metaDataCategory':'customerTag'},'*')
    except:
        raise RuntimeError("The table does not contain a customerTag column")

    # Iterate through the tags to find the row number of the specified tag
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

# dtbl = dict(measuringSystemRef="ms1", tableId="MS120")
# tbl = getTableFromResult(dtbl)
# dcol = dict(dataCategory="Value", measurand="Measure.Volume", metaDataCategory="Data", scope="reference")
# col = getColumnFromTable(tbl,dcol,searchunit="*")[0]
# print_node(col)
# getRowFromColumn(col, "p5")


#%%
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
    ns = root.nsmap

    searchValue="-"
    warning="-"
    usertagwarning="-"
    colwarning="-"
    cols=[]

    try:
        """Find the right table using measuringSystemRef and tableId"""
        tbl=getTableFromResult(tableAttrib)
        try:
            """Find the rigt column using attributes and unit"""
            cols=getColumnFromTable(tbl,colAttrib,unit)
            try:
                if type(customerTag)!=type(None):
                    searchValue=[]
                    for col in cols:
                        searchValue.append(getRowFromColumn(col,customerTag))
                else:
                    #searchValue=col[2].text.split()
                    searchValue=cols
            except Exception as e:
                usertagwarning=e.args[0]
        except Exception as e:
            colwarning=e.args[0]
    except Exception as e:
        warning=e.args[0]

    # for col in cols:
    #     print("")
    #     selectstring="heading[@lang='"+lang+"']"
    #     heading=col.find("dcc:"+selectstring,ns)
    #     if type(heading) != type(None):
    #         print(heading.text)
    #     #print(col.attrib)
    #     print("Unit:" +col.find('dcc:unit',ns).text)
    #     if type(customerTag)!=type(None):
    #         print(getRowFromColumn(col,customerTag))
    #     else:
    #         print(col[-1].text.split())

    #return [searchValue, usertagwarning, colwarning, warning]
    return searchValue

# dtbl = dict(measuringSystemRef="ms1", tableId="MS120")
# dcol = dict(dataCategory="Value", measurand="Measure.Volume", metaDataCategory="Data", scope="reference")
# rowtag = "p5"
# print_node(search(root,dtbl, dcol, "\micro\litre" )[0])
# search(root,dtbl, dcol, "\micro\litre", customerTag="p5" )
#%%
def get_statement(root, ID='*') -> list:
    ns = root.nsmap
    statements=root.findall(".//dcc:statement", ns)
    returnstatement=[]
    for statement in statements:
        if ID==statement.attrib['statementId'] or ID=='*':
            returnstatement.append(statement)
    return returnstatement

# print_node(get_statement(root,'meth1')[0])
#%%
def get_measuringSystem(root, ID='*',lang='en', show=False):
    ns = root.nsmap
    # items=root.findall("./dcc:administrativeData/dcc:measuringSystemsUnderCalibration",ns)
    items = root.findall(".//dcc:measuringSystem",ns)
    returnitem = []
    for item in items:
        if ID==item.attrib['measuringSystemId'] or ID=='*':
            returnitem.append(item)
            if show:
                print('------------'+item.attrib['measuringSystemId']+'------------')
                for heading in item.findall("dcc:heading",ns):
                    if heading.attrib['lang']==lang:
                        print(heading.text)
                for identification in item.findall("dcc:identification",ns):
                    for heading in identification.findall("dcc:heading",ns):
                        if heading.attrib['lang']==lang:
                            print("------")
                            print(heading.text)
                    print(identification.find("dcc:value").text)
    return returnitem
# print_node(get_measuringSystem(root,show=True)[0])
#%%
def get_tables(root, tableId='*', lang='en', show=False) -> list:
    ns = root.nsmap
    returntable=[]
    tables=root.find('dcc:measurementResults',ns).findall('dcc:table',ns)
    for table in tables:
        if tableId==table.attrib['tableId'] or tableId=='*':
            returntable.append(table)
            if show:
                print('----------------table-------------')
                for k,v in table.attrib.items():
                    print(f"{k}: {v}")
    return returntable
# print_node(get_tables(root, tableId="MS120")[0])
#%%
def get_setting(root, settingId='*', lang='en', show=False) -> list:
    """ Returns a list of elements fullfilling ID requirements"""
    ns = root.nsmap
    returnsetting=[]
    settings=settings = root.find("dcc:administrativeData/dcc:settings",ns)
    for setting in settings:
        if settingId==setting.attrib['settingId'] or settingId=='*':
            returnsetting.append(setting)
            if show:
                print('---------------'+setting.attrib['settingId']+'-------------')
                for heading in setting.findall("dcc:heading",ns):
                    if heading.attrib['lang']==lang:
                        print(heading.text)
                for body in setting.findall("dcc:body", ns):
                    if body.attrib['lang']==lang:
                        print(body.text)
                print('value: '+setting.find('dcc:value', ns).text)
    return returnsetting

# print_node(get_setting(root)[0])
#%%

def printelement(element):
    #INPUT xml-element
    #Print out the whole structure of an element to screen
    xmlstring=minidom.parseString(et.tostring(element)).toprettyxml(indent="   ")
    print(xmlstring)
    return

#------------------------------------------------------------------
#%%
def schema_get_restrictions(xsd_root: et._Element, 
                            type_names=['yesno', 'scopeType', 'dataCategoryType', 
                                        'statementCategoryType', 'stringPerformanceLocationType',
                                        'metaDataCategoryType', 'measurandType' ]
                            ) -> dict: 
    """schema_get_restrictions is used for finding the valid tokens for as specified in type_name:
        - yesno
        - scopeType
        - dataCategoryType
        - metaDataCategoryType
        - statementCategoryType
        - measurandType

        returns: 
            A dictionary with keys being the type_names passed in the function arguments,
            and values are the restrictions found in the schema.  
    """
    def get_restrictions(type_name, xsd_root=xsd_root):
        # xsd_ns = {'xs':"http://www.w3.org/2001/XMLSchema"}
        # type_name = 'measurandType'
        xsd_ns = xsd_root.nsmap
        s = f"xs:simpleType[@name='{type_name}']"
        r = xsd_root.findall(s, xsd_ns)
        measurandTypes = r[0].find("xs:restriction", xsd_ns)
        measurandTypes = measurandTypes.findall("xs:enumeration", xsd_ns)
        strs = [mt.get('value') for mt in measurandTypes]
        return strs

    if type(type_names) is str: 
        type_names = [type_names]
    return dict(zip(type_names, [get_restrictions(tn, xsd_root) for tn in type_names]))


# schema_get_restrictions(xsd_root)
#%%
def schema_find_all_restrictions(xsd_root):
    """Retrieves all restrictions listed in the schema and 
        returns a dictionary with kyes being the name of the parent, and restrictions in the values."""
    r = xsd_root.findall("*/xs:restriction", xsd_root.nsmap)
    names = [e.getparent().get('name') for e in r]
    restrics =  [[c.get('value') for c in e.getchildren()] for e in r]
    d = dict(zip(names, restrics))
    return d

# schema_find_all_restrictions(xsd_root)

#%%
def xpath_query(node, xpath_str: str) -> et._Element: 
    """
        xpath_query(root, "//*[@measuringSystemRef='ms2' and @tableId]")
        special operators such 'and' is not supported by lxml. 
    """
    s = xpath_str
    if xpath_str.startswith("/dcc:digitalCalibrationCertificate"):
        s = './'+xpath_str.split("/dcc:digitalCalibrationCertificate")[1]
    else: 
        s = '.'+s
    ns = node.nsmap
    for k in ns.keys():
        v = ns[k]
        v = "{"+f"{ns[k]}"+"}"
        # print(k, v)
        s = s.replace(k+":", v)
    # print(s)
    elm = node.findall(s)
    return elm    
#%%
    
def rev_ns_tag(node): 
    """Reverse the namespace tag i.e. from URI to local namespace"""
    rev_NS = {v: k for k, v in node.nsmap.items()}
    qname = et.QName(node.tag)
    ns = qname.namespace
    rev_ns = rev_NS[ns]
    ln = qname.localname
    revtag = f"{rev_ns}:{ln}"
    return revtag

#%%
def node_info(node): 
    ios = io.StringIO()
    node_display_name = rev_ns_tag(node)
    attributes = ", ".join([f"@{k}='{v}'" for k, v in node.attrib.items()])
    text_content = (node.text or '').strip()
    ios.write(f"{node_display_name}")
    if attributes:
        ios.write(f" [{attributes}]")
    if text_content:
        ios.write(f" '{text_content}'")
    ios.write("\n")
    return ios.getvalue()

# print(node_info(node))

#%%

def format_tag_name(tag_name):
    """Format the tag name by removing the namespace URL enclosed in curly braces."""
    return tag_name.rpartition('}')[-1] if '}' in tag_name else tag_name

def write_node_to_file(file, node, prefix="", is_last=False):
    # Elements for visually structured tree branches
    space = "    "
    branch = '├── '
    last_branch = '└── '
    vertical = '│   '
 
    connector = last_branch if is_last else branch
    branch_prefix = prefix + (space if is_last else vertical)
 
    node_display_name = rev_ns_tag(node)
    attributes = "; ".join([f"{k}='{v}'" for k, v in node.attrib.items()])
    text_content = (node.text or '').strip()
 
    # Check if node tag contains "XMLList" and split the text content
    if 'XMLList' in node_display_name and text_content:
        text_content = ', '.join(text_content.split())
    text_content = text_content
 
    # if len(text_content) > 150:
        # text_content = text_content[:150] + '...'
 
    file.write(f"{prefix}{connector}{node_display_name}")
    if attributes:
        file.write(f" [{attributes}]")
    if text_content:
        file.write(f" '{text_content}'")
    file.write("\n")
 
    children = list(node)
    for index, child in enumerate(children):
        write_node_to_file(file, child, branch_prefix, (index == len(children) - 1))

def node_to_str(node):
    sio = io.StringIO()
    write_node_to_file(sio, node)
    return sio.getvalue()

def print_node(node):
    print(node_to_str(node))

# print_node(root)
#%% Run tests on dcc-xml-file
if False: 
    tree, root = load_xml("SKH_10112_2.xml")
    dtbl = dict(measuringSystemRef="ms1", tableId="MS120")
    tbl = getTableFromResult(dtbl)
    dcol = dict(dataCategory="Value", measurand="Measure.Volume", metaDataCategory="Data", scope="reference")
    col = getColumnFromTable(tbl,dcol,searchunit="*")[0]
    print_node(col)
    rowtag = "p5"
    getRowFromColumn(col, rowtag)
    getColumnValues(col) 
    print_node(search(root,dtbl, dcol, "\micro\litre" )[0])
    print("SEARCH RESULT:")
    print_node(search(root,dtbl, dcol, "\micro\litre").pop())
    print(search(root,dtbl, dcol, "\micro\litre", customerTag="p5" ))
    print("----------------------GET MeasuringSystem----------------")
    # print_node(get_measuringSystem(root,show=True)[0])
    [print_node(n) for n in get_measuringSystem(root,"ms2")]
    print("----------------------get_table----------------")
    print(get_tables(root,show=False))
    print_node(get_tables(root,tableId="MS120")[0])
    print_node(get_setting(root)[0])

#%% Run tests on dcc-xml-file
if False: 
    xsd_tree, xsd_root = load_xml("dcc.xsd")
    print(schema_get_restrictions(xsd_root))
    print(schema_find_all_restrictions(xsd_root))


elif __name__ == "__main__":
    validate( "certificate2.xml", "dcc.xsd")
