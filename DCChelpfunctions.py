#%%
import io
from lxml import etree as et
from urllib.request import urlopen
from xml.dom import minidom
from typing import Tuple
#import openpyxl as pyxl

#%%
# xsd_ns = {'xs':"http://www.w3.org/2001/XMLSchema"}
# DCC='{https://dfm.dk}'
# SI='{https://ptb.de/si}'
LANG='en'
XSD_RESTRICTION_NAMES = [
                        'stringISO3166Type',
                        'stringISO639Type',
                        'yesno', 
                        "transactionContentType",
                        'statementCategoryType', 
                        'accreditationApplicabilityType',
                        'accreditationNormType',
                        'equipmentCategoryType',
                        'operationalStatusType', 
                        'performanceLocationType',
                        'conformityStatusType',
                        'scopeType',
                        'dataCategoryType', 
                        'quantityType',
                        'tableCategoryType',
                        'approachToTargetRestrictType',
                        'quantityCodeSystemType',
                        'positionSystemType',
                        'contactAndLocationColumnHeadingNamesType',
                        'statementColumnHeadingNamesType',
                        'equipmentColumnHeadingNamesType',
                        'settingColumnHeadingNamesType',
                        'measurementSetupColumnHeadingNamesType',
                        'quantityUnitDefsColumnHeadingNamesType',
                        ]
# et.register_namespace("si", SI.strip('{}'))
# et.register_namespace("dcc", DCC.strip('{}'))


#%%
def transpose_2d_list(matrix):
    return [list(row) for row in zip(*matrix)]
#%%

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



#%%
def load_xml(xml_path: str) -> Tuple[et._ElementTree, et._Element]:
    ## return the root element from an xml file
    # et.register_namespace("si", SI.strip('{}'))
    # et.register_namespace("dcc", DCC.strip('{}'))
    parser=et.XMLParser(encoding='utf-8', remove_comments=True)
    tree=et.parse(xml_path,parser)
    root=tree.getroot()
    for k,v in root.nsmap.items():
        et.register_namespace(k,v)
    return tree, root

#%%
def match_table_attributes(att,searchAttrib,searchTableType='*'):
    for key in att.keys():
        if att[key]!='-' and searchAttrib[key]!='*' and att[key]!=searchAttrib[key]:
            return False
    return True

#%%
def getTables(root: et._Element,search_attrib={}, tableType='*') -> list:
    ns = root.nsmap
    default_search_attrib = dict(tableId='*',
                                 measuringSystemRef="*", 
                                 serviceCategory="*", 
                                 customServiceCategory='*', 
                                 statementRef='*',
                                 numRows='*',
                                 numCols='*')
    default_search_attrib.update(search_attrib)
    search_attrib = default_search_attrib
    # print(search_attrib)
    returntable=[]
    tables=root.find('dcx:measurementResults',ns).getchildren() #findall('dcx:table',ns)
    for table in tables:
        if (tableType == '*' or tableType == rev_ns_tag(table)) and match_table_attributes(table.attrib, search_attrib):
            returntable.append(table)

    # count = len(returntable)
    # if count==0:
    #     raise ValueError('Warning: DCC contains no tables with the required combination of setting and item Ids.')
    # if count>1:
    #     raise ValueError('Warning: DCC contains ' + str(count) + ' tables with the required Id.\n Returning only the first instance')
    return returntable
    
#%%
def match_column_attributes(att,searchatt, dataCategory, searchDataCategory, unit="-"):
    # print(f"On Column with: {att}, {dataCategory}, {unit}")
    for key in att.keys():
        if att[key]!='-' and searchatt[key]!='*' and att[key]!=searchatt[key]:
            return False
    if dataCategory!='-' and searchDataCategory!='*' and dataCategory!=searchDataCategory:
        return False
    return True



def getColumnsFromTable(table,searchattributes, searchDataCategory="") -> list:
    #INPUT: xml-element of type dcx:table
    #INPUT: attribute dictionary
    #INPUT: searchunit as string.
    #OUTPUT: list of xml-element of type dcx:column
    ns = table.nsmap
    cols=[]
    # print(f"searching for: {searchattributes}, {searchDataCategory}, {searchunit}")
    for col in table.findall('dcx:column',ns):
        unit=""
        dataCategory=rev_ns_tag(col.getchildren()[-1]).replace("dcx:","")
        #if col.attrib==searchattributes and searchunit==unit:
        if match_column_attributes(col.attrib, searchattributes, dataCategory, searchDataCategory):
            cols.append(col)
            #return col
    if len(cols)==0: 
        raise ValueError("No column found with the required attributes")
        return []
    return cols


#%%
def getRowData(column: et._Element, search_idxs=[]) -> list:
    # Iterate through the tags to find the row number of the specified tag
    search_idxs = list(map(str, search_idxs))
    dataList = column.getchildren()[-1]
    dataType = rev_ns_tag(dataList)
    dataList = {row.attrib['idx']: row.text for row in dataList}
    if search_idxs == []:
        search_idxs = dataList.keys()
    search_result = {}
    for idx in search_idxs:  
        if idx in dataList.keys(): 
            rowData = dataList[idx]
            if dataType == 'dcx:real': 
                rowData = eval(f"float({rowData})")
            elif dataType == 'dcx:int':
                rowData = eval(f"int({rowData})")
            search_result[idx] = rowData
        else:
            raise Exception(f"Row index idx={idx} not found in Column dataList!") 
            return None
    return search_result

#%%
def getRowTagColumns(tbl) -> list: 
    # return tbl.findall("./dcx:column[@dataCategory='rowTag']",tbl.nsmap)
    tagCols =  [c.getparent() for c in tbl.findall("*/dcx:rowTag",tbl.nsmap)]
    return tagCols

def getRowTagsFromRowTagColumn(col: et._Element) -> dict:
    rowTags = {elm.attrib["idx"]: elm.text for elm in col.findall(".//*[@idx]", col.nsmap)}
    return rowTags

def rowTagsToIndexs(rowTagColumn: et._Element) -> dict: 
    """Returns {tag:idx}"""
    rowTags = getRowTagsFromRowTagColumn(rowTagColumn)
    return {v: k for k, v in rowTags.items()}

#%%
def search(root, tableAttrib, colAttrib, dataCategory, tableType="dcx:calibrationResult", rowTags=[], idxs=[], lang="en") -> list:
    """Searches for a specific value in a DCC measurement table. 

    Args:
        root: eTree._Element
            etree root element of the DCC
        tableAttributes:
            itemRef, settingRef and tableId as dictionary of string values
        coAttributes scope, dataCategory and quantity  as dictionary of string values
        unit as string
        customerTag (optional)  as string
    
    Returns:
        search result as string (or list of strings if customerTag is not specified)
        warnings as strings 
    
    :NOTE: rowTags takes prior rank to idxs if both are provided. 
        
    :NOTE: This function is not really necessary any more, as the following xpath expressions
          serve the same purpose, except that rowTags can be used directly in this function.: 
            root.findall('*//dcx:calibrationResult[@measuringSystemRef="ms1"]/dcx:column[@scope="reference"][@dataCategoryRef="-"][@quantity="3-4|volume|m3"]/dcx:value/dcx:row[@idx="1"]',root.nsmap)
            root.findall('*//*[@measuringSystemRef="ms1"]/*[@scope="reference"][@dataCategoryRef="-"][@quantity="3-4|volume|m3"]/dcx:value/*[@idx="1"]',root.nsmap)
    """
    ns = root.nsmap

    searchValue=[]
    warning="-"
    usertagwarning="-"
    colwarning="-"
    cols=[]

    try:
        """Find the right table using measuringSystemRef and tableId"""
        tbls=getTables(root, tableAttrib, tableType)
        # print(tbls)
        if len(tbls) != 1: 
                raise Exception("Found multiple columns - search should be unique")
        tbl = tbls[0]
        print(tbl.attrib['tableId'])
        try:
            """Find the rigt column using attributes and unit"""
            cols=getColumnsFromTable(tbl,colAttrib,dataCategory)
            # print(cols)
            if len(cols) != 1: 
                raise Exception("Found multiple columns - search should be unique")
            col = cols[0]
            # print(col.attrib)
            try:
                """Convert rowTags to index's - checks for uniquenes of rowTag column"""
                if rowTags!=[]: 
                    tagColumns = getRowTagColumns(tbl)
                    if len(tagColumns) != 1: 
                        raise Exception("Multiple rowTag columns identified, please use another method to identify required row indexes.")
                    tagColumn = tagColumns[0]
                    rowTagIdxs=rowTagsToIndexs(tagColumns[0])
                    idxs = [rowTagIdxs[rowTag] for rowTag in rowTags]
                try: 
                    searchValue=getRowData(col, idxs)
                except Exception as e: 
                    getrowdataWarning = e.args[0]
            except Exception as e:
                rowtagwarning=e.args[0]
        except Exception as e:
            colwarning=e.args[0]
    except Exception as e:
        warning=e.args[0]
    return searchValue

# dtbl = dict(measuringSystemRef="ms1", tableId="MS120")
# dcol = dict(dataCategory="Value", quantity="Measure.Volume", metaDataCategory="Data", scope="reference")
# rowtag = "p5"
# print_node(search(root,dtbl, dcol, "\micro\litre" )[0])
# search(root,dtbl, dcol, "\micro\litre", customerTag="p5" )
#%%
def get_languages(root) -> list:
    ns = root.nsmap
    # mandatory_lang = root.findall(".//dcx:mandatoryLangCodeISO639_1",ns)
    # used_lang = root.findall(".//dcx:usedLangCodeISO639_1",ns)
    # langs = mandatory_lang + used_lang
    # langs = [x.text for x in langs]
    # unique_langs = []
    # [unique_langs.append(x) for x in langs if x not in unique_langs]
    unique_langs = list(set(root.xpath("//@lang", namespaces=root.nsmap)))
    return unique_langs

#%%
def get_statements(root, ID='*') -> list:
    ns = root.nsmap
    statements=root.findall(".//dcx:statement", ns)
    returnstatement=[]
    for statement in statements:
        if ID==statement.attrib['id'] or ID=='*':
            returnstatement.append(statement)
    return returnstatement

# print_node(get_statement(root,'meth1')[0])
#%%
def get_measuringSystems(root, ID='*',lang='en', show=False) -> list:
    ns = root.nsmap
    # items=root.findall("./dcx:administrativeData/dcx:measuringSystemsUnderCalibration",ns)
    items = root.findall(".//dcx:measuringSystem",ns)
    returnitem = []
    for item in items:
        if ID==item.attrib['id'] or ID=='*':
            returnitem.append(item)
            if show:
                print('------------'+item.attrib['id']+'------------')
                for heading in item.findall("dcx:heading",ns):
                    if heading.attrib['lang']==lang:
                        print(heading.text)
                for identification in item.findall("dcx:identification",ns):
                    for heading in identification.findall("dcx:heading",ns):
                        if heading.attrib['lang']==lang:
                            print("------")
                            print(heading.text)
                    print(identification.find("dcx:value").text)
    return returnitem
# print_node(get_measuringSystem(root,show=True)[0])
#%%
def get_setting(root, settingId='*', lang='en', show=False) -> list:
    """ Returns a list of elements fullfilling ID requirements"""
    ns = root.nsmap
    returnsetting=[]
    settings=settings = root.findall("dcx:settings/dcx:setting",ns)
    for setting in settings:
        if settingId==setting.attrib['settingId'] or settingId=='*':
            returnsetting.append(setting)
            if show:
                print('---------------'+setting.attrib['id']+'-------------')
                for heading in setting.findall("dcx:heading",ns):
                    if heading.attrib['lang']==lang:
                        print(heading.text)
                for body in setting.findall("dcx:body", ns):
                    if body.attrib['lang']==lang:
                        print(body.text)
                print('value: '+setting.find('dcx:value', ns).text)
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
                            type_names=XSD_RESTRICTION_NAMES
                            ) -> dict: 
    """schema_get_restrictions is used for finding the valid tokens for as specified in type_name:
        - yesno
        - statementCategoryType
        - scopeType
        - dataCategoryType
        - metaDataCategoryType
        - quantityType
        - and more 

        returns: 
            A dictionary with keys being the type_names passed in the function arguments,
            and values are the restrictions found in the schema.  
    """
    def get_restrictions(type_name, xsd_root=xsd_root):
        # xsd_ns = {'xs':"http://www.w3.org/2001/XMLSchema"}
        # type_name = 'quantityType'
        xsd_ns = xsd_root.nsmap
        s = f"xs:simpleType[@name='{type_name}']"
        r = xsd_root.findall(s, xsd_ns)
        if len(r) == 0: 
            raise KeyError(f"No element found with name='{type_name}'") 
        quantityTypes = r[0].find("xs:restriction", xsd_ns)
        quantityTypes = quantityTypes.findall("xs:enumeration", xsd_ns)
        strs = [mt.get('value') for mt in quantityTypes]
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
def schemaFindAdministrativeDataChildren(xsd_root):
    adminNode = xsd_root.xpath('.//*[@name="administrativeDataType"]')[0]
    tmp = adminNode.find("xs:all",adminNode.nsmap).getchildren()
    adminDataTags = ['dcx:'+e.attrib['name'] for e in tmp]
    return adminDataTags

def getNodePath(node: et._Element) -> str:
    """Returns the xpath of the element"""
    return node.getroottree().getpath(node)

def getNamesAndTypes(xsd_root: et._Element ,typeName: str) -> Tuple[list, list]: 
    """Extract data from schema regarding Type Element content names and types
      of elements and attribtues within the type element"""
    elements = xpath_query(xsd_root, f'.//*[@name="{typeName}Type"]/*/xs:element')
    elementNames = [e.get('name') for e in elements if not e.get('name') == 'heading']
    elementTypes = [e.get('type') for e in elements if not e.get('name') == 'heading']
    attributes = xpath_query(xsd_root, f'.//*[@name="{typeName}Type"]/xs:attribute')
    attributeNames = [e.get('name') for e in attributes]
    attributeTypes = [e.get('type') for e in attributes]
    return elementNames+attributeNames, elementTypes+attributeTypes

#%%
def schemaGetAdministrativeDataStructure(xsd_root: et._Element) -> list:
    """schemaGetAdministrativeDataStructure is used for finding the structure
    of the administrative data in the schema.

    args:
        xsd_root: et._Element
            The root element of the schema

    Returns:
        A list of tuples containing the following information:
            - level: int
            - discr: str
            - xsdType: str
            - path: str
    """
    adminNode = xsd_root.xpath('.//*[@name="administrativeDataType"]')[0]
    root_path = '/dcx:digitalCalibrationExchange'
    discrs = ['title', 'subtitle', 'administrativeData']
    levels = [0, 0, 0]
    xsdTypes = ['dcx:transactionContentFieldType', 'dcx:stringWithLangType', 'dcx:administrativeDataType']
    paths = [root_path+'/dcx:title',
             root_path+'/dcx:subtitle',
             root_path+'/dcx:administrativeData',]

    # adminElementNames = xpath_query(xsd_root, './/*[@name="administrativeDataType"]/xs:sequence/*/@name')
    # adminElementNames = [e for e in adminElementNames if e not in ['heading']]
    # adminElementTypes = [xpath_query(xsd_root, f'.//*[@name="administrativeDataType"]/xs:sequence/*[@name="{e}"]/@type') for e in adminElementNames]

    def append_to_data(discr, level, path, xsdType):
        discrs.append(discr)
        levels.append(level)
        paths.append(path)
        xsdTypes.append(xsdType)

    adminElementNames, adminElementTypes = getNamesAndTypes(xsd_root, 'administrativeData')
    N = adminElementNames.index('contactColumnHeadings')
    adminPath = root_path+'/dcx:administrativeData'

    for idx,sectionName in enumerate(adminElementNames[:N]):
               
        discr = sectionName
        level = 1
        sectionPath = adminPath+f'/dcx:{sectionName}'
        xsdType = adminElementTypes[idx]
        append_to_data(discr, level, sectionPath, xsdType)
        
        elementNames, elementTypes = getNamesAndTypes(xsd_root, sectionName) 

        for idx2,elementName in enumerate(elementNames):
            discr = elementName
            level = 2
            subsecPath = sectionPath+f'/dcx:{elementName}'
            xsdType = elementTypes[idx2]
            append_to_data(discr, level, subsecPath, xsdType)
        
            if xsdType == "dcx:responsiblePersonType": 
                level = 3
                xsdType = 'dcx:stringFieldType'
                discr = 'name'
                path = subsecPath+'/dcx:name'
                append_to_data(discr, level, path, xsdType)
                discr = 'email'
                path = subsecPath+'/dcx:email'
                append_to_data(discr, level, path, xsdType)
    values = [None]*len(discrs)
    return list(map(list,zip(levels, discrs, values, values, values, xsdTypes, paths))), (adminElementNames[N:], adminElementTypes[N:])

# list(schemaGetAdministrativeDataStructure(xsd_root)[0])

#%% get id element
def getNodeById(root, ID:str):
    nodes = root.xpath(f'//*[@*="{ID}"]')
    if len(nodes) == 0: 
        raise KeyError( "No elements found")
    elif len(nodes)>1: 
        raise KeyError( f"Too many elements: found {len(nodes)} elements expected 1.")
    node = nodes[0]
    nTag = rev_ns_tag(node)
    return nTag, node

#%%
# def xpath_query(node, xpath_str: str) -> et._Element: 
#     """
#     xpath_query is a wrapper around the lxml findall function.

#     Example: 
#         xpath_query(root, "//*[@measuringSystemRef='ms2' and @tableId]")
#         xpath_query(root, "//*[@measurementSetupRef='ms1'][@tableId='calRes0']")
        
#     Note: special operators such 'and' is not supported by lxml. 
#     """
#     s = xpath_str
#     if xpath_str.startswith("/dcx:digitalCalibrationCertificate"):
#         s = './'+xpath_str.split("/dcx:digitalCalibrationCertificate")[1]
#     else: 
#         s = '.'+s
#     ns = node.nsmap
#     for k in ns.keys():
#         v = ns[k]
#         v = "{"+f"{ns[k]}"+"}"
#         # print(k, v)
#         s = s.replace(k+":", v)
#     # print(s)
#     elm = node.findall(s)
#     return elm    
#%%

def xpath_query(node: et._Element, xpath_str: str) -> et._Element:
    """
    xpath_query is a wrapper around the lxml findall function.

    Example: 
        xpath_query(root, "//*[@measuringSystemRef='ms2' and @tableId]")
        xpath_query(root, "//*[@measurementSetupRef='ms1'][@tableId='calRes0']")
        
    Note: special operators such 'and' is not supported by lxml. 
    """
    nodes = node.xpath(xpath_str, namespaces=node.nsmap)
    return nodes
    
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
    #%%    
    tree, root = load_xml("dcc-example.xml")
    dtbl = dict(tableId='*',measuringSystemRef="ms1", serviceCategory="*")
    print("----------------------get_table----------------")
    tbl = getTables(root,dtbl,tableType="dcx:calibrationResult")[0]
    print(tbl)
    #%%
    # print_node(get_measuringSystem(root,show=True)[0])
    dcol = dict(quantity="3-4|volume|m3", dataCategoryRef="*", scope="reference", unit='µL')
    col = getColumnsFromTable(tbl,dcol, searchDataCategory="value")[0]

    # dcol = dict(quantity="quantityUnitDefRef", dataCategoryRef="-", scope="environment", quantityUnitDefRef="ms1")
    # dataCategoryRef="-" quantity="quantityUnitDefRef" scope="environment" quantityUnitDefRef="ms1"
    # col = getColumnsFromTable(tbl,dcol, searchDataCategory="value", searchunit="*")[0]
    print(col)
    print_node(col)
    #%%
    rowtag = "1"
    idxs = [1,3]
    search_data = getRowData(col, idxs)
    print("Search_data:  ", search_data)

    tagCols = getRowTagColumns(tbl)
    print(tagCols)
    print_node(tagCols[0])

    print(getRowTagsFromRowTagColumn(tagCols[0]))
    print(rowTagsToIndexs(tagCols[0]))


    search_result = search(root, dtbl, dcol, "value", tableType="dcx:calibrationResult")
    print("SEARCH RESULT for Column: ", search_result)
    #%%

    search_result = search(root, dtbl, dcol, "value", rowTags=['pt1','pt3'] ,idxs=[1,2])
    print("SEARCH RESULT for specific Rows", search_result)

    #%%
    print("----------------------GET MeasuringSystem----------------")
    for n in get_measuringSystems(root,"ms1"): print_node(n)
    #%%
    get_setting(root)
    print_node(get_setting(root)[0])
    print_node(getTables(root,dict(tableId="ser13"))[0])
    #%%
    
    statementIds = [elm.attrib['id'] for elm in get_statements(root)]
    statementIds
    #%%
    dtbl = dict(tableId='*',measuringSystemRef="ms1")
    print("----------------------get_table----------------")
    tbl = getTables(root,dtbl,tableType="dcx:calibrationResult")[0]
    print(tbl)
    # print_node(get_measuringSystem(root,show=True)[0])
    #%%
    dcol = dict( quantity="3-4|volume|m3", dataCategoryRef='*', scope='reference', unit="*")
    col = getColumnsFromTable(tbl,dcol,searchDataCategory="value")
    print_node(col[0])
    #%%
    search(root, dtbl, dcol, dataCategory='value', rowTags=['pt1'])




if False: 
#%% Run tests on dcc-xml-file
    xsd_tree, xsd_root = load_xml("dcx.xsd")
    da = schema_find_all_restrictions(xsd_root)
    d = schema_get_restrictions(xsd_root)
    # v = validate( "SKH_10112_2.xml", "dcc.xsd")
#%%
    v = validate( "output.xml", "dcx.xsd")
    print(v)
    v = validate( "dcc-example.xml", "dcx.xsd")
    print(v)
    # print(validate("Examples\\Stip-230063-V1.xml", "dcc.xsd"))
#%%
if False:
    #%%
    tree, root = load_xml("dcc-example.xml")
    nodes = root.findall('.//*[@measurementSetupRef="ms1"]/*[@scope="reference"][@dataCategoryRef="-"][@quantity="3-4|volume|m3"]/dcx:value/*[@idx="1"]',root.nsmap)
    print_node(nodes[0])
    #%%
    nodes = root.findall('*//*[@measurementSetupRef="ms1"]/*[@scope="reference"][@dataCategoryRef="-"][@quantity="3-4|volume|m3"][@unit="µL"]/dcx:value/*[@idx="1"]',root.nsmap)
    print_node(nodes[0])
    #%%
    nodes = root.findall('*//*[@measurementSetupRef="ms1"]/*[@scope="reference"][@dataCategoryRef="-"][@unit="µL"]/dcx:value/*[@idx="1"]',root.nsmap)
    print(nodes)
    print_node(nodes[0])

#%%
if __name__ == "__main__":
    validate( "SKH.xml", "dcc.xsd")
