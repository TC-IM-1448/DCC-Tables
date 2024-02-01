# 
# by DBH 2024-01-11
#
# Uses xlwings see following references: 
# https://docs.xlwings.org/en/stable/syntax_overview.html 
# https://docs.xlwings.org/en/latest/ 
#%%
__ver__ = "DCC-EXCEL-GUI v. 0.0.1"
#%%
import os
import openpyxl as pyxl
import xlwings as xw
from  lxml import etree as et
from lxml import builder as etb
from  xml.etree import ElementTree as xmlEt 
import tkinter as tk
import tkinter.filedialog as tkfd
import DCChelpfunctions as dcchf
from DCChelpfunctions import search


LANG='da'
DCC='{https://dfm.dk}'

xmlEt.register_namespace("dcc", DCC.strip('{}'))
xlValidateList = xw.constants.DVType.xlValidateList
#%%
class DccQuerryTool(): 
    xsdDefInitCol = 3  

    colors = dict(  yellow = "#ffd966",
                    light_yellow = '#fff2cc',
                    green = '#c6e0b4',
                    light_green = '#e2efda',
                    gray = '#bfbfbf',
                    light_gray = '#d9d9d9',
                    blue = '#bdd7ee',
                    light_blue = '#ddebf7',
                    red = '#f8cbad',
                    light_red = '#fce4d6')

    def loadExcelWorkbook(self, workBookFilePath: str):
        self.wb = xw.Book(workBookFilePath)
        wb = self.wb
        self.sheetDef = wb.sheets['Definitions']
        self.wb = wb
        self.loadSchemaFile()
        self.loadSchemaRestrictions()
        self.langs = ['da','en']
        wb.activate(steal_focus=True)

    def loadSchemaFile(self, xsdFileName="dcc.xsd"):
        self.xsdTree, self.xsdRoot  = dcchf.load_xml(xsdFileName)
        
    def loadDCCFile(self, xmlFileName="SKH_10112_2.xml"):
        self.dccTree, self.dccRoot = dcchf.load_xml(xmlFileName)
        
    def loadSchemaRestrictions(self): 
        xsd_root = self.xsdRoot
        sht_def = self.sheetDef
        drestr = dcchf.schema_get_restrictions(xsd_root)
        j = self.xsdDefInitCol
        for i, (k,vs) in enumerate(drestr.items()):
            sht_def.range((1, i+j)).value = [k]
            sht_def.range((2, i+j)).value = [[v] for v in vs]
            sht_def.range((2,i+j)).expand('down').name = k  
        self.dccDefInitCol = i+1

    # def loadDccStatements(self, after='AdministrativeData'):
    #     root = self.dccRoot
    #     wb = self.wb
    #     lang1 = 'da'
    #     lang2 = 'en'

    #     if not 'Statements' in wb.sheet_names:
    #         wb.sheets.add('Statements', after=after)
    #     sht = wb.sheets['Statements']
    #     statements = dcchf.get_statements(root)
    #     heading = ['in DCC', 'category', 'id', 'heading lang1', 'body lang1', 'heading lang2', 'body lang2']
    #     sht.range((1,1), (1,7)).value = heading
    #     ns = root.nsmap

    #     statementIds = [elm.attrib['statementId'] for elm in dcchf.get_statements(root)]
    #     statementCat = [elm.attrib['category'] for elm in dcchf.get_statements(root)]
    #     for idx, statement in enumerate(statements):
    #         inDCC = 'y'
    #         statementId = statement.attrib['statementId']
    #         category = statement.attrib['category']
    #         lang1Heading = statement.find(f'./dcc:heading[@lang="{lang1}"]',ns).text
    #         lang1Body = statement.find(f'./dcc:body[@lang="{lang1}"]',ns).text
    #         lang2Heading = statement.find(f'./dcc:heading[@lang="{lang2}"]',ns).text
    #         lang2Body = statement.find(f'./dcc:body[@lang="{lang2}"]',ns).text
    #         rng = sht.range((idx+2,1))
    #         rng.value = [inDCC, category, statementId, lang1Heading, lang1Body, lang2Heading, lang2Body]
        
    #     rng = sht.range((2,2),(1024,2))
    #     rng.api.Validation.Delete()
    #     rng.api.Validation.Add(Type=xlValidateList, Formula1='=statementCategoryType')

    # def loadDCCEquipment(self, after='Statements'):
    #     root = self.dccRoot
    #     ns = root.nsmap
    #     wb = self.wb
    #     lang1 = 'da'
    #     lang2 = 'en'

    #     if not 'Equipment' in wb.sheet_names:
    #         wb.sheets.add('Equipment', after=after)
    #     sht = wb.sheets['Equipment']
    #     equipment = root.find(".//dcc:equipment",ns) 
    #     equipItems = equipment.getchildren()

    #     heading = ['in DCC', 'category', 'equipId', 'heading lang1', 'heading lang2', 'manufacturer', 'productName', 
    #                 'id1 value', 'id1 issuer', 'id1 heading lang1', 'id1 heading lang2', 
    #                 'id2 value', 'id2 issuer', 'id2 heading lang1', 'id2 heading lang2' ]   

    #     sht.range((1,1), (1,7)).value = heading

    #     equipId = [elm.attrib['equipId'] for elm in equipItems]
    #     equipCat = [elm.attrib['category'] for elm in equipItems]
    #     for idx, equip in enumerate(equipItems):
    #         inDCC = 'y'
    #         category = equip.attrib['category']
    #         equipId = equip.attrib['equipId']
    #         ids = [inDCC, category, equipId]
    #         # dcchf.print_node(equip)
    #         headingLang1 = equip.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
    #         headingLang2 = equip.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
    #         productName = equip.findall('dcc:productName',ns)
    #         manufacturer = equip.findall('dcc:manufacturer',ns)
    #         tmp = [headingLang1, headingLang2, manufacturer, productName]
    #         tmp =  [None if i == [] else i[0].text for i in tmp]
    #         ids = ids+tmp
    #         # print(ids)
    #         identifications = equip.findall('.//dcc:identification',ns)
    #         for Id in identifications:
    #             issuer = Id.attrib['issuer']
    #             lang1Heading = Id.find(f'./dcc:heading[@lang="{lang1}"]',ns).text
    #             lang2Heading = Id.find(f'./dcc:heading[@lang="{lang2}"]',ns).text
    #             idValue = Id.find('./dcc:value',ns).text
    #             ids.extend([idValue, issuer, lang1Heading,lang2Heading])
    #         rng = sht.range((idx+2,1))
    #         rng.value = ids
        
    #     rng = sht.range((2,2),(1024,2))
    #     rng.api.Validation.Delete()
    #     rng.api.Validation.Add(Type=xlValidateList, Formula1='=equipmentCategoryType')
    #     rng = sht.range((2,9),(1024,9))

    #     rng.api.Validation.Delete()
    #     rng.api.Validation.Add(Type=xlValidateList, Formula1='=issuerType')
    #     sht.activate()

    # def loadDCCSettings(self, after='Equipment'): 
    #     root = self.dccRoot
    #     ns = root.nsmap
    #     wb = self.wb
    #     lang1 = 'da'
    #     lang2 = 'en'

    #     if not 'Settings' in wb.sheet_names:
    #         wb.sheets.add('Settings', after=after)
    #     sht = wb.sheets['Settings']
    #     settings = root.find(".//dcc:settings",ns) 
    #     settingList = settings.getchildren()

    #     heading = ['in DCC', 'settingId', 'refId', 
    #                'value', 'unit',  
    #                'heading lang1', 'body lang1', 
    #                'heading lang2', 'body lang2']

    #     sht.range((1,1), (1,7)).value = heading

    #     for idx, setting in enumerate(settingList):
    #         inDCC = 'y'
    #         settingRefId = setting.attrib['refId'] if 'refId' in setting.attrib else None
    #         settingId = setting.attrib['settingId']
    #         ids = [inDCC, settingId, settingRefId]
    #         headingLang1 = setting.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
    #         bodyLang1 = setting.findall(f'./dcc:body[@lang="{lang1}"]',ns)
    #         headingLang2 = setting.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
    #         bodyLang2 = setting.findall(f'./dcc:body[@lang="{lang2}"]',ns)
    #         value = setting.findall('dcc:value',ns)
    #         unit = setting.findall('dcc:unit',ns)
    #         tmp = [value, unit, headingLang1, bodyLang1, headingLang2, bodyLang2]
    #         tmp =  [None if i == [] else i[0].text for i in tmp]
    #         ids = ids+tmp
    #         rng = sht.range((idx+2,1))
    #         rng.value = ids
        
    #     rng = sht.range((1,1)).expand('table')
    #     self.resizeXlTable(rng,sht,'TableSettings')
    #     rng.api.WrapText = True
    #     rng = sht.range((1,3)).expand('table')
    #     rng.columns

        
    # def loadDCCMeasurementSystem(self, after='Settings'):
    #     root = self.dccRoot
    #     ns = root.nsmap
    #     wb = self.wb
    #     lang1 = 'da'
    #     lang2 = 'en'

    #     if not 'MeasuringSystems' in wb.sheet_names:
    #         wb.sheets.add('MeasuringSystems', after=after)
    #     sht = wb.sheets['MeasuringSystems']
    #     msuc = root.find(".//dcc:measuringSystems",ns) 

    #     heading = ['in DCC', 'Id', 'Instrument & Settings Refs', 'headingLang1', 'bodyLang1', 'headingLang2', 'bodyLang2']

    #     ids = []
    #     mssHeadings = msuc.findall("dcc:heading",ns)
    #     for h in mssHeadings:
    #         if h.attrib['lang'] == lang1: sht.range((1,2)).value = h.text
    #         if h.attrib['lang'] == lang2: sht.range((2,2)).value = h.text

    #     tblRowIdx = 3    
    #     sht.range((tblRowIdx,1)).value = heading

    #     msList = msuc.findall("./dcc:measuringSystem",ns)
    #     for idx, ms in enumerate(msList):
    #         inDCC = 'y'
    #         msId = ms.attrib['measuringSystemId']
    #         ids = [inDCC, msId]
    #         headingLang1 = ms.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
    #         bodyLang1 = ms.findall(f'./dcc:body[@lang="{lang1}"]',ns)
    #         headingLang2 = ms.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
    #         bodyLang2 = ms.findall(f'./dcc:body[@lang="{lang2}"]',ns)
    #         refs = ms.findall('./dcc:ref',ns)
    #         refs = " ".join([elm.text for elm in refs])
    #         tmp = [ headingLang1, bodyLang1, headingLang2, bodyLang2] 
    #         tmp =  [None if i == [] else i[0].text for i in tmp] 
    #         ids = ids+ [refs]+tmp
    #         rng = sht.range((tblRowIdx+1+idx,1))
    #         rng.value = ids
    #         sht.activate()

    def resizeXlTable(self,rng,sht,tableName:str):
        if tableName not in [tbl.name for tbl in sht.tables]:
            sht.tables.add(source=rng, name=tableName)
        else:
            sht.tables[tableName].resize(rng)

    def getHeadingOrBodyFromXlHeadingTag(self, node: et._Element, headingTag:str) -> str: 
        h = headingTag

        lang = h[h.index('[')+1:h.index(']')]
        headOrBody = h.split('[')[0]
        searchStr = f'./dcc:{headOrBody}[@lang="{lang}"]'
        nodes = node.findall(searchStr, node.nsmap)
        if len(nodes) > 0: return nodes[0].text
        else: return None

    def loadDccInfoTable(self, 
                         heading=['in DCC', '@category', '@statementId', 
                                  'heading[en]', 'body[en]', 
                                  'heading[da]', 'body[da]'], 
                         nodeTag="dcc:statements", 
                         subNodeTag="dcc:statement",
                         place_sheet_after='AdministrativeData' ): 
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'


        shtName = nodeTag.replace('dcc:','')    # stratements
        subNodeAttrId = shtName[:-1]+'Id'        # statementId
        if not shtName in wb.sheet_names:
            wb.sheets.add(shtName, after=place_sheet_after)
        sht = wb.sheets[shtName]
        node = root.find(".//"+nodeTag,ns) 

        # Write the Sheet headings
        nodeHeadings = node.findall("dcc:heading",ns)
        sht.range((1,1)).value = [[f'heading[{self.langs[0]}]'], [f'heading[{self.langs[1]}]']]        
        for h in nodeHeadings:
            if h.attrib['lang'] == lang1: sht.range((1,2)).value = h.text
            if h.attrib['lang'] == lang2: sht.range((2,2)).value = h.text

        # Write the table column headings and set column widths
        
        tblRowIdx = 3    
        sht.range((tblRowIdx,1)).value = heading
        sht.range((1,1),(3,len(heading))).columns.autofit()

        hIdxs = [idx for idx,val in enumerate(heading) if val.startswith('heading')]
        bIdxs = [idx for idx,val in enumerate(heading) if val.startswith('body')]
        for i in hIdxs:
            sht.range((1,i+1)).column_width = 30
        for i in bIdxs:
            sht.range((1,i+1)).column_width = 50

        # load the information into the table
        ids = []
        rows = node.findall(subNodeTag,ns)
        issuerIdMap = {'owner_id':'owner', 'manufact_id': 'manufacturer', 'calLab_id': 'calibrationLaboratory'}
        for idx, subNode in enumerate(rows):
            rowData = []
            for i, h in enumerate(heading):
                if i == 0: 
                    rowData.append('y')
                    # if is an attribute
                elif h.startswith('@'): 
                    if h.strip('@') in subNode.attrib.keys():
                        rowData.append(subNode.attrib[h.strip('@')])
                    else: 
                        rowData.append(None)
                elif h.startswith('heading['): 
                    rowData.append(self.getHeadingOrBodyFromXlHeadingTag(subNode, h))
                elif h.startswith('body['):
                    rowData.append(self.getHeadingOrBodyFromXlHeadingTag(subNode, h))
                elif nodeTag == "dcc:equipment" and any([x in issuerIdMap.keys() for x in h.split()]):  # For parsing equipment identifications
                    who, what = h.split(' ')
                    issuer = issuerIdMap[who]
                    subSubNode = subNode.find(f'dcc:identification[@issuer="{issuer}"]', ns)
                    if subSubNode is None: 
                        rowData.append(None)
                    elif what == 'value': 
                        # print('subnode: ',subSubNode.attrib['issuer'])
                        rowData.append(subSubNode.find('./dcc:value', ns).text)
                    elif what.startswith('heading') or what.startswith('body'): 
                        rowData.append(self.getHeadingOrBodyFromXlHeadingTag(subSubNode, what))
                    else:
                        rowData.append(None)
                elif nodeTag == "dcc:measuringSystems":
                    refs = subNode.findall('./dcc:ref',ns)
                    refNodes = [dcchf.getNodeById(root, ref.text) for ref in refs]
                    if h=='equipmentRefs': 
                        # print("EquipmentRefs: ",refNodes)
                        refs = " ".join([ref.text for ref,(tag,node) in zip(refs,refNodes) if tag == "dcc:equipmentItem"])
                        rowData.append(refs)
                    elif nodeTag == "dcc:measuringSystems" and h=='settingRefs': 
                        refs = " ".join([ref.text for ref,(tag,node) in zip(refs,refNodes) if tag == "dcc:setting"])
                        rowData.append(refs)
                    elif nodeTag == "dcc:measuringSystems" and h=='statementRefs': 
                        refs = " ".join([ref.text for ref,(tag,node) in zip(refs,refNodes) if tag == "dcc:statement"])
                        rowData.append(refs)
                    else: 
                        nodes =  subNode.findall(f'./dcc:{h}', ns)
                        if len(nodes)>0: rowData.append(nodes[0].text) 
                        else: rowData.append(None)
                else: 
                    nodes =  subNode.findall(f'./dcc:{h}', ns)
                    if len(nodes)>0: rowData.append(nodes[0].text) 
                    else: rowData.append(None)

            rng = sht.range((tblRowIdx+1+idx,1))
            rng.value = rowData

            if nodeTag == "dcc:measuringSystems": 
                sht.range((1,3)).column_width = 30
        
        rng = sht.range((tblRowIdx,1),(tblRowIdx+1+idx,len(heading)))
        self.resizeXlTable(rng,sht,'Table_'+shtName)
        rng.api.WrapText = True
        rng.columns

        if shtName == "statements": 
            #Apply statement category validator to the statement@category column
            rng = sht.range("Table_"+shtName+"['@category]") 
            self.applyValidationToRange(rng, 'statementCategoryType')
        
        if shtName == "equipment":
            #Apply equipment category type validator to the equipment@category column
            rng = sht.range("Table_"+shtName+"['@category]")
            self.applyValidationToRange(rng, 'equipmentCategoryType')
            # Give a name to the equipmentId column
            equipIdRng = wb.sheets['equipment'].range("Table_equipment['@equipId]")
            equipIdRng.name = "equipIdRange"
        
        if shtName == "settings":
            #Apply equipmentId validator to the setting@refId column
            rng = sht.range("Table_"+shtName+"['@refId]")
            self.applyValidationToRange(rng, 'equipIdRange')

        if shtName == 'measuringSystems': 
            # Give a name to the measurementId column
            measuringSysIdRng = sht.range("Table_"+shtName+"['@measuringSystemId]")
            measuringSysIdRng.name = "measuringSystemIdRange"
            #Apply operationalStatus validator to the measuringSystems@operationalStatus column
            rng = sht.range("Table_"+shtName+"[operationalStatus]")
            self.applyValidationToRange(rng, 'operationalStatusType')


        # Apply Validation
        # validatorMap= {'dcc:statements': }

    def applyValidationToRange(self, rng, restrictionName:str):
            # insert validation on table headings
            formula = '='+restrictionName
            rng.api.Validation.Delete()
            rng.api.Validation.Add(Type=xlValidateList, Formula1=formula) 



    def loadDCCMeasurementResults(self):
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'en'
        lang2 = 'da'

        measurementResults = root.find(".//dcc:measurementResults",ns) 
        print(measurementResults)
        tableIds = [c.attrib["tableId"] for c in measurementResults.getchildren()]
        calibrationResults = measurementResults.findall("dcc:calibrationResult",ns)
        calibResIds = [tbl.attrib['tableId'] for tbl in calibrationResults]
        measurementSeries = measurementResults.findall("dcc:measurementSeries",ns)
        measSerIds = [tbl.attrib['tableId'] for tbl in measurementSeries]

        for tableId in tableIds: 
            if not tableId in wb.sheet_names:
                wb.sheets.add(tableId, after=wb.sheet_names[-1])
        sht = wb.sheets[tableIds[0]]
        sht.activate()

        for tbl in measurementResults.getchildren(): 
            tblType = dcchf.rev_ns_tag(tbl).split(':')[-1]
            tableId = tbl.attrib['tableId']
            # print(f"{tblId} : {tblCategoryType}")
            tabelHeadings = ['tableId', 'tableCategory', 'serviceCategory', 'measuringSystemRef', 'customServiceCategory', 'statementRef','Heading Lang1', 'Heading Lang2', 'numRows', 'numCols']
                        
            sht = wb.sheets[tableId]
            sht.range("A1").value = [[txt] for txt in tabelHeadings]
            sht.range("A1").expand('down').columns.autofit()
            
            idx = tabelHeadings.index('tableCategory')+1
            sht.range((idx,2)).value = tblType

            # insert validation on table headings
            for k in ['tableCategory', 'serviceCategory']:
                i = tabelHeadings.index(k)+1
                rng = sht.range((i,2))
                formula = '='+k+'Type'
                rng.api.Validation.Delete()
                rng.api.Validation.Add(Type=xlValidateList, Formula1=formula) 

            rng = sht.range((tabelHeadings.index('measuringSystemRef')+1,2))
            self.applyValidationToRange(rng, 'measuringSystemIdRange')

            for k, a in tbl.attrib.items(): 
                idx = tabelHeadings.index(k)+1
                sht.range((idx,2)).value = a
            

            colInitRowIdx = idx+2
            numRows = int(tbl.attrib['numRows'])
            numCols = int(tbl.attrib['numCols'])

            # Now load the columns
            columns = tbl.findall("dcc:column", ns)
            columnHeading = ['metaDataCategory', 'scope', 'dataCategory', 'measurand', 'unit', 'headingLang1', 'headingLang2', 'idx \ dataType']
            headingColors = ['light_gray', 'light_yellow', 'light_yellow', 'light_yellow', 'light_green', 'light_blue', 'light_blue', 'light_gray']
            headingColors = [self.colors[k] for k in headingColors]

            sht.range((colInitRowIdx,1)).value = [[h] for h in columnHeading]
            # insert the index column
            rng = sht.range((colInitRowIdx+len(columnHeading),1))
            rng.value = [[i+1] for i in range(numRows)]
            rng = rng.expand('down')
            rng.color = headingColors[0]
            rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

            # 
            for colIdx, col in enumerate(columns):
                # insert the metadata heading attribute values
                cIdx = colIdx + 2
                for k, a in col.attrib.items():
                    rowIdx = colInitRowIdx + columnHeading.index(k)
                    sht.range((rowIdx,cIdx)).value = a 

                # instert the unit
                unit = col.find("dcc:unit",ns).text
                rowIdx = colInitRowIdx + columnHeading.index('unit')
                sht.range(((rowIdx,cIdx))).value = unit

                # Insert human readable heading in two languages. 
                for idx, lang in enumerate(self.langs):
                    xpath =  './/dcc:heading[@lang="{lang}"]'.format(lang=lang)
                    elm = col.find(xpath,ns)
                    if elm is None: 
                        val = None
                    else:
                        val = elm.text
                    rowIdx = colInitRowIdx + columnHeading.index('headingLang1')+idx
                    sht.range(((rowIdx,cIdx))).value = val



                # Insert the data 
                rowIdx = colInitRowIdx + len(columnHeading) - 1
                dataList = col.find('dcc:dataList',ns)
                dataPoints = {int(pt.attrib['idx']): pt.text for pt in dataList}
                if len(dataPoints) > 0:  
                    dataType = dcchf.rev_ns_tag(dataList.getchildren()[0]).strip("dcc:")
                    sht.range((rowIdx,cIdx)).value = dataType
                    for k,v in dataPoints.items(): 
                        sht.range((rowIdx+k,cIdx)).value = v 

            # set colors of the column heading rows
            for i in range(len(columnHeading)):
                idx = colInitRowIdx + i
                rng = sht.range((idx,1)).expand('right')
                rng.color = headingColors[i]
                
            # set Validator for the first four rows in the column heading-rows
            for i in range(4):
                idx = colInitRowIdx + i
                rng = sht.range((idx,2)).expand('right')
                rng.api.Validation.Delete()
                formula = '='+columnHeading[i]+'Type'
                rng.api.Validation.Add(Type=xlValidateList, Formula1=formula) 

            
            # Set the column widths
            idx = colInitRowIdx + len(columnHeading) - 1
            rng = sht.range((colInitRowIdx,1),(colInitRowIdx+4,1)).expand('right')
            rng.columns.autofit()
            rng = sht.range((1,1)).expand()
            rng.columns.autofit()
            # Set the borders visible for the table
            rng = sht.range((colInitRowIdx,1)).expand()
            rng.api.Borders.Weight = 2 
            # Set the color of the data range in the table. 
            rng = sht.range((colInitRowIdx+len(columnHeading),2)).expand()
            rng.color = self.colors["light_red"]


    def write_to_admin(self, sht, root, startline, section) -> int:
        line=startline
        for element in section.iter():
            head=[]
            for child in element.getchildren():
                if dcchf.rev_ns_tag(child) == "dcc:heading":
                    head.append(child.text)
            if len(head) > 0:
                sht.range((line,1)).value = head + ['', '', self.dccTree.getpath(element)]
                sht.range((line,1)).expand('right').color = self.colors["light_yellow"]
                sht.range((line,3)).value = dcchf.rev_ns_tag(element).replace('dcc:','')
                sht.range((line,3)).font.bold = True
                line+=1
            if type(element.text)!=type(None): 
                if dcchf.rev_ns_tag(element)!="dcc:heading": 
                    if not(element.text.startswith('\n')):
                        # print(element)
                        sht.range((line,3)).value = dcchf.rev_ns_tag(element).replace("dcc:",'')
                        rng = sht.range((line,4))
                        rng.value = [element.text, self.dccTree.getpath(element)]
                        rng.color = self.colors["light_yellow"]
                        line+=1
        return line

                
    def loadDCCAdministrativeInformation(self, after='Definitions'):
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'AdministrativeData' in wb.sheet_names:
            wb.sheets.add('AdministrativeData', after=after)
        
        sht = wb.sheets['AdministrativeData']
        sht.clear()
        toprow = ["heading lang1", "heading lang2", "Description", "Value", "XPatht"]
        sht.range((1,1)).value = toprow
        sht.range((1,1)).expand('right').font.bold = True
        head = [h.text for h in root.findall('./dcc:heading', ns)]
        sht.range((2,1)).value = head
        lineIdx = 3
        adm=root.find("dcc:administrativeData", ns)
        soft=adm.find("dcc:dccSoftware", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx, soft)
        core=adm.find("dcc:coreData", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx, core)
        callab=adm.find("dcc:calibrationLaboratory", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx,callab)
        core=adm.find("dcc:accreditation", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx, core)
        cust=adm.find("dcc:customer", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx,cust)
        rng = sht.range((1,1)).expand()
        sht.autofit(axis="columns")

        # heading = ['in DCC', 'settingId', 'refId', 
        #            'value', 'unit',  
        #            'heading lang1', 'body lang1', 
        #            'heading lang2', 'body lang2']

        # sht.range((1,1), (1,7)).value = heading

        # for idx, setting in enumerate(settingList):
        #     inDCC = 'y'
        #     settingRefId = setting.attrib['refId'] if 'refId' in setting.attrib else None
        #     settingId = setting.attrib['settingId']
        #     ids = [inDCC, settingId, settingRefId]
        #     headingLang1 = setting.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
        #     bodyLang1 = setting.findall(f'./dcc:body[@lang="{lang1}"]',ns)
        #     headingLang2 = setting.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
        #     bodyLang2 = setting.findall(f'./dcc:body[@lang="{lang2}"]',ns)
        #     value = setting.findall('dcc:value',ns)
        #     unit = setting.findall('dcc:unit',ns)
        #     tmp = [value, unit, headingLang1, bodyLang1, headingLang2, bodyLang2]
        #     tmp =  [None if i == [] else i[0].text for i in tmp]
        #     ids = ids+tmp
        #     rng = sht.range((idx+2,1))
        #     rng.value = ids            

    def minimal_DCC():
        schemaVersion = self.xsd_root.attrib['version']
        #%%
        nsmap = {'xsi': "http://www.w3.org/2001/XMLSchema-instance", 
                 'dcc': DCC.strip('{}')}
        schemalocation= DCC.strip('{}') + " dcc.xsd"
        xsi="http://www.w3.org/2001/XMLSchema-instance"
        newRoot = xmlEt.Element('digitalCalibrationCertificate', 
                             nsmap=nsmap, 
                             attrib={"schemaVersion":schemaVersion,
                                     "xmlns:xsi":xsi, 
                                     "xsi:schemaLocation":schemalocation})
        #%%
        myNameSpaces = DCC.strip('{}') + " dcc.xsd"
        em = etb.ElementMaker(namespace=myNameSpaces, 
                               nsmap={None: DCC.strip('{}'), 
                                      'dcc' : DCC.strip('{}'), 
                                      'xsi' : "http://www.w3.org/2001/XMLSchema-instance"
                               })
        g = em.root(label="Test", directed="1")
        print(et.tostring(graph, pretty_print=True))
        
        #%%
        from lxml.builder import ElementMaker
        from lxml import etree

        # create an ElementMaker instance with multiple namespaces
        myNameSpace = DCC.strip('{}')
        ns = {'dcc': 'https://dfm.dk',
              'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}
        E = ElementMaker(namespace=myNameSpace, 
                         nsmap=ns)

        # create elements with attributes
        r = E("digitalCalibrationCertificate")
        r.set("{http://www.w3.org/2001/XMLSchema-instance}schemaLocation", myNameSpace+" dcc.xsd")

        child1 = E("administrativeData", name="value")
        child2 = E("{http://www.example2.com}child", name="value2")

        # add the children to the root
        r.append(child1)
        r.append(child2)

        # write the XML to file with pretty print
        with open('output.xml', 'wb') as f:
            f.write(etree.tostring(r, pretty_print=True))


        #%%
        from lxml import etree as ET

        # Define your element names
        element_name = 'root'
        sub_element_name = 'child'
        sub_sub_element_name = 'grandchild'

        # Create the elements
        element = ET.Element(element_name)
        sub_element = ET.SubElement(element, sub_element_name)
        sub_sub_element = ET.SubElement(sub_element, sub_sub_element_name, 
                                        {"{http://www.w3.org/2001/XMLSchema-instance}nil": "true"})
        
        pretty_xml = ET.tostring(element, pretty_print=True)
        print(pretty_xml.decode())

        #%%

    #     newRoot=et.Element(DCC+'digitalCalibrationCertificate',
    #             attrib={"schemaVersion":schemaVersion,   
    #                     "xmlns:xsi":xsi, 
    #                     "xsi:schemaLocation":xsilocation})
    #     return newRoot

    # def storeExcelDataToXML(self):
    #     self.storeAdministrativeDataToEtree()
    #     self.storeStatementsToEtree()
    #     self.storeEquipmentToEtree() 
    #     self.store
        pass

    def storeAdministrativeDataToEtree(self): 
        pass
    
    def storeStatementsToEtree(self):
        pass

    def storeEquipmentToEtree(self):
        pass

    def storeSettingsToEtree(self):
        pass
    
    def storeMeasureingSystemsToEtree(self):
        pass

    def storeDataTablesToEtree(self):
        pass






class MainApp(tk.Tk):



    
    def __init__(self, queryTool: DccQuerryTool):
        super().__init__()
        app = self
        self.queryTool = queryTool
        self.setup_gui(app)

        # self.configure(background='white')
        # self.bind()
        # self.bindVirtualEvents()
        uiExcelFileName = 'DCC_UI_blank.xlsx'
        self.queryTool.loadExcelWorkbook(workBookFilePath=uiExcelFileName)
        # self.label1.config(text='DCC_pipette_blank.xlsx')
        self.label1.config(text=uiExcelFileName)
        # self.queryTool.loadDCCFile('I:\\MS\\4006-03 AI metrologi\\Software\\DCCtables\\master\\Examples\\Stip-230063-V1.xml')
        # self.label2.config(text='Stip-230063-V1.xml')
        dccFileName = 'Examples\\Stip-230063-V1.xml'
        dccFileName = 'SKH_10112_2.xml'
        self.queryTool.loadDCCFile(dccFileName)
        self.label2.config(text=dccFileName)
        self.loadDCCsequence()

    def setup_gui(self,app):
        self.wm_title("DCC EXCEL UI Tool")
        # self.canvas = tk.Canvas(self, width = 1851, height = 1041)
        self.geometry('300x200')
        self.attributes('-topmost', True)
        # img = tk.PhotoImage(file = 'BG_design.png')
        # self.background_image = img
        # self.background_label = tk.Label(app, image=img)
        # self.background_label.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Create three buttons
        button1 = tk.Button(app, text='load excel-file', command=self.loadExcelBook)
        button1.pack(pady=10)
        self.label1 = tk.Label(app, text = "Select a schema file")
        self.label1.pack()

        button2 = tk.Button(app, text='load DCC.xml', command=self.loadDCC)
        button2.pack(pady=10)
        self.label2 = tk.Label(app, text = "Select a DCC file")
        self.label2.pack()

        button3 = tk.Button(app, text='Merge DCC into gui') #, command=self.queryTool.runDccQuery)
        button3.pack(pady=10)

    def loadDCCsequence(self):
            self.statementHeadings = ['in DCC', '@category', '@statementId', 
                                        'heading[en]', 'body[en]', 
                                        'heading[da]', 'body[da]']

            self.equipmentHeadings = ['in DCC', '@equipId', '@category',
                                        'heading[da]', 'heading[en]', 'manufacturer', 'productName', 
                                        'customer_id heading[en]', 'customer_id heading[da]','customer_id value', 
                                        'manufact_id heading[en]', 'manufact_id heading[da]','manufact_id value', 
                                        'calLab_id heading[en]', 'calLab_id heading[da]', 'calLab_id value']
            
            self.settingsHeadings = ['in DCC', '@settingId', '@refId', 
                                    'parameter', 'value', 'unit', 'softwareInstruction', 
                                    'heading[en]', 'body[en]', 
                                    'heading[da]', 'body[da]']
            
            self.measuringSystemsHeadings = ['in DCC', '@measuringSystemId', 
                                            'equipmentRefs', 'settingRefs', 'statementRefs', 
                                            'operationalStatus',
                                            'heading[en]', 'body[en]', 
                                            'heading[da]', 'body[da]']

            self.queryTool.loadDCCAdministrativeInformation()
            
            self.queryTool.loadDccInfoTable(heading = self.statementHeadings, 
                                            nodeTag="dcc:statements",
                                            subNodeTag="dcc:statement",
                                            place_sheet_after='AdministrativeData')
            
            self.queryTool.loadDccInfoTable(heading = self.equipmentHeadings, 
                                            nodeTag="dcc:equipment", 
                                            subNodeTag="dcc:equipmentItem",
                                            place_sheet_after='statements')
            
            self.queryTool.loadDccInfoTable( heading = self.settingsHeadings, 
                                            nodeTag="dcc:settings", 
                                            subNodeTag="dcc:setting",
                                            place_sheet_after='equipment')
            self.queryTool.loadDccInfoTable( heading = self.measuringSystemsHeadings, 
                                            nodeTag="dcc:measuringSystems", 
                                            subNodeTag="dcc:measuringSystem",
                                            place_sheet_after='settings')
            self.queryTool.loadDCCMeasurementResults()
            # self.queryTool.loadDccStatements()
            # self.queryTool.loadDCCEquipment()
            # self.queryTool.loadDCCSettings()
            # self.queryTool.loadDCCMeasurementSystem()

    def loadExcelBook(self):
        file_path = tkfd.askopenfilename(initialdir=os.getcwd())
        self.queryTool.loadExcelWorkbook(workBookFilePath=file_path)
        self.label1.config(text=file_path)


    def loadSchema(self):
        file_path = tkfd.askopenfilename(initialdir=os.getcwd())
        self.queryTool.loadShemaFile(xsdFileName=file_path)
        self.queryTool.loadSchemaRestrictions()
        self.label1.config(text=file_path)

    def loadDCC(self):
        file_path = tkfd.askopenfilename(initialdir=os.getcwd())
        for sht in self.queryTool.wb.sheets: 
            if not sht.name == "Definitions": sht.delete()
        self.queryTool.loadDCCFile(file_path)
        # print(f"Label is: {self.label1['text']}")
        dcchf.validate(file_path, 'dcc.xsd')    
        self.loadDCCsequence()
        self.label2.config(text=file_path)
        
# if __name__ == "__main__":

if __name__=="__main__":
    #################
    #first argument is dcc xml file
    #second argument is excel template to use
    import sys
    args=sys.argv[1:]
    print(len(args))
    

    dccQuerryTool = DccQuerryTool()
    app = MainApp(dccQuerryTool)
    app.mainloop()
    # if len(args)==0:
    #     mapFileName ='Examples'+os.sep+'Mapping_Novo_temperatur_Certifikat.xlsx'
    #     dccFileName = 'Examples'+os.sep+'Stip-230063-V1.xml'
    #     lookupFromMappingFile(mapFileName, dccFileName)
    # elif len(args)==2:
    #     mapFileName = args[0]
    #     dccFileName = args[1]
    #     lookupFromMappingFile(mapFileName, dccFileName)
    # else: 
    #     helpstatement = """call dccquery.py using the following arguments: \n 
    #     >> python dccquery.py [mapping file e.g. mapping.xlsx] [DCC file e.g. dcc.xml] """
    #     print(helpstatement)

         
