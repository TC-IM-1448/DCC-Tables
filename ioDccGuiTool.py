# 
# by DBH 2024-01-11
#
# Uses xlwings see following references: 
# https://docs.xlwings.org/en/stable/syntax_overview.html 
# https://docs.xlwings.org/en/latest/ 
#%%
__ver__ = "DCC-EXCEL-GUI v. 0.0.2"
#%%
import os
import re
import base64
import openpyxl as pyxl
import xlwings as xw
from  lxml import etree as et
from lxml import builder as etb
from lxml.builder import ElementMaker 
import tkinter as tk
import tkinter.filedialog as tkfd
import DCChelpfunctions as dcchf
from DCChelpfunctions import search


LANG='da'
DCC='{https://dfm.dk}'

xlValidateList = xw.constants.DVType.xlValidateList
HEADINGS = dict(statementHeadings = ['in DCC', '@id', '@category', 
                            'heading[en]', 'heading[da]', 
                            'body[en]', 'body[da]', 
                            'externalReference'],

    equipmentHeadings = ['in DCC', '@id', '@category',
                                'heading[da]', 'heading[en]', 'manufacturer', 'productName', 'productType',
                                'customer_id heading[en]', 'customer_id heading[da]','customer_id value', 
                                'manufact_id heading[en]', 'manufact_id heading[da]','manufact_id value', 
                                'calLab_id heading[en]', 'calLab_id heading[da]', 'calLab_id value'],

    settingsHeadings = ['in DCC', '@settingId', '@equipmentRef', 
                            'parameter', 'value', 'unit', 'softwareInstruction', 
                            'heading[en]', 'heading[da]', 'body[en]', 'body[da]'
                            ],

    measuringSystemsHeadings = ['in DCC', '@id', 
                                    'heading[en]', 'heading[da]', 
                                    'equipmentRefs', 'settingRefs', 'statementRefs',
                                    'operationalStatus',
                                    'body[en]', 'body[da]'
                                    ], 

    embeddedFilesHeadings = ['in DCC', '@id', 
                             'heading[da]', 
                             'heading[en]',
                             'body[da]',
                             'body[en]',
                             'fileExtension'],

    administrativeDataHeadings = [ "heading[en]", 
                                        "heading[da]", 
                                        "Description", 
                                        "Value", 
                                        "XPath"], 

    measurementResultHeadings = ['tableCategory', 
                                        '@tableId', 
                                        '@serviceCategory', 
                                        '@measuringSystemRef', 
                                        '@customServiceCategory', 
                                        'statementRef',
                                        'heading[da]', 
                                        'heading[en]', 
                                        '@numRows', 
                                        '@numCols'], 

    columnHeading = ['scope', 'dataCategory', 'dataCategoryRef', 'measurand', 'unit', 'heading[da]', 'heading[en]', 'idx'])

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
        self.createEmbeddedFilesFolder(wb)

    def createEmbeddedFilesFolder(self, wb):
        wbpath = wb.fullname.replace(wb.name, '')
        if not os.path.exists(wbpath+'embeddedFiles'):
            os.mkdir(wbpath+'embeddedFiles')
        self.embeddedFilesPath = wbpath+'embeddedFiles'+os.sep

    def loadSchemaFile(self, xsdFileName="dcc.xsd"):
        self.xsdTree, self.xsdRoot  = dcchf.load_xml(xsdFileName)
        
    def loadDCCFile(self, xmlFileName="SKH_10112_2.xml"):
        self.dccTree, self.dccRoot = dcchf.load_xml(xmlFileName)
        errors = dcchf.validate(xmlFileName, 'dcc.xsd')    
        
        for sht in self.wb.sheets: 
            if not sht.name == "Definitions": sht.delete()
        self.loadSchemaRestrictions()
        self.loadDccSequence()
        return errors
        
    def loadSchemaRestrictions(self): 
        xsd_root = self.xsdRoot
        sht_def = self.sheetDef
        drestr = dcchf.schema_get_restrictions(xsd_root)
        rng = sht_def.range((1,3)).expand()
        rng.clear()
        j = self.xsdDefInitCol
        for i, (k,vs) in enumerate(drestr.items()):
            rng = sht_def.range((1, i+j))
            rng.value = [k]
            rng.offset(1,0).value = [[v] for v in vs]
            rng.expand('down').name = k  
            rng.font.bold = True
        self.dccDefInitCol = i+1

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

    def loadDccSequence(self):
        self.loadDCCAdministrativeInformation(after='Definitions',
                                              heading=HEADINGS['administrativeDataHeadings'])
        
        self.loadDccInfoTable(heading = HEADINGS['statementHeadings'], 
                                nodeTag="dcc:statements",
                                subNodeTag="dcc:statement",
                                place_sheet_after='AdministrativeData')
        
        self.loadDccInfoTable(heading = HEADINGS['equipmentHeadings'], 
                                nodeTag="dcc:equipment", 
                                subNodeTag="dcc:equipmentItem",
                                place_sheet_after='statements')
        
        self.loadDccInfoTable( heading = HEADINGS['settingsHeadings'], 
                                nodeTag="dcc:settings", 
                                subNodeTag="dcc:setting",
                                place_sheet_after='equipment')
        self.loadDccInfoTable( heading = HEADINGS['measuringSystemsHeadings'], 
                                nodeTag="dcc:measuringSystems", 
                                subNodeTag="dcc:measuringSystem",
                                place_sheet_after='settings')
        self.loadDccInfoTable( heading = HEADINGS['embeddedFilesHeadings'], 
                                nodeTag="dcc:embeddedFiles", 
                                subNodeTag="dcc:embeddedFile",
                                place_sheet_after='measuringSystems')
        self.loadDCCMeasurementResults(heading=HEADINGS['measurementResultHeadings'])


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
        issuerIdMap = {'customer_id':'customer', 'manufact_id': 'manufacturer', 'calLab_id': 'calibrationLaboratory'}
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
                else: 
                    nodes =  subNode.findall(f'./dcc:{h}', ns)
                    if len(nodes)>0: rowData.append(nodes[0].text) 
                    else: rowData.append(None)
                if nodeTag == "dcc:embeddedFiles" and h=="@id": 
                    fileId = subNode.attrib[h.strip('@')]
                    filePath = self.embeddedFilesPath+fileId
                    fileData = subNode.find("dcc:fileContent", ns).text
                    fileIsSaved = self.saveEmbeddedFileToFolder(filePath, fileData)
                    cidx = len(heading)+1
                    if fileId.split('.')[-1].lower() in ['png', 'emf', 'jpg'] and fileIsSaved:
                        sht.pictures.add(filePath, name=fileId, anchor=sht.range((4+idx,1+cidx+idx)))
                    sht.range((tblRowIdx+1+idx,cidx)).value = filePath
                    sht.range((1,2)).column_width = 27


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
            equipIdRng = wb.sheets['equipment'].range("Table_equipment['@id]")
            equipIdRng.name = "equipIdRange"
        
        if shtName == "settings":
            #Apply equipmentId validator to the setting@refId column
            rng = sht.range("Table_"+shtName+"['@equipmentRef]")
            self.applyValidationToRange(rng, 'equipIdRange')

        if shtName == 'measuringSystems': 
            # Give a name to the measurementId column
            measuringSysIdRng = sht.range("Table_"+shtName+"['@id]")
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

    def saveEmbeddedFileToFolder(self, filepath, base64_string):
        try:
            image_bytes = base64.b64decode(base64_string)

            with open(filepath, "wb") as img_file:
                img_file.write(image_bytes)
                
            print(f"Image saved to {filepath}")
            return True
        except Exception as e:
            print(f"Error decoding base64 string: {e}")
            return False
        
    def loadDCCMeasurementResults(self, heading=['tableId', 
                                                 'tableCategory', 
                                                 'serviceCategory', 
                                                 'measuringSystemRef', 
                                                 'customServiceCategory', 
                                                 'statementRef',
                                                 'Heading Lang1', 
                                                 'Heading Lang2', 
                                                 'numRows', 
                                                 'numCols']):
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
            tabelHeadings = heading
                        
            sht = wb.sheets[tableId]
            sht.range("A1").value = [[txt] for txt in tabelHeadings]
            sht.range("A1").expand('down').columns.autofit()
            

            for i,h in enumerate(tabelHeadings):
                idx = i+1
                if h=='tableCategory':
                    # idx = tabelHeadings.index('tableCategory')+1
                    rng = sht.range((idx,2))
                    rng.value = tblType
                    self.applyValidationToRange(rng, 'tableCategoryType')
                elif h == '@serviceCategory':
                    rng = sht.range((idx,2))
                    try:
                        rng.value = tbl.attrib['serviceCategory']
                    except KeyError:
                        rng.value = None
                    self.applyValidationToRange(rng,'serviceCategoryType')
                elif h == '@measuringSystemRef': 
                    rng = sht.range((idx,2))
                    rng.value = tbl.attrib['measuringSystemRef']
                    self.applyValidationToRange(rng, 'measuringSystemIdRange')
                elif h.startswith('heading['): 
                    rng = sht.range((idx,2))
                    rng.value = self.getHeadingOrBodyFromXlHeadingTag(tbl, h)
                elif h.startswith('@'): 
                    rng = sht.range((idx,2))
                    if h.strip('@') in tbl.attrib.keys():
                        rng.value = tbl.attrib[h.strip('@')]
                else:
                    rng = sht.range((idx,2)) 
                    nodes =  tbl.findall(f'./dcc:{h}', ns)
                    if len(nodes)>0: 
                        rng.value = nodes[0].text
                    else:
                        rng.value = None

            rng = sht.range((1,2), (idx,2))
            rng.color = self.colors["light_yellow"]
            rng.api.Borders.Weight = 2 
            
            colInitRowIdx = idx+2
            numRows = int(tbl.attrib['numRows'])
            numCols = int(tbl.attrib['numCols'])

            # Now load the columns
            columns = tbl.findall("dcc:column", ns)
            columnHeading = ['dataCategoryRef', 'scope', 'dataCategory', 'measurand', 'unit', 'heading[da]', 'heading[en]', 'idx']
            columnHeading = HEADINGS['columnHeading']
            headingColors = ['yellow', 'yellow', 'light_yellow', 'green', 'light_green', 'light_blue', 'light_blue', 'light_gray']
            headingColors = [self.colors[k] for k in headingColors]

            sht.range((colInitRowIdx,1)).value = [[h] for h in columnHeading]
            # insert the index column
            rng = sht.range((colInitRowIdx+len(columnHeading),1))
            rng.value = [[i+1] for i in range(numRows)]
            rng = rng.expand('down')
            rng.color = headingColors[-1] 
            rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

            # 
            for colIdx, col in enumerate(columns):
                # insert the metadata heading attribute values
                cIdx = colIdx + 2
                for k, a in col.attrib.items():
                    rowIdx = colInitRowIdx + columnHeading.index(k)
                    sht.range((rowIdx,cIdx)).value = a 

                # insert the dataCategory. 
                # dataList = col.find('dcc:dataList',ns)
                rowIdx = colInitRowIdx + columnHeading.index('dataCategory')
                dataCategory = dcchf.rev_ns_tag(col.getchildren()[-1])
                dataCategory = dataCategory.replace("dcc:", "", 1)
                sht.range((rowIdx,cIdx)).value = dataCategory 

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
                    rowIdx = colInitRowIdx + columnHeading.index(f'heading[{lang}]')
                    sht.range(((rowIdx,cIdx))).value = val


                # Insert the data 
                rowIdx = colInitRowIdx + len(columnHeading) - 1
                dataList = col.getchildren()[-1]
                dataPoints = {int(pt.attrib['idx']): pt.text for pt in dataList}
                if len(dataPoints) > 0:  
                    # dataType = dcchf.rev_ns_tag(dataList.getchildren()[0]).strip("dcc:")
                    # sht.range((rowIdx,cIdx)).value = dataType
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
                formula = columnHeading[i] if not columnHeading[i][-3:] == "Ref" else 'dataCategory'
                formula = '='+formula+'Type'
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
            rng = sht.range((colInitRowIdx+len(columnHeading),2),(colInitRowIdx+len(columnHeading)+numRows-1,1+numCols))
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
                for k,v in element.attrib.items(): 
                    sht.range((line,3)).value = "@"+k
                    rng = sht.range((line,4))
                    rng.value = str(v)
                    rng.color = self.colors["light_yellow"]
                    sht.range((line,5)).value = self.dccTree.getpath(element)
                    line+=1
            if type(element.text)!=type(None): 
                if dcchf.rev_ns_tag(element)!="dcc:heading": 
                    if not(element.text.startswith('\n')):
                        # print(element)
                        sht.range((line,3)).value = dcchf.rev_ns_tag(element).replace("dcc:",'')
                        sht.range((line,5)).value = self.dccTree.getpath(element)
                        rng = sht.range((line,4))
                        rng.value = str(element.text)
                        rng.color = self.colors["light_yellow"]
                        line+=1
        return line

                
    def loadDCCAdministrativeInformation(self, after='Definitions', 
                                         heading=["heading[en]", 
                                                  "heading[da]", 
                                                  "Description", 
                                                  "Value", 
                                                  "XPath"]):
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'administrativeData' in wb.sheet_names:
            wb.sheets.add('administrativeData', after=after)
        
        sht = wb.sheets['administrativeData']
        sht.clear()
        toprow = heading
        sht.range((1,1)).value = toprow
        sht.range((1,1)).expand('right').font.bold = True
        hIdx = [i for i, elm in enumerate(heading) if elm.startswith('heading[')]
        titles = [self.getHeadingOrBodyFromXlHeadingTag(root, heading[i]) for i in hIdx]
        
        # head = [h.text for h in root.findall('./dcc:heading', ns)]
        sht.range((2,1)).value = titles
        sht.range((2,1),(2,2)).color = self.colors["light_yellow"]
        sht.range((2,3)).value = ['Certificate-Title', None, '/dcc:digitalCalibrationCertificate']
        sht.range((2,3)).font.bold = True
        lineIdx = 3
        adm=root.find("dcc:administrativeData", ns)
        soft=adm.find("dcc:dccSoftware", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx, soft)
        core=adm.find("dcc:coreData", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx, core)
        callab=adm.find("dcc:calibrationLaboratory", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx,callab)
        respPers=adm.find("dcc:respPersons", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx,respPers)
        accr=adm.find("dcc:accreditation", ns)
        lineIdx = self.write_to_admin(sht, root, lineIdx, accr)
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

def extractHeadingLang(s):
    r = re.findall(r'\[(.*?)\]', s)
    if len(r) == 1:
        return r[0]
    else:
        return re.findall(r'\[(.*?)\]', s)
#%%    

def exportToXmlFile(wb, fileName='output.xml'):
    # create an ElementMaker instance with multiple namespaces
    myNameSpace = DCC.strip('{}')
    ns = {'dcc': 'https://dfm.dk',
            'xsi': 'http://www.w3.org/2001/XMLSchema-instance'}
    elmMaker = ElementMaker(namespace=myNameSpace, 
                        nsmap=ns)

    # create the root element of the output tree  with attributes
    exportRoot = elmMaker("digitalCalibrationCertificate", schemaVersion="3.0.0")
    exportRoot.set("{http://www.w3.org/2001/XMLSchema-instance}schemaLocation", myNameSpace+" dcc.xsd")

    # wb = xw.Book('DCC_pipette_blank.xlsx')
    
    adminSht = wb.sheets['administrativeData']
    adminHeading = HEADINGS['administrativeDataHeadings']
    
    colIdxXpath = adminHeading.index('XPath')+1
    rowInitXpath = 2
    rngXpath = adminSht.range((rowInitXpath,colIdxXpath)).expand('down')
    rngDescription = adminSht.range((rowInitXpath,adminHeading.index('Description'))).expand('down')
    descriptions = [c.value for c in rngDescription]

    def exportHeading(node, elmMaker, xlSheet, sheetHeading, rowIdx): 
        headingColIdxLang = [(idx, extractHeadingLang(h)) for idx,h in enumerate(sheetHeading) if h.startswith('heading[')]
        headingText = [(xlSheet.range((rowIdx,idx+1)).value, lang) for idx, lang in headingColIdxLang]
        for txt,lang in headingText:
            if not txt == None:
                elm = elmMaker("heading", txt, lang=lang)
                node.append(elm)

    def add_subtree(root, xpath_list: list):
        for idx, xpath in enumerate(xpath_list):
            rowIdx = idx + rowInitXpath
            data = adminSht.range((idx+rowInitXpath,colIdxXpath-1)).api.Text
            data = "" if data is None else str(data)
            current_node = root
            xpathTags = xpath.split('/')
            for i, node in enumerate(xpathTags):
                next_node = current_node.find(node, ns)
                # print(node, next_node)
                if next_node is None and i < len(xpathTags)-1:
                    next_node = elmMaker(node.replace('dcc:','')) 
                    current_node.append(next_node)
                elif next_node is None:
                    tag = node.replace('dcc:','')
                    print(f'tag is: {tag}  data:{data} ', current_node)
                    if not data == None:
                        next_node = elmMaker(tag, data)
                        current_node.append(next_node)
                current_node = next_node
            if adminSht.range((rowIdx,colIdxXpath-2)).font.bold :
                exportHeading(current_node, elmMaker, adminSht, adminHeading, rowIdx) 
                
    s = '/dcc:digitalCalibrationCertificate'
    xpath_list = [xpth.value.replace(s,'.') for xpth in rngXpath]
    add_subtree(exportRoot, xpath_list)
    # dcchf.print_node(exportRoot)
    mainSignerNode = exportRoot.find('./dcc:administrativeData/dcc:respPersons/dcc:respPerson/dcc:mainSigner', ns)
    if mainSignerNode != None: 
        mainSignerNode.text = mainSignerNode.text.lower()

    

    # set TRUE/FALSE to true/false

    adminNode = exportRoot.find('./dcc:administrativeData', ns)
    print(adminNode)
    dcchf.print_node(exportRoot)

    exportInfoTable(exportRoot, elmMaker, wb, nodeName = 'statements')
    exportEquipment(exportRoot, elmMaker, wb)
    exportInfoTable(exportRoot, elmMaker, wb, nodeName = 'settings')
    exportInfoTable(exportRoot, elmMaker, wb, nodeName = 'measuringSystems')
    msIdx = wb.sheet_names.index('embeddedFiles')+1
    measurementResultsNode = elmMaker('measurementResults')
    exportRoot.append(measurementResultsNode)
    for tblId in wb.sheet_names[msIdx:]:
        print(tblId)
        exportDataTable(measurementResultsNode, elmMaker, wb, tblId)
    
    efNode = exportInfoTable(exportRoot, elmMaker, wb, nodeName='embeddedFiles')
    exportEmbeddedFiles(exportRoot, elmMaker, wb, efNode)

    # measurementResultsNode = elmMaker("measurementResults", name="test")

    # write the XML to file with pretty print
    with open(fileName, 'wb') as f:
        # xml_str = et.tostring(exportRoot, pretty_print=True, xml_declaration=True, encoding='utf-8').decode()
        # xml_str_crlf  = xml_str.replace('\n', '\r\n')
        # f.write(xml_str_crlf.encode())
        f.write(et.tostring(exportRoot, pretty_print=True, xml_declaration=True, encoding='utf-8'))

    dcchf.validate(fileName, 'dcc.xsd')

def exportSheetHeading(parentNode, sht, elmMaker): 
    for i in range(1,3): 
        head = sht.range((i,1)).value
        lang = extractHeadingLang(head)
        txt = sht.range((i,2)).value if not sht.range((i,2)).value == None else "" 
        elm = elmMaker("heading", txt, lang=lang)
        parentNode.append(elm)

def exportEmbeddedFiles(exportRoot, elmMaker, wb, embeddedFilesNode):
    ns = exportRoot.nsmap
    shtIdx = wb.sheet_names.index('embeddedFiles')
    sht = wb.sheets[shtIdx]
    
    def encode_file_to_base64(file_path):
        try:
            with open(file_path, "rb") as img_file:
                file_bytes = img_file.read()
                base64_encoded = base64.b64encode(file_bytes).decode("utf-8")
                return base64_encoded
        except FileNotFoundError:
            print(f"Error: File '{file_path}' not found.")
            return None

    tbl = sht.tables['Table_embeddedFiles']      
    rng_header = tbl.data_body_range
    nrow, ncols = rng_header.shape
    print(rng_header.shape)
    dcchf.print_node(embeddedFilesNode)

    for i, efn in enumerate(embeddedFilesNode.findall("./dcc:embeddedFile",ns)): 
        filepath = sht.range((4+i,ncols+1)).value
        print(f"loading file: {filepath}")
        base64str = encode_file_to_base64(filepath)
        node = elmMaker('fileContent',base64str)
        efn.append(node)


def exportEquipment(adminNode, elmMaker,wb): 
    ns = adminNode.nsmap
    shtIdx = wb.sheet_names.index('equipment')
    sht = wb.sheets[shtIdx]
    equipmentNode = elmMaker.equipment()

    exportSheetHeading(equipmentNode, sht, elmMaker)
    ns = equipmentNode.nsmap
    tbl = sht.tables['Table_equipment']
    rng = tbl.data_body_range
    rng_header = tbl.header_row_range
    headings = rng_header.value
    nrow, ncols = rng.shape
    tbl_data = rng.value
    # print(rng.value) 
    for row in tbl_data: 
        if row[0].startswith('y'):
            a = {h[1:]:  row[i] for i,h in enumerate(headings) if h.startswith('@')}
            node = elmMaker('equipmentItem', **a)
            for i, h in enumerate(headings):
                if i == 0 or row[i] == None: 
                    continue
                if h.startswith('@'):
                    continue
                elif h.startswith('heading['): 
                    lang = extractHeadingLang(h)
                    elm = elmMaker('heading', str(row[i]),lang=lang)
                    node.append(elm)
                elif h.startswith('body['):
                    lang = extractHeadingLang(h)
                    elm = elmMaker('body', str(row[i]),lang=lang)
                    node.append(elm)
                elif len(h.split('_id ')) > 1: 
                    issuer, what = h.split('_id ')
                    if issuer.startswith('manufac'): issuer = 'manufacturer'
                    if issuer.startswith('calLab'): issuer = 'calibrationLaboratory'
                    idNode = node.find(f'./dcc:identification[@issuer="{issuer}"]', ns)
                    if idNode is None:
                        idNode = elmMaker('identification',issuer=issuer)
                        node.append(idNode)
                    if what.startswith('heading'): 
                        # print(what)
                        lang = extractHeadingLang(what)
                        elm = elmMaker('heading',row[i], lang=lang)
                        idNode.append(elm)
                    else: 
                        elm = elmMaker('value',str(row[i]))
                        idNode.append(elm)
                else: 
                    elm = elmMaker(h,str(row[i]))
        equipmentNode.append(node)
    adminNode.append(equipmentNode)

def exportInfoTable(adminNode, elmMaker,wb, nodeName = 'settings'): 
    shtIdx = wb.sheet_names.index(nodeName)
    sht = wb.sheets[shtIdx]
    statementsNode = elmMaker(nodeName)

    exportSheetHeading(statementsNode, sht, elmMaker)

    tbl = sht.tables['Table_'+nodeName]
    rng = tbl.data_body_range
    rng_header = tbl.header_row_range
    headings = rng_header.value
    nrow, ncols = rng.shape
    tbl_data = rng.value
    if nrow == 1: 
        tbl_data = [tbl_data]
    for row in tbl_data:
        if row[0].startswith('y'):
            a = {h[1:]:  row[i] for i,h in enumerate(headings) if h.startswith('@')}
            a = {k:v for k,v in a.items() if v != None}
            node = elmMaker(nodeName[:-1], **a)
            for i, h in enumerate(headings):
                if i == 0 or row[i] == None: 
                    continue
                if h.startswith('@'):
                    continue
                elif h.startswith('heading['): 
                    lang = extractHeadingLang(h)
                    elm = elmMaker('heading', str(row[i]),lang=lang)
                elif h.startswith('body['):
                    lang = extractHeadingLang(h)
                    elm = elmMaker('body', str(row[i]),lang=lang)
                else: 
                    elm = elmMaker(h,str(row[i]))
                node.append(elm)
        statementsNode.append(node)
    adminNode.append(statementsNode)
    return statementsNode

def exportDataTable(parentNode, elmMaker, wb, tableId): 
    shtIdx = wb.sheet_names.index(tableId)
    sht = wb.sheets[shtIdx]
    heading = [c.value for c in sht.range((1,1)).expand('down')]
    values = [c.value for c in sht.range((1,2),(len(heading),2))]
    nodeTag = values[0]

    attrib = {}
    for i,h in enumerate(heading): 
        if h.startswith('@') and values[i] != None: 
            val = values[i]
            if h[1:].startswith('num'): 
                val = int(val)
            attrib[h[1:]] = str(val)
    # print(attrib)
    tblNode = elmMaker(nodeTag, **attrib)

    for i,h in enumerate(heading): 
        if h.startswith('heading['):
            lang = extractHeadingLang(h)
            elm = elmMaker('heading', values[i],lang=lang)
            tblNode.append(elm)

    ncols = sht.range((len(heading)+2,1)).expand('right')
    colIdxs = range(2,len(ncols)+1)
    # print(list(colIdxs))
    for colIdx in colIdxs:
        exportDataColumn(tblNode, sht, elmMaker, wb, len(heading)+2, colIdx)

    parentNode.append(tblNode)

def exportDataColumn(parentNode, tblSheet, elmMaker, wb, rowInitIdx, colIdx): 
    typecast_dict = {'int': int, 'real': float, 'string': str, 'bool': bool, 'conformityStatus': str, 'ref': str}
    numRows = int(parentNode.attrib['numRows'])
    # print('numRows = ', numRows)
    colHeading = HEADINGS['columnHeading']
    dataInitRowIdx = rowInitIdx+len(colHeading)-1
    colAttrRange = tblSheet.range((rowInitIdx,colIdx),(dataInitRowIdx, colIdx)) 
    colAttrNameRange = tblSheet.range((rowInitIdx,1),(dataInitRowIdx, 1)) 
    colDataRange = tblSheet.range((dataInitRowIdx+1,colIdx),(dataInitRowIdx+numRows, colIdx)) 
    colIndexRange = tblSheet.range((dataInitRowIdx+1,1),(dataInitRowIdx+numRows, 1)) 
    
    
    colAttrNames = [c.value for c in colAttrNameRange]
    colAttrValues = [c.value for c in colAttrRange]
    colData = [c.api.Text for c in colDataRange]
    colIndex = [int(c.value) for c in colIndexRange]

    colHeadDict = dict(zip(colAttrNames,colAttrValues))

    colAttrKeys = ['scope', 'dataCategoryRef', 'measurand'] 
    colAttr = {k: colHeadDict[k] for k in colAttrKeys}
    # colNode = elmMaker('column', **dict(zip(colAttrNames[:4], colAttrValues[:4])))
    colNode = elmMaker('column', **colAttr)

    colHeadingDict = {k: colHeadDict[k] for k in colAttrNames if k.startswith('heading')}
    for k,v in colHeadingDict.items():
        if not v == None:
            lang = extractHeadingLang(k)
            elm = elmMaker('heading', v,lang=lang)
            colNode.append(elm)
    # I AM HERE
    unitNode = elmMaker('unit', colHeadDict['unit'])
    colNode.append(unitNode)

    # print(colAttrNames)
    # print(colAttrValues)
    # print(colHeading)
    # print(colIndex)
    # print(colData)
    # print(len(colData))
    # dataType = colAttrValues[-1]
    # typecast = typecast_dict[dataType]
    # dataList = elmMaker('dataList')
    dataCategoryNode = elmMaker(colHeadDict['dataCategory'])
    # dataList.append(dataCategoryNode)
    


    
    for i, data in enumerate(colData):
        if not data == "":
            txt = str(data)
            # print(txt)
            # elm = elmMaker(dataType, txt, idx=str(colIndex[i]))
            elm = elmMaker('row', txt, idx=str(colIndex[i]))
            dataCategoryNode.append(elm)
    # colNode.append(dataList)
    colNode.append(dataCategoryNode)
    parentNode.append(colNode)

# exportToXmlFile()    





class MainApp(tk.Tk):

    def __init__(self, queryTool: DccQuerryTool):
        super().__init__()
        app = self
        self.queryTool = queryTool
        self.setup_gui(app)

        # self.configure(background='white')
        # self.bind()
        # self.bindVirtualEvents()
        uiExcelFileName = 'DCC_pipette_blank.xlsx'
        uiExcelFileName = 'DCC_UI_blank.xlsx'
        self.queryTool.loadExcelWorkbook(workBookFilePath=uiExcelFileName)
        # self.label1.config(text='DCC_pipette_blank.xlsx')
        self.label1.config(text=uiExcelFileName)
        # self.queryTool.loadDCCFile('I:\\MS\\4006-03 AI metrologi\\Software\\DCCtables\\master\\Examples\\Stip-230063-V1.xml')
        # self.label2.config(text='Stip-230063-V1.xml')
        dccFileName = 'Examples\\Stip-230063-V1.xml'
        dccFileName = 'Examples\\Template_TemperatureCal.xml'
        dccFileName = 'SKH_10112_2.xml'
        #self.queryTool.loadDCCFile(dccFileName)
        #self.label2.config(text=dccFileName)
        # self.loadDCCsequence()
        # exportToXmlFile('output.xml')

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

        button3 = tk.Button(app, text='Export DCC from GUI', command=self.exportDCC) #, command=self.queryTool.runDccQuery)
        button3.pack(pady=10)

    def loadDCCsequence(self):
        self.queryTool.loadDccSequence()


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
        errors = self.queryTool.loadDCCFile(file_path)
        if not errors: 
            self.label2.config(text=file_path)
        else: 
            self.label2.config(text="File does not validate!")
        
    def exportDCC(self):
        # file_path = tkfd.asksaveasfilename(initialdir=os.getcwd())
        file_path = "output.xml"
        self.label2.config(text="EXPORTING!")
        exportToXmlFile(self.queryTool.wb, file_path)
        self.label2.config(text=f"exported to: {file_path}")
        
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

         
