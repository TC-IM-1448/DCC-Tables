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
import tkinter as tk
import tkinter.filedialog as tkfd
import DCChelpfunctions as dcchf
from DCChelpfunctions import search

#%%
LANG='da'
DCC='{https://dfm.dk}'
xlValidateList = xw.constants.DVType.xlValidateList


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
        self.loadDccStatements()
        
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

    def loadDccStatements(self):
        root = self.dccRoot
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'Statements' in wb.sheet_names:
            wb.sheets.add('Statements', after='Definitions')
        sht = wb.sheets['Statements']
        statements = dcchf.get_statements(root)
        heading = ['in DCC', 'category', 'id', 'heading lang1', 'body lang1', 'heading lang2', 'body lang2']
        sht.range((1,1), (1,7)).value = heading
        ns = root.nsmap

        statementIds = [elm.attrib['statementId'] for elm in dcchf.get_statements(root)]
        statementCat = [elm.attrib['category'] for elm in dcchf.get_statements(root)]
        for idx, statement in enumerate(statements):
            inDCC = 'y'
            statementId = statement.attrib['statementId']
            category = statement.attrib['category']
            lang1Heading = statement.find(f'./dcc:heading[@lang="{lang1}"]',ns).text
            lang1Body = statement.find(f'./dcc:body[@lang="{lang1}"]',ns).text
            lang2Heading = statement.find(f'./dcc:heading[@lang="{lang2}"]',ns).text
            lang2Body = statement.find(f'./dcc:body[@lang="{lang2}"]',ns).text
            rng = sht.range((idx+2,1))
            rng.value = [inDCC, category, statementId, lang1Heading, lang1Body, lang2Heading, lang2Body]
        
        rng = sht.range((2,2),(1024,2))
        rng.api.Validation.Delete()
        rng.api.Validation.Add(Type=xlValidateList, Formula1='=statementCategoryType')

    def loadDCCEquipment(self):
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'Equipment' in wb.sheet_names:
            wb.sheets.add('Equipment', after='Statements')
        sht = wb.sheets['Equipment']
        equipment = root.find(".//dcc:equipment",ns) 
        equipItems = equipment.getchildren()

        heading = ['in DCC', 'category', 'equipId', 'heading lang1', 'heading lang2', 'manufacturer', 'productName', 
                    'id1 value', 'id1 issuer', 'id1 heading lang1', 'id1 heading lang2', 
                    'id2 value', 'id2 issuer', 'id2 heading lang1', 'id2 heading lang2' ]   

        sht.range((1,1), (1,7)).value = heading

        equipId = [elm.attrib['equipId'] for elm in equipItems]
        equipCat = [elm.attrib['category'] for elm in equipItems]
        for idx, equip in enumerate(equipItems):
            inDCC = 'y'
            category = equip.attrib['category']
            equipId = equip.attrib['equipId']
            ids = [inDCC, category, equipId]
            # dcchf.print_node(equip)
            headingLang1 = equip.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
            headingLang2 = equip.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
            productName = equip.findall('dcc:productName',ns)
            manufacturer = equip.findall('dcc:manufacturer',ns)
            tmp = [headingLang1, headingLang2, manufacturer, productName]
            tmp =  [None if i == [] else i[0].text for i in tmp]
            ids = ids+tmp
            # print(ids)
            identifications = equip.findall('.//dcc:identification',ns)
            for Id in identifications:
                issuer = Id.attrib['issuer']
                lang1Heading = Id.find(f'./dcc:heading[@lang="{lang1}"]',ns).text
                lang2Heading = Id.find(f'./dcc:heading[@lang="{lang2}"]',ns).text
                idValue = Id.find('./dcc:value',ns).text
                ids.extend([idValue, issuer, lang1Heading,lang2Heading])
            rng = sht.range((idx+2,1))
            rng.value = ids
        
        rng = sht.range((2,2),(1024,2))
        rng.api.Validation.Delete()
        rng.api.Validation.Add(Type=xlValidateList, Formula1='=equipmentCategoryType')
        rng = sht.range((2,9),(1024,9))

        rng.api.Validation.Delete()
        rng.api.Validation.Add(Type=xlValidateList, Formula1='=issuerType')
        sht.activate()

    def loadDCCSettings(self): 
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'Settings' in wb.sheet_names:
            wb.sheets.add('Settings', after='AdministrativeInfo')
        sht = wb.sheets['Settings']
        settings = root.find(".//dcc:settings",ns) 
        settingList = settings.getchildren()

        heading = ['in DCC', 'settingId', 'refId', 
                   'value', 'unit',  
                   'heading lang1', 'body lang1', 
                   'heading lang2', 'body lang2']

        sht.range((1,1), (1,7)).value = heading

        for idx, setting in enumerate(settingList):
            inDCC = 'y'
            settingRefId = setting.attrib['refId'] if 'refId' in setting.attrib else None
            settingId = setting.attrib['settingId']
            ids = [inDCC, settingId, settingRefId]
            headingLang1 = setting.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
            bodyLang1 = setting.findall(f'./dcc:body[@lang="{lang1}"]',ns)
            headingLang2 = setting.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
            bodyLang2 = setting.findall(f'./dcc:body[@lang="{lang2}"]',ns)
            value = setting.findall('dcc:value',ns)
            unit = setting.findall('dcc:unit',ns)
            tmp = [value, unit, headingLang1, bodyLang1, headingLang2, bodyLang2]
            tmp =  [None if i == [] else i[0].text for i in tmp]
            ids = ids+tmp
            rng = sht.range((idx+2,1))
            rng.value = ids

        rng = sht.range((1,1)).expand('table')
        self.resizeXlTable(rng,sht,'TableSettings')
        rng.api.WrapText = True
        rng = sht.range((1,3)).expand('table')
        rng.columns

    def resizeXlTable(self,rng,sht,tableName):
        if 'TableSettings' not in [tbl.name for tbl in sht.tables]:
            sht.tables.add(source=rng, name='TableSettings')
        else:
            sht.tables['TableSettings'].resize(rng)

        
    def loadDCCMeasurementSystem(self):
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'MeasuringSystems' in wb.sheet_names:
            wb.add('MeasuringSystems', after='Settings')
        sht = wb.sheets['MeasuringSystems']
        msuc = root.find(".//dcc:measuringSystems",ns) 

        heading = ['in DCC', 'Id', 'Instrument & Settings Refs', 'headingLang1', 'bodyLang1', 'headingLang2', 'bodyLang2']

        ids = []
        mssHeadings = msuc.findall("dcc:heading",ns)
        for h in mssHeadings:
            if h.attrib['lang'] == lang1: sht.range((1,2)).value = h.text
            if h.attrib['lang'] == lang2: sht.range((2,2)).value = h.text

        tblRowIdx = 3    
        sht.range((tblRowIdx,1)).value = heading

        msList = msuc.findall("./dcc:measuringSystem",ns)
        for idx, ms in enumerate(msList):
            inDCC = 'y'
            msId = ms.attrib['measuringSystemId']
            ids = [inDCC, msId]
            headingLang1 = ms.findall(f'./dcc:heading[@lang="{lang1}"]',ns)
            bodyLang1 = ms.findall(f'./dcc:body[@lang="{lang1}"]',ns)
            headingLang2 = ms.findall(f'./dcc:heading[@lang="{lang2}"]',ns)
            bodyLang2 = ms.findall(f'./dcc:body[@lang="{lang2}"]',ns)
            refs = ms.findall('./dcc:ref',ns)
            refs = " ".join([elm.text for elm in refs])
            tmp = [ headingLang1, bodyLang1, headingLang2, bodyLang2] 
            tmp =  [None if i == [] else i[0].text for i in tmp] 
            ids = ids+ [refs]+tmp
            rng = sht.range((tblRowIdx+1+idx,1))
            rng.value = ids
            sht.activate()

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

                
    def loadDCCAdministrativeInformation(self):
        root = self.dccRoot
        ns = root.nsmap
        wb = self.wb
        lang1 = 'da'
        lang2 = 'en'

        if not 'AdministrativeInfo' in wb.sheet_names:
            wb.sheets.add('AdministrativeInfo', after='Definitions')
        
        sht = wb.sheets['AdministrativeInfo']
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


    def storeExcelDataToXML(self): 
        pass

    # def runDccQuery(self):
    #     sht = self.sheetMap
    #     root = self.dccRoot
    #     vals = sht.range("A1").expand("down").value
    #     [print(v) for v in vals]
    #     n_rows = vals.index("--END--")+1
    #     print(n_rows)

    #     for i in range(1, n_rows):
    #         queryType = sht.range((i,3)).value

    #         if queryType == 'xpath':
    #             xpath_str = sht.range((i,4)).value
    #             val = dcchf.xpath_query(root, xpath_str)
    #             print(vals[i-1], queryType, val)
    #             if len(val)>0:
    #                 sht[f"M{i}"].value = val[0].text
    #             else:
    #                 sht[f"M{i}"].value = "ERROR not Found"
    #         elif queryType == 'data':
    #             dtbl = dict(zip(["measuringSystemRef", "tableId"], sht.range((i,5),(i,6)).value))
    #             dcol = dict(zip(["metaDataCategory", "scope", "dataCategory","measurand"], sht.range((i,7), (i,10)).value))
    #             unit = sht.range((i,11)).value
    #             customerTag = sht.range((i,12)).value
    #             if customerTag is None:
    #                 data = search(root, dtbl, dcol, unit)
    #             else: 
    #                 data = search(root, dtbl, dcol, unit, rowTags=[customerTag])
    #             print(dtbl, dcol, unit, customerTag, ":", data)
    #             if len(data) == 0 : 
    #                 data_val = "ERROR not Found"
    #             else:
    #                 data_val = list(data.values())
    #             # elif len(data) > 1: 
    #                 # data_val = "ERROR too many values"
    #             sht[f"M{i}"].value = data_val





class MainApp(tk.Tk):
    def __init__(self, queryTool: DccQuerryTool):
        super().__init__()
        app = self
        self.queryTool = queryTool
        self.setup_gui(app)

        # self.configure(background='white')
        # self.bind()
        # self.bindVirtualEvents()
        self.queryTool.loadExcelWorkbook(workBookFilePath='DCC_pipette_blank.xlsx')
        self.label1.config(text='DCC_pipette_blank.xlsx')
        # self.queryTool.loadDCCFile('I:\\MS\\4006-03 AI metrologi\\Software\\DCCtables\\master\\Examples\\Stip-230063-V1.xml')
        # self.label2.config(text='Stip-230063-V1.xml')
        self.queryTool.loadDCCFile('SKH_10112_2.xml')
        self.label2.config(text='SKH_10112_2.xml')
        self.loadDCCprocedure()

    def setup_gui(self,app):
        self.wm_title("DCC EXCEL GUI Tool")
        # self.canvas = tk.Canvas(self, width = 1851, height = 1041)
        self.geometry('500x200')
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

    def loadDCCprocedure(self):
            self.queryTool.loadDCCAdministrativeInformation()
            # self.queryTool.loadDCCSettings()
            self.queryTool.loadDCCEquipment()
            # self.queryTool.loadDCCMeasurementSystem()
            # self.queryTool.loadDCCMeasurementResults()

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
        self.queryTool.loadDCCFile(file_path)
        # dcchf.validate(file_path, self.label1)    
        self.loadDCCprocedure()
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

         
