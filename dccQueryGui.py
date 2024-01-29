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

#%%
def lookupFromMappingFile(mapFileName:str, dccFileName:str):
    """LookupFromMappingFile """
    pass


class DccQuerryTool(): 
    xsdDefInitCol = 3  
    def loadExcelWorkbook(self, workBookFilePath: str):
        self.wb = xw.Book(workBookFilePath)
        wb = xw.Book(workBookFilePath)
        self.sheetDef = wb.sheets['Definitions']
        self.sheetMap = wb.sheets['Mapping']
        self.wb = wb
        self.loadSchemaFile()
        self.loadSchemaRestrictions()
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
        
        j = 7
        sht = self.sheetMap
        for r in ['metaDataCategoryType', 'scopeType','dataCategoryType', 'measurandType']: 
            rng = sht.range((2,j),(1024,j))
            rng.api.Validation.Delete()
            fml = "="+r
            rng.api.Validation.Add(Type=xlValidateList, Formula1=fml)
            j += 1


    def loadDccAttributes(self):
        root = self.dccRoot
        sht_def = self.sheetDef
        sht_map = self.sheetMap

        i = self.dccDefInitCol
        statementIds = [elm.attrib['statementId'] for elm in dcchf.get_statements(root)]
        sht_def.range((1,i)).value = 'statementId'
        sht_def.range((2,i)).value = [[s] for s in statementIds]
        sht_def.range((2,i)).expand('down').name = 'statementId' 
        
        i+=1 
        msucIds = [elm.attrib['measuringSystemId'] for elm in dcchf.get_measuringSystems(root)]
        sht_def.range((1,i)).value = 'measuringSystemId'
        sht_def.range((2,i)).value = [[ms] for ms in msucIds]
        sht_def.range((2,i)).expand('down').name = 'measuringSystemId'
        rng = sht_map.range((2,5),(1024,5))
        rng.api.Validation.Delete()
        rng.api.Validation.Add(Type=xlValidateList, Formula1='=measuringSystemId')

        i+=1
        tableIds = [elm.attrib['tableId'] for elm in dcchf.getTables(root)]
        sht_def.range((1,i)).value = 'tableId'
        sht_def.range((2,i)).value = [[tbl] for tbl in tableIds]
        sht_def.range((2,i)).expand('down').name = 'tableId'
        rng = sht_map.range((2,6),(1024,6))
        rng.api.Validation.Delete()
        rng.api.Validation.Add(Type=xlValidateList, Formula1='=tableId') 
        
    def runDccQuery(self):
        sht = self.sheetMap
        root = self.dccRoot
        vals = sht.range("A1").expand("down").value
        [print(v) for v in vals]
        n_rows = vals.index("--END--")+1
        print(n_rows)

        for i in range(1, n_rows):
            queryType = sht.range((i,3)).value

            if queryType == 'xpath':
                xpath_str = sht.range((i,4)).value
                val = dcchf.xpath_query(root, xpath_str)
                print(vals[i-1], queryType, val)
                if len(val)>0:
                    sht[f"M{i}"].value = val[0].text
                else:
                    sht[f"M{i}"].value = "ERROR not Found"
            elif queryType == 'data':
                dtbl = dict(zip(["measuringSystemRef", "tableId"], sht.range((i,5),(i,6)).value))
                dcol = dict(zip(["metaDataCategory", "scope", "dataCategory","measurand"], sht.range((i,7), (i,10)).value))
                unit = sht.range((i,11)).value
                customerTag = sht.range((i,12)).value
                if customerTag is None:
                    data = search(root, dtbl, dcol, unit)
                else: 
                    data = search(root, dtbl, dcol, unit, rowTags=[customerTag])
                print(dtbl, dcol, unit, customerTag, ":", data)
                if len(data) == 0 : 
                    data_val = "ERROR not Found"
                else:
                    data_val = list(data.values())
                # elif len(data) > 1: 
                    # data_val = "ERROR too many values"
                sht[f"M{i}"].value = data_val



#%%
if False: 
    xsd_tree, xsd_root  = dcchf.load_xml("dcc.xsd")
    tree, root = dcchf.load_xml("SKH_10112_2.xml")
    wb = xw.Book('SKH_10112_2_Mapping.xlsx')
    sht = wb.sheets['Mapping']
    sht_def = wb.sheets['Definitions']

    #%% Load schema definitions 
    drestr = dcchf.schema_get_restrictions(xsd_root)
    for i, (k,vs) in enumerate(drestr.items()):
        sht_def.range((1, i+3)).value = [k]
        sht_def.range((2, i+3)).value = [[v] for v in vs]
        sht_def.range((2,i+3)).expand('down').name = k  
        
    #%% Get DCC attribute names and set Validators for tables and measuringSystems
    xlValidateList = xw.constants.DVType.xlValidateList
    statementIds = [elm.attrib['statementId'] for elm in dcchf.get_statements(root)]
    sht_def.range((1,i+4)).value = 'statementId'
    sht_def.range((2,i+4)).value = [[s] for s in statementIds]
    sht_def.range((2,i+4)).expand('down').name = 'statementId'  

    msucIds = [elm.attrib['measuringSystemId'] for elm in dcchf.get_measuringSystems(root)]
    sht_def.range((1,i+5)).value = 'measuringSystemId'
    sht_def.range((2,i+5)).value = [[ms] for ms in msucIds]
    sht_def.range((2,i+5)).expand('down').name = 'measuringSystemId'
    rng = sht.range((2,5),(1024,5))
    rng.api.Validation.Delete()
    rng.api.Validation.Add(Type=xlValidateList, Formula1='=measuringSystemId')

    tableIds = [elm.attrib['tableId'] for elm in dcchf.getTables(root)]
    sht_def.range((1,i+6)).value = 'tableId'
    sht_def.range((2,i+6)).value = [[tbl] for tbl in tableIds]
    sht_def.range((2,i+6)).expand('down').name = 'tableId'
    rng = sht.range((2,6),(1024,6))
    rng.api.Validation.Delete()
    rng.api.Validation.Add(Type=xlValidateList, Formula1='=tableId')

    #%% get schema restrictions and set validations 
    j = 7
    for r in ['metaDataCategoryType', 'scopeType','dataCategoryType', 'measurandType']: 
        rng = sht.range((2,j),(1024,j))
        rng.api.Validation.Delete()
        fml = "="+r
        rng.api.Validation.Add(Type=xlValidateList, Formula1=fml)
        j += 1



    #%% Do the lookup  
    vals = sht.range("A1").expand("down").value
    [print(v) for v in vals]
    n_rows = vals.index("--END--")+1
    print(n_rows)

    for i in range(1, n_rows):
        queryType = sht.range((i,3)).value

        if queryType == 'xpath':
            xpath_str = sht.range((i,4)).value
            val = dcchf.xpath_query(root, xpath_str)
            print(vals[i-1], queryType, val)
            if len(val)>0:
                sht[f"M{i}"].value = val[0].text
            else:
                sht[f"M{i}"].value = "ERROR not Found"
        elif queryType == 'data':
            dtbl = dict(zip(["measuringSystemRef", "tableId"], sht.range((i,5),(i,6)).value))
            dcol = dict(zip(["metaDataCategory", "scope", "dataCategory","measurand"], sht.range((i,7), (i,10)).value))
            unit = sht.range((i,11)).value
            customerTag = sht.range((i,12)).value
            if customerTag is None:
                data = search(root, dtbl, dcol, unit)
            else: 
                data = search(root, dtbl, dcol, unit, rowTags=[customerTag])
            print(dtbl, dcol, unit, customerTag, ":", data)
            if len(data) == 0 : 
                data_val = "ERROR not Found"
            else:
                data_val = list(data.values())
            # elif len(data) > 1: 
                # data_val = "ERROR too many values"
            sht[f"M{i}"].value = data_val
        #     cell = sht.range((i, 14, value=data)
        # else: 
        #     cell = sht.range((i, 14, value="FAILED")



class MainApp(tk.Tk):
    def __init__(self, queryTool: DccQuerryTool):
        super().__init__()
        app = self
        self.queryTool = queryTool
        self.setup_gui(app)

        # self.configure(background='white')
        # self.bind()
        # self.bindVirtualEvents()
        
        
        
    def setup_gui(self,app):
        self.wm_title("DCC EXCEL Mapping Tool")
        # self.canvas = tk.Canvas(self, width = 1851, height = 1041)
        self.geometry('500x200')
        self.attributes('-topmost', True)
        # img = tk.PhotoImage(file = 'BG_design.png')
        # self.background_image = img
        # self.background_label = tk.Label(app, image=img)
        # self.background_label.place(x=0, y=0, relwidth=1, relheight=1)
        
        # Create three buttons
        button1 = tk.Button(app, text='load query excel-file', command=self.loadExcelBook)
        button1.pack(pady=10)
        self.label1 = tk.Label(app, text = "Select a schema file")
        self.label1.pack()

        button2 = tk.Button(app, text='load DCC.xml', command=self.loadDCC)
        button2.pack(pady=10)
        self.label2 = tk.Label(app, text = "Select a DCC file")
        self.label2.pack()

        button3 = tk.Button(app, text='run query', command=self.queryTool.runDccQuery)
        button3.pack(pady=10)

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
        self.queryTool.loadDccAttributes()
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

         
