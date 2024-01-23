#%%import os
import openpyxl as pyxl
import xlwings as xw
from  lxml import etree as et
import DCChelpfunctions as dcchf
from DCChelpfunctions import search

#%%
LANG='da'
DCC='{https://dfm.dk}'

#%%
def lookupFromMappingFile(mapFileName:str, dccFileName:str):
    """LookupFromMappingFile """
    pass

#%%
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

tableIds = [elm.attrib['tableId'] for elm in dcchf.get_tables(root)]
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



#%%
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
        data = search(root, dtbl, dcol, unit, customerTag=customerTag )
        print(dtbl, dcol, unit, customerTag, ":", data)
        sht[f"M{i}"].value = data
    #     cell = sht.range((i, 14, value=data)
    # else: 
    #     cell = sht.range((i, 14, value="FAILED")



#%% Make a test of the code
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


if __name__=="__main__":
    #################
    #first argument is dcc xml file
    #second argument is excel template to use

    import sys
    args=sys.argv[1:]
    print(len(args))
    if len(args)==0:
        mapFileName ='Examples'+os.sep+'Mapping_Novo_temperatur_Certifikat.xlsx'
        dccFileName = 'Examples'+os.sep+'Stip-230063-V1.xml'
        lookupFromMappingFile(mapFileName, dccFileName)
    elif len(args)==2:
        mapFileName = args[0]
        dccFileName = args[1]
        lookupFromMappingFile(mapFileName, dccFileName)
    else: 
        helpstatement = """call dccquery.py using the following arguments: \n 
        >> python dccquery.py [mapping file e.g. mapping.xlsx] [DCC file e.g. dcc.xml] """
        print(helpstatement)

         
