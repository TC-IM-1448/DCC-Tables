import win32com.client
import lookupFunktions as lookup

"""Open an active excel sheet"""
ExApp=win32com.client.GetActiveObject("Excel.Application")
ExApp.Visible="True"
wb=ExApp.Workbooks.Open("DCC-lookup4.xlsm")

"""Get serch criteria from sheet"""
xmlfile=ExApp.Range("B1").Value
tableId=ExApp.Range("B2").Value
attribname1=ExApp.Range("A9").Value
attribname2=ExApp.Range("A10").Value
attribname3=ExApp.Range("A11").Value

attribval1=ExApp.Range("B9").Value
attribval2=ExApp.Range("B10").Value
attribval3=ExApp.Range("B11").Value
unit=ExApp.Range("B12").Value
customerTag=ExApp.Range("B14").Value

"""Find the right table using tableId"""
tab=lookup.getTableFromXML(xmlfile,tableId)
"""Find the rigt column using attributes and unit"""
attrib={attribname1:attribval1,attribname2:attribval2,attribname3:attribval3}
col=lookup.getColumnFromTable(tab,attrib,unit)

"""Find the column containing the user Tags"""
tagcol=lookup.getColumnFromTable(tab,{'scope':'dataInfo','dataCategory':'customerTag','measurand':'metaData'},'nan')

"""Iterate through the tags to find the row number of the specified tag"""
tags=tagcol[2].text.split()
for i, tag in enumerate(tags):
    if tag==customerTag:
        found=True
        break

TargetCell="D1"
if found and type(col)!=type(None):
    searchValue=col[2].text.split()[i]
    ExApp.Range(TargetCell).Value=searchValue
    print('1')
else:
    ExApp.Range(TargetCell).Value="No Data Found"
    print('0')

