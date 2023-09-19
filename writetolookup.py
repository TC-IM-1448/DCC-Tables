import win32com.client
import lookupFunktions as lookup

def search(root, tableAttrib, colAttrib, unit, customerTag=None):
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

   searchValue="-"
   warning="-"
   usertagwarning="-"
   colwarning="-"

       try:
           """Find the right result using resId"""
           res=lookup.getResultFromRoot(root, resId="")
           try:
               """Find the right table using itemRef and settingRef"""
               tab=lookup.getTableFromResult(res, tableAttrib)
               try:
                   """Find the rigt column using attributes and unit"""
                   col=lookup.getColumnFromTable(tab,colAttrib,unit)
                   try:
                       if type(customerTag)!=type(None):
                          searchValue=lookup.getRowFromColumn(col,tab,customerTag)
                       else:
                           searchValue=col[2].text.split()
                   except Exception as e:
                       usertagwarning=e.args[0]
               except Exception as e:
                   colwarning=e.args[0]
           except Exception as e:
               warning=e.args[0]
       except Exception as e:
           warning=e.args[0]

   return [searchValue, usertagwarning, colwarning, warning]


#Import search values from sheet

"""Open an active excel sheet"""
ExApp=win32com.client.GetActiveObject("Excel.Application")
ExApp.Visible="True"
wb=ExApp.Workbooks.Open("DCC-lookup5e.xlsm")

"""Get serch criteria from sheet"""
xmlfile=ExApp.Range("B1").Value
tableId=ExApp.Range("B2").Value
itemRef=ExApp.Range("B3").Value
settingRef=ExApp.Range("B4").Value
tableAttrib={'tableId':tableId, 'itemRef':itemRef, 'settingRef':settingRef}

attribname1=ExApp.Range("A9").Value
attribname2=ExApp.Range("A10").Value
attribname3=ExApp.Range("A11").Value
attribval1=ExApp.Range("B9").Value
attribval2=ExApp.Range("B10").Value
attribval3=ExApp.Range("B11").Value
colAttrib={attribname1:attribval1,attribname2:attribval2,attribname3:attribval3}
unit=ExApp.Range("B12").Value
customerTag=ExApp.Range("B14").Value
root=lookup.getRoot(xmlfile)

#Perform search
[searchValue, usertagwarning, colwarning, warning] = search(root, tableAttrib, colAttrib, unit, customerTag)

#Export result to sheet
print(usertagwarning, colwarning,warning)
print(colwarning)
print(warning)
print("searchValue")
print(searchValue)
ExApp.Range("E4").Value=searchValue
ExApp.Range("E5").Value=warning
ExApp.Range("E6").Value=colwarning
ExApp.Range("E7").Value=usertagwarning

