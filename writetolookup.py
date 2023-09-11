import win32com.client
import lookupFunktions as lookup

"""Open an active excel sheet"""
ExApp=win32com.client.GetActiveObject("Excel.Application")
ExApp.Visible="True"
wb=ExApp.Workbooks.Open("DCC-lookup5d.xlsm")

"""Get serch criteria from sheet"""
xmlfile=ExApp.Range("B1").Value
resultId=ExApp.Range("B2").Value
itemRef=ExApp.Range("B3").Value
settingRef=ExApp.Range("B4").Value
attribname1=ExApp.Range("A9").Value
attribname2=ExApp.Range("A10").Value
attribname3=ExApp.Range("A11").Value

attribval1=ExApp.Range("B9").Value
attribval2=ExApp.Range("B10").Value
attribval3=ExApp.Range("B11").Value
unit=ExApp.Range("B12").Value
customerTag=ExApp.Range("B14").Value
searchValue="-"
warning="-"
usertagwarning="-"
colwarning="-"

try:
    """get the root element of the DCC"""
    root=lookup.getRoot(xmlfile)
    try:
        """Find the right result using resId"""
        res=lookup.getResultFromRoot(root, resId=resultId)
        try:
            """Find the right table using itemRef and settingRef"""
            tab=lookup.getTableFromResult(res,itemRef,settingRef )
            try:
                """Find the rigt column using attributes and unit"""
                attrib={attribname1:attribval1,attribname2:attribval2,attribname3:attribval3}
                col=lookup.getColumnFromTable(tab,attrib,unit)
                try:
                    searchValue=lookup.getRowFromColumn(col,tab,customerTag)
                except Exception as e:
                    usertagwarning=e.args[0]
            except Exception as e:
                colwarning=e.args[0]
        except Exception as e:
            warning=e.args[0]
    except Exception as e:
        warning=e.args[0]
except:
    warning="Could not open the file"

print(searchValue)
ExApp.Range("E4").Value=searchValue
ExApp.Range("E5").Value=warning
ExApp.Range("E6").Value=colwarning
ExApp.Range("E7").Value=usertagwarning

