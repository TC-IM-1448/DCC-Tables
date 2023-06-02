

class DccTableColumn():
    """ """ 
    columnType = ""
    relationType = ""
    measurandType = ""
    unit = ""
    columnData = []

def __init__(self,  
             columnType="", 
             relationType="",
             measurandType="", 
             unit="", 
             columnData=[]):
    """ """
    self.columnType = columnType
    self.relationType = relationType
    self.measurandType = measurandType
    self.unit = unit
    self.columnData = columntData

class DccTabel():
    """ """
    tableID = ""
    itemID = ""
    columns = []

    def __init__(self, tableID="", itemID="", columns=""): 
        self.tableID = tableID
        self.itemID = itemID
        self.columns = columns