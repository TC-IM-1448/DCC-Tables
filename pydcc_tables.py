

class DccTableColumn():
    """ """ 
    columnType = ""
    relationType = ""
    measurandType = ""
    unit = ""
    humanHeading = ""
    columnData = []

    def __init__(self,  
                columnType="", 
                relationType="",
                measurandType="", 
                unit="", 
                humanHeading = "",
                columnData=[]):
        """ """
        self.columnType = columnType
        self.relationType = relationType
        self.measurandType = measurandType
        self.unit = unit
        self.humanHeading = humanHeading
        self.columnData = columnData

    def get_attributes(self):
        attr = [attr for attr in dir(self) if not callable(getattr(self, attr)) and not attr.startswith("__")]
        # local_attributes = {k: v for k, v in vars(self).items() if k in locals()}
        return attr

    def print(self):
        attr = self.get_attributes()
        str = ""
        for a in attr: 
            print(a, ": \t", getattr(self, a))


class DccTabel():
    """ """
    tableID = ""
    itemID = ""
    columns = []

    def __init__(self, tableID="", itemID="", columns=""): 
        self.tableID = tableID
        self.itemID = itemID
        self.columns = columns