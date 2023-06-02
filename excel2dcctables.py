from pydcc_tables import DccTabel, DccTableColumn
import openpyxl as pyxl 
import pandas as pd

def transpose_2d_list(matrix):
    return [list(row) for row in zip(*matrix)]

def _test_get_tables_from_sheet():
    """ Function that finds all the tables in a given sheet """
    
    wb = pyxl.load_workbook("DCC-Table_example3.xlsx", data_only=True)

    #access specific sheet
    ws = wb["Table2"]

    rngs = {key: value for key, value in ws.tables.items()}

    # mapping = {}
    # for entry, data_boundary in ws.tables.items():
    mapping = {}
    columns = []

    for entry, data_boundary in ws.tables.items():
        d = {}
        tableID = ws["B2"].value
        itemID = ws["B3"].value

        #parse the data within the ref boundary
        data = ws[data_boundary]
        #extract the data
        #the inner list comprehension gets the values for each cell in the table
        content = [[cell.value for cell in ent] for ent in data]
        content = transpose_2d_list(content)

        header = content[0]
        for c in content:
            col = DccTableColumn(   relationType=c[1],
                                    columnType=c[2], 
                                    measurandType=c[3], 
                                    unit=c[4], 
                                    humanHeading = c[5],
                                    columnData=c[6:])
            columns.append(col)
        
        #the contents ... excluding the header
        # rest = content[1:]

        #create dataframe with the column names
        #and pair table name with dataframe
        
        # df = pd.DataFrame(rest, columns = header)
        # mapping[entry] = df

        # d['header'] = header
        # d['data'] = data
        # d[
        tbl = DccTabel(tableID, itemID, columns)
    wb.close()
    return tbl, tableID, itemID, columns, content, mapping



if __name__ == "__main__": 
    tbl, tableID, itemID, columns, content, mapping = _test_get_tables_from_sheet()
    columns[4].print()