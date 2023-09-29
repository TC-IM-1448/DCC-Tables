# DCC-Tables

This repository represents a solution for Digital Calibration Certificates (DCC) based on matching the dataformats to that of relational datases such as used with SQL.  

The overall architechture of DCC' xml's generated in this framework is represented in the diagram below.
![image](https://github.com/TC-IM-1448/DCC-Tables/assets/123001590/71f3c4dd-6516-4710-9e6e-8b8e77f4d8f6)


dcc.xsd : is the xml-Schema for the DCC.  

To generate xml-dcc's that are in accordance with schema, an Excel sheet is used as interface in order to simplify the process of providing the input for the DCC. Excel sheets for this is provided in the Examples folder. 

To generate the xml-based DCC the python file excel2dcc.py can used using the following commandline from the main folder:

'> python excel2dcc.py Examples/DCC_GUI.xlsx


To view the content of a xml-based DCC the python file excel2dcc.py using the commandline below. This will generate the file view_content.xlsx. 

'> python dcc2excel.py Examples/DFM-T220000.xml DCC_template.xlsx 


ToDo:
- [ ] Make environment.yml for the python code
- [ ] Comple python code. 
- [ ] In dcc2excel.py: Add items and settings import functions.
