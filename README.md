# DCC-Tables

This repository represents a solution for industrial oriented Digital Calibration Certificates (iDCC).  It has been developped in the atempt of generating a python based middleware for generating and loading digital calibration certificates with an Excel user-interface using the python package "xlwings". The provided middleware and xml-schema has been developed in parallel, which has been crucial to choices made for the xml-schema design, to facilitate flexibility for the DCC content while keeping middleware maintenance as low as possible. Being build around an excel user interface inherently makes it necessary to adapt to datastructures to the table-formats of the Excel sheets, which conveniently makes a good match to the typical datastructure used in relational SQL databases.  

The overall data structure of iDCC-xml's generated in this framework is represented in the diagram below.
![image](https://github.com/TC-IM-1448/DCC-Tables/assets/123001590/71f3c4dd-6516-4710-9e6e-8b8e77f4d8f6)

To generate xml-dcc's that are in accordance with schema, an Excel sheet is used as interface in order to simplify the process of providing the input for the DCC. A base pipette example is provided in "SKH_10112_2.xml" and other examples are provided in the Examples folder. 

To run the gui interface run the following program: 

'> python ioDccGuiTool.py

# Primary files
* dcc.xsd : is the xml-Schema for the iDCC.  
* ioDccGuiTool.py : A demo UI tool loading, editing and generating DCC's, DCR's and templates. 
* dccQueryGui.py : A demo tool intented for clients when wanting to load specific data from receied DCC's. An base-examples is provided in the SKH_10112_2_Mapping.xlsx file. 


# Screenshots

![image](https://github.com/TC-IM-1448/DCC-Tables/assets/123001590/25861559-a820-412f-9e90-73f2b591674a)

![image](https://github.com/TC-IM-1448/DCC-Tables/assets/123001590/108218d7-7ca2-4340-9f8b-c98ed45766ef)

![image](https://github.com/TC-IM-1448/DCC-Tables/assets/123001590/8ca06471-30eb-4e59-b5bd-9d33c325e9b6)


# ToDo:
- [ ] make languages dynamic in the middleware, presently static to EN and DA.
- [ ] Provide more examples. 
- [ ] Find a solution to the 

contact info: Daid Balslev-Harder please write to (dbh @ dfm.dk) 