## ProExcel

## Objective:
The ProExcel module empowers developers to effortlessly interact with Excel files by providing a seamless way to read, write, and edit existing Excel documents. Whether you're working with data, lists, or tables, the module allows you to access and manipulate Excel content with minimal effort. By simply specifying key parameters such as column, row, and sheet names, developers can quickly perform complex operations on Excel files, significantly streamlining workflows and reducing development time. With ProExcel, handling intricate Excel-related tasks becomes easier, making it an invaluable tool for any developer working with Excel in Mendix.

## Dependencies
•	Studio pro version 9.24.32\
•	Excel Importer\
•	Community Commons\
•	Mx-Model Reflection


## Note : Excel importer , community commons and Mx-Model reflection are mandatory to import before importing the module

## Configuration
•	Import the module Excel Importer, Community Commons, Mx-Model Reflection Module from the Mendix marketplace\
•	Import the module ProExcel Module from the Mendix marketplace.

## Requirement 1: If we want to read the data from specific cells.

<img width="872" alt="image" src="https://github.com/user-attachments/assets/45564fb3-e1c2-4137-87ee-cd3bc000ebb4" />

<img width="722" alt="image" src="https://github.com/user-attachments/assets/5f2e6983-f24f-444a-a909-7fb51c4fbb10" />

## Note : You can pass your own excel file too( Refer first 3 activities in microflow )
## Steps to pass your own file
1.First get the file by any data source.\
2. After that use java action named as DuplicateFileDocument to copy your file content to ExcelFile entity object.

## Steps to read the data
1.Retrieve the ExcelRowCellRead list of objects which will contain the data for in which row, column and sheet we want to read the data.\
( You can make this as a master data by adding the page in navigation or you can create it in runtime depending on the requirement ).\
•	Here sheet name is the name of the sheet from which you want to read the data.\
•	Row number is same as the numbers default there in excel sheet.\
•	Column number will represent as A = 1, B = 2 , C = 3 ….\
•	Click on the Read data by rows and columns button\
2.Pass the ExcelRowCellRead list of objects and Excel file in which we want to read the data to the java action name as Java_action_ExcelReadByRowColumn.\
3.After triggering the microflow , you will get the data in the format of string . You can refer the image below for better clarification.

## Requirement 2: If we want to read the list of data horizontally

## Example
<img width="884" alt="image" src="https://github.com/user-attachments/assets/2155287e-c100-4979-9533-25da2abcbe5f" />

<img width="869" alt="image" src="https://github.com/user-attachments/assets/86e23879-0dc9-4fcf-aefd-2d845ae9143c" />

## Steps to read the table

<img width="219" alt="image" src="https://github.com/user-attachments/assets/ea585885-15b1-41cb-8fd4-c4af5c7ac597" />


1.First 3 activities are the same as mentioned in Note : You can pass your own excel file too( Refer first 3 activities in microflow ).\
2.Activity 4  is Number of rows you want to read from excel, by taking the example.. here I want to read 5 rows\ 
3.Activity 5  is Empty object list, here we are passing empty object list of entity name DynamicEntity to store the data in it. If we want to read the 5 line of data then in that case we need to create 5 number of objects.\
4.Activity 6 states that how many colums are there in the table . In example case there are 4 columns . Name, Roll no , ID , Location\
5.Activity 7 states from which row number we want to read the data . Taking example case , the table is starting from row number 8.\
6. Activity 8 states that from which column number we want to read the data . Taking example case , the table is starting from C that means 3.\
7. Activity S9 states that from which sheet you want to read the data.\
8.Now pass these all data to the java action Java_action_ExcelRead_Horizontally.\
After triggering the microflow , it will give the data in first four attributes in DynamicEntity entity.

Data will be\
List of DynamicEntity we get after java action 

<img width="869" alt="image" src="https://github.com/user-attachments/assets/ca9a7d69-bc77-47ed-858b-477dc0581658" />\
This list of data we can manipulate in our application .

## Requirement 3: If we want to read the list of data Vertically
## Example

![image](https://github.com/user-attachments/assets/7ddd31ea-93c0-4aa6-ba27-30ad08bfe8ad)

<img width="867" alt="image" src="https://github.com/user-attachments/assets/5725e87a-6e90-49b4-a65a-0d4d73ffeb80" />

After triggering the microflow , it will give the data in first four attributes in DynamicEntity entity.
Data will be
List of DynamicEntity we get after java action 

<img width="866" alt="image" src="https://github.com/user-attachments/assets/7962a3f4-b2a8-41c8-8216-7093929319c0" />

## Requirement 4: If we want to write data in cells
## Example

<img width="545" alt="image" src="https://github.com/user-attachments/assets/dba73680-7f63-42e0-8003-a89c55563694" />

<img width="871" alt="image" src="https://github.com/user-attachments/assets/78da3ec3-239d-4c0a-a062-89c013e82c23" />

1.Retrieve the ExcelRowCellWrite list of objects which will contain the data for in which row, column and sheet name and value we want to write in  the Excel.\
( You can make this as a master data by adding the page in navigation or you can create it in runtime depending on the requirement )\
•	Here sheet name is the name of the sheet from which you want to read the data.\
•	Row number is same as the numbers default there in excel sheet.\
•	Column number will represent as A = 1, B = 2 , C = 3 ….\
•	Value will be the data you want to write in excel.\
2.Pass the ExcelRowCellWrite list of objects and Excel file in which we want to write the data to the java action name as Java_action_WriteByRowAndColumn.\
3.After triggering the microflow, the Excel file will get downloaded with filled data.


## Requirement 5: If we want to read the list of data horizontally from specific cell

Provided Excel in which in blue portion I want to write the list.					
Downloaded Excel after triggering microflow.

<img width="812" alt="image" src="https://github.com/user-attachments/assets/e23c7f8a-83c8-4b75-a0cd-9e3ffe70b996" />

<img width="868" alt="image" src="https://github.com/user-attachments/assets/d4db6282-a72a-493a-9bae-a38cb3b3feca" />

## Requirement 6: If we want to read the list of data vertically from specific cell

<img width="918" alt="image" src="https://github.com/user-attachments/assets/34ebe1c9-2834-494d-8946-a44ce080bca9" />

<img width="871" alt="image" src="https://github.com/user-attachments/assets/dbce94e6-01dc-475d-be3a-ea70025143b7" />











