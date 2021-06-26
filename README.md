# Microsoft Office Specialist: Excel Associate (Excel and Excel 2019) - Cheatsheet
The plan for this repository is to create a comprehensive-ish cheatcheet for the MO-200 / Excel 2019 Associate Exam.

The automatic program checking what you've done is lacking, to say the least. Though there's multiple ways of carrying out the tasks, it seems to only accept your work if you suffer through the ribbon menues (who thought of this a brilliant idea? 🤦). 

Hence why I make a start on this list. Do help out if there's some I missed. 👍
I numbered the list for easier discovery 😄

Syllabus list from https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE3WKkC.
Mike's Office videos are also pretty good: https://www.youtube.com/channel/UC92eTyKmq0zOIjIK1ShycKw

## 0. This is for the ones which I have no idea where to locate under yet.
Change a string: Home -> Editing -> Find and Select -> Replace "stringA" with "stringB" -> Replace all

subtital feature: calculate the number of students by clolour.
insert a page break after each change, display a grand total in g39


## 1. Manage worksheets and workbooks (10-15%)


### 1.1 Import data into workbooks

#### 1.1.1 Import data from .txt files
Data -> From Text -> Import Text. Pick Delimiter (usually comma or tab). Takr care on where to load.

#### 1.1.2 Import data from .csv files
Too similar to .txt files above.


### 1.2 Navigate within workbooks

#### 1.2.1 Search for data within a workbook
I have no idea

#### 1.2.2 Navigate to named cells, ranges, or workbook elements
location table is top left of A1 cell. 

#### 1.2.3 Insert and remove hyperlinks
right click -> insert/remove hyperlink. Sometimes you have to link cells to one another, other times it is websites.


### 1.3 Format worksheets and workbooks

#### 1.3.1 Modify page setup
#### 1.3.2 Adjust row height and column width
#### 1.3.3 Customize headers and footers
  
  
### 1.4 Customize options and views

#### 1.4.1 Customize the Quick Access toolbar
#### 1.4.2 Display and modify workbook content in different views
#### 1.4.3 Freeze worksheet rows and columns
#### 1.4.4 Change window views
#### 1.4.5 Modify basic workbook properties
#### 1.4.6 Display formulas
  
  
### 1.5 Configure content for collaboration
  #### 1.5.1 Set a print area
  #### 1.5.2 Save workbooks in alternative file formats
  #### 1.5.3 Configure print settings
  #### 1.5.4 Inspect workbooks for issues



## 2. Manage data cells and ranges


### 2.1 Manipulate data in worksheets

  #### 2.1.1	Paste data by using special paste options
  usual left click -> paste or ctrl+v, but then hit the drop-down arrow in the lower right corner of what you just pasted.
  
  #### 2.1.2	Fill cells by using Auto Fill
  Grab the cell's lower-right corner, and drag down/left/etc to your heart's desire.
  
  #### 2.1.3	Insert and delete multiple columns or rows
  #### 2.1.4	Insert and delete cells
  
  
### 2.2 Format cells and ranges
  #### 2.2.1	Merge and unmerge cells
  Home -> Alignment -> Merge And Centre 
  Note: Merge Across hidden in the Merge and Centre drop arrow - if "without changing the alignment of the text"
  
  #### 2.2.2	Modify cell alignment, orientation, and indentation
  Home -> Alignment. you have left, right, indentation, orientation and wrap under there. make yourself aware of all the WYSIWYG buttons. 
  
  #### 2.2.3	Format cells by using Format Painter
  #### 2.2.4	Wrap text within cells
  #### 2.2.5	Apply number formats
  #### 2.2.6	Apply cell formats from the Format Cells dialog box
  #### 2.2.7	Apply cell styles
  #### 2.2.8 Clear cell formatting


> Define and reference named ranges
  •	Define a named range
  •	Name a table
  
  
> Summarize data visually
  •	Insert Sparklines
  •	Apply built-in conditional formatting
  •	Remove conditional formatting



Manage tables and table data


Create and format tables

•	Create Excel tables from cell ranges
•	Apply table styles
•	Convert tables to cell ranges

Modify tables
 
•	Add or remove table rows and columns
•	Configure table style options
•	Insert and configure total rows

Filter and sort table data

•	Filter records
•	Sort data by multiple columns

Perform operations by using formulas and functions
Insert references

•	Insert relative, absolute, and mixed references
•	Reference named ranges and named tables in formulas

Calculate and transform datas

•	Perform calculations by using the AVERAGE(), MAX(), MIN(), and SUM() functions
SUM ()

•	Count cells by using the COUNT(), COUNTA(), and COUNTBLANK() functions
" enter a function which will count the total number of cells with 'x'":
number of cells -> COUNT () - int; or COUNTA () - str



•	Perform conditional operations by using the IF() function
if

conditional count
countif(range, criteria) e.g. countif(a1:a7, "green")

conditional average XXXX


Format and modify text

•	Format text by using RIGHT(), LEFT(), and MID() functions
•	Format text by using UPPER(), LOWER(), and LEN() functions
•	Format text by using the CONCAT() and TEXTJOIN() functions

Manage charts
Create charts

•	Create charts
•	Create chart sheets

Modify charts

•	Add data series to charts
•	Switch between rows and columns in source data
•	Add and modify chart elements

Format charts

•	Apply chart layouts
 
•	Apply chart styles
•	Add alternative text to charts for accessibility

