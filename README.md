# Microsoft Office Specialist: Excel Associate (Excel and Excel 2019) - Cheatsheet
The plan for this repository is to create a comprehensive-ish cheatcheet for the MO-200 / Excel 2019 Associate Exam.

The automatic program checking what you've done is lacking, to say the least. Though there's multiple ways of carrying out the tasks, it seems to only accept your work if you suffer through the ribbon menues (who thought of this a brilliant idea? 🤦). 

Hence why I make a start on this list. Do help out if there's some I missed. 👍
I numbered the list for easier discovery 😄

Syllabus list from https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE3WKkC.
Mike's Office videos are also pretty good: https://www.youtube.com/channel/UC92eTyKmq0zOIjIK1ShycKw

## 0. This is for the ones which I have no idea where to locate under yet. Maybe they were taken out in the transition from Excel 2016 to 2019?

* Subtotal feature: 
"calculate the number of students by colour. insert a page break after each change, display a grand total in g39".
highlight range (not table). Data -> Outline -> Subtotal. Select according details they ask you for.

* Image manipulation.
Click on image -> Picture format -> you've all kinds of picture effects.

* Tab manipulation.
This is all in the lowermost ribbon. Add, delete are pretty simple; change colour and other more niche features are a right click away.
  * Hide: Right click hide on a worksheet.
  * Unhide: Home -> Cells -> Format -> Hide & Unhide -> Unhide Sheet. Select that sheet.

## 1. Manage worksheets and workbooks (10-15%)


### 1.1 Import data into workbooks

#### 1.1.1 Import data from .txt files
Data -> From Text -> Import Text. Pick Delimiter (usually comma or tab). Takr care on where to load.

#### 1.1.2 Import data from .csv files
Too similar to .txt files above.


### 1.2 Navigate within workbooks

#### 1.2.1 Search for data within a workbook
Find: Home -> Editing -> Find and Select -> Find
Replace: Home -> Editing -> Find and Select -> Replace "stringA" with "stringB" -> Replace all

#### 1.2.2 Navigate to named cells, ranges, or workbook elements
Location table is top left of A1 cell. 

#### 1.2.3 Insert and remove hyperlinks
Hyperlinks: right click on cell -> insert/remove hyperlink. Sometimes you have to link cells to one another, other times it is websites.


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
Change/Add/Delete property "X"
File -> Info -> Properties. You may have to click the button to change properties.

"Hide personal information"
File -> Info -> Inspect Workbook -> Check for issues.

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
  Naming: Select cell/range -> top left of A1 -> whrite the name there
  
> Summarize data visually
  •	Insert Sparklines
  •	Apply built-in conditional formatting
  •	Remove conditional formatting



Manage tables and table data


Create and format tables

•	Create Excel tables from cell ranges

•	Apply table styles
Click table -> Table design -> Table Styles -> More. Select correct style

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
Click on a chart -> Chart Design -> Location -> Move Chart -> to a new sheet. Name it accordingly.

Modify charts

•	Add data series to charts


•	Switch between rows and columns in source data

•	Add and modify chart elements
Click the chart -> (top-right of the chart) click the "+" button -> Select chart element

Format charts

•	Apply chart layouts
Click on chart -> Chart Design -> Chart Layouts -> Quick Layouts.
 
•	Apply chart styles
Click on chart -> Chart Design -> Chart Styles -> Select whichever from the dropdown menu.

•	Add alternative text to charts for accessibility

