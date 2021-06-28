# Microsoft Office Specialist: Excel Associate (Excel and Excel 2019) - Cheatsheet
Passing grade: 700

The plan for this repository is to create a comprehensive-ish cheatcheet for the MO-200 / Excel 2019 Associate Exam.
I forgot to add sample questions, but if you want then do help out with some of them.

The automatic program checking what you've done is lacking, to say the least. Though there's multiple ways of carrying out the tasks, it seems to only accept your work if you suffer through the ribbon menues (who thought of this a brilliant idea? 🤦). 

Hence why I make a start on this list. Do help out if there's some I missed. 👍
I numbered the list for easier discovery 😄

v0.9.0 update: With this list, I got 978/1000 in the Excel 2019 Exam. :)

v0.9.1 update: I can't figure out how to change the license. Anyway, do as you will provided you give some attribution in your footnotes: https://creativecommons.org/licenses/by/4.0/

Syllabus list from https://query.prod.cms.rt.microsoft.com/cms/api/am/binary/RE3WKkC.
Mike's Office videos are also pretty good: https://www.youtube.com/channel/UC92eTyKmq0zOIjIK1ShycKw

## 0.0 These are some questions I have seen and have zero idea about
Configure print settings -> "sett to always print cells A1:h17" "print out header in all sheets".

## 0.1 This is for the ones which I have no idea where to locate under yet. Maybe they were taken out in the transition from Excel 2016 to 2019?

* Subtotal feature: 
"calculate the number of students by colour. insert a page break after each change, display a grand total in g39".
highlight range (not table). Data -> Outline -> Subtotal. Select according details they ask you for e.g. add each to city. show in amount. 

* Image manipulation.
Click on image -> Picture format -> you've all kinds of picture effects.
there's also rotate, size, etc all on there.

  * Hide: Right click hide on a worksheet.
  * Unhide: Home -> Cells -> Format -> Hide & Unhide -> Unhide Sheet. Select that sheet.

* Switch axis of a chart:
Click on chart -> Chart Design -> Data -> Switch Row/Column

## 1. Manage worksheets and workbooks (10-15%)

### 1.1 Import data into workbooks

#### 1.1.1 Import data from .txt files
Data -> From Text -> Import Text. Pick Delimiter (usually comma or tab). Takr care on where to load.

#### 1.1.2 Import data from .csv files
Too similar to .txt files above.

### 1.2 Navigate within workbooks

#### 1.2.1 Search for data within a workbook
Find: 
Home -> Editing -> Find and Select -> Find
Replace: 
Home -> Editing -> Find and Select -> Replace "stringA" with "stringB" -> Replace all

#### 1.2.2 Navigate to named cells, ranges, or workbook elements
Location table is top left of A1 cell. 

#### 1.2.3 Insert and remove hyperlinks
Hyperlinks: 
right click on cell -> insert/remove hyperlink. Sometimes you have to link cells to one another, other times it is websites.

### 1.3 Format worksheets and workbooks

#### 1.3.1 Modify page setup
Tab manipulation: 
This is all in the lowermost ribbon. Add, delete are pretty simple; change colour and other more niche features are a right click away.

#### 1.3.2 Adjust row height and column width
Just click a column and drag it up/dowm, or right click and gfo through the more comprehensive settings.

#### 1.3.3 Customize headers and footers
  
### 1.4 Customize options and views

#### 1.4.1 Customize the Quick Access toolbar

#### 1.4.2 Display and modify workbook content in different views
Lower Right corner. has three buttons for different views. not a very common question.

#### 1.4.3 Freeze worksheet rows and columns

#### 1.4.4 Change window views

#### 1.4.5 Modify basic workbook properties
Change/Add/Delete property "X": 
File -> Info -> Properties. You may have to click the button to "show all properties" sometimes. 

#### 1.4.6 Display formulas
 Formulas -> Show Formulas.
  
### 1.5 Configure content for collaboration

  #### 1.5.1 Set a print area
  Highlight the table -> File -> Print -> Selection
  """??Page Layout -> can set page breaks there?"""
  
  #### 1.5.2 Save workbooks in alternative file formats
  File -> Save As -> whatever other format. or ctrl+shift+v... 
  
  #### 1.5.3 Configure print settings
  
  #### 1.5.4 Inspect workbooks for issues
  "Hide personal information": 
File -> Info -> Inspect Workbook -> Check for issues.
  
## 2. Manage data cells and ranges (20-25%)

### 2.1 Manipulate data in worksheets

  #### 2.1.1	Paste data by using special paste options
  usual left click -> paste or ctrl+v, but then hit the drop-down arrow in the lower right corner of what you just pasted.
  
  #### 2.1.2	Fill cells by using Auto Fill
  Grab the cell's lower-right corner, and drag down/left/etc to your heart's desire.
  Auto Fill Options -> without formatting.
  
  #### 2.1.3	Insert and delete multiple columns or rows
  
  #### 2.1.4	Insert and delete cells
  Right click -> delete. or insert. move up/left/down/right dep. on what the question asks you.
  
  Else, if it's delete contents just highlight and delete with the "Del" key.
   
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

### 2.3 Define and reference named ranges
  #### 2.3.1	Define a named range
  #### 2.3.2	Name a table
  Naming: Select cell/range -> top left of A1 -> whrite the name there
  
### 2.4 Summarize data visually
  ### 2.4.1	Insert Sparklines
  Sparklines are little diagrams or visual graphs contained in one cell. 
  insert -> sparklines -> line (or others). select its data range.
  
  ### 2.4.2	Apply built-in conditional formatting
  Home -> Format -> Conditional Formatting. 
  
  ### 2.4.3	Remove conditional formatting
  Home -> Format -> Conditional Formatting. 

## 3. Manage tables and table data (15-20%)

### 3.1 Create and format tables

#### 3.1.1 Create Excel tables from cell ranges
Structured range -> Table:
select cells. insert -> table. select headers.

#### 3.1.2 Apply table styles
Click table -> Table design -> Table Styles -> More. Select correct style.

Note: Every other row is "banded rows". yw!

#### 3.1.3 Convert tables to cell ranges
Select the table. Table -> Tools -> convert to range -> accept. 

### 3.2 Modify tables
 
#### 3.2.1 Add or remove table rows and columns
#### 3.2.2 Configure table style options
#### 3.2.3 Insert and configure total rows

### 3.3 Filter and sort table data

#### 3.3.1 Filter records
Click table -> Table tools -> data -> remove duplicates

#### 3.3.2 Sort data by multiple columns
I believe it is Table tools -> Data -> Sort. It allows you to set multiple hierarchies of sorting. I might be mistaken.

## 4. Perform operations by using formulas and functions (20-25%)
quick tip: if you start righting a formula, below the top ribbon menues there lies an fx button. That'll give you a popup if you're uncomfortable with just writing it out like the pros, xd.

### 4.1 Insert references

#### 4.1.1	Insert relative, absolute, and mixed references
#### 4.1.2	Reference named ranges and named tables in formulas

### 4.2 Calculate and transform datas

#### 4.2.1	Perform calculations by using the AVERAGE(), MAX(), MIN(), and SUM() functions
AVERAGE(range) = mean of the values inside. e.g. AVERAGE(A1:A7)

MAX(range) = highest value in its range.  e.g. MAX(B2:B17)

MIN(range) = highest value in its range.  e.g. MIN(B2:B17)

SUM(range) = adds all the values in its range.  e.g. SUM(B2:B17)

#### 4.2.2	Count cells by using the COUNT(), COUNTA(), and COUNTBLANK() functions
" enter a function which will count the total number of cells with 'x'":

COUNT(range) = counts the amount of cells with a number inside
COUNTA(range) = counts the amount of cells with anything inside. used for strings
COUNTBLANK(range) = counts the amount of cells with nothing inside

#### 4.2.3	Perform conditional operations by using the IF() function

IF(logicaltest, true, false) : will do T or F on the basis of the logical test.

For example, IF(A7>16, "Scooter", "Bicycle")

Other conditionals:
* SUMIF(range for criteria, criteria, sumrange) adds all cells if the condition is true, per cell basis e.g. SUMIF(B44:B77, "Food", C:44:C77) 
* AVERAGEIF(range for criteria, criteria, avgrange)  averages all cells for which the condition holds true
* COUNTIF(range for criteria, criteria). conditional count. (range, criteria) e.g. countif(a1:a7, "green")

### 4.3 Format and modify text

#### 4.3.1	Format text by using RIGHT(), LEFT(), and MID() functions
* RIGHT(cell, n) gives the rightmost n characters of a cell
* LEFT(cell, n) gives the leftmost n characters of a cell
* MID(cell, n) gives the middle n characters of a cell

#### 4.3.2	Format text by using UPPER(), LOWER(), and LEN() functions
* UPPER(cell) gives you the cell but all in uppercase 
* LOWER(cell) gives you the cell but all in lowercase 
* LEN(cell) returns the length of a string
* PROPER(cell) gives you the cell but all in proper English 

#### 4.3.3	Format text by using the CONCAT() and TEXTJOIN() functions
* CONCAT() adds two strings together.
* TEXTJOIN() is a new 2019 Excel formula which has a lot more fuunctionality (https://exceljet.net/excel-functions/excel-textjoin-function)
* TRIM(cell) removes odd spacings.

## 5. Manage charts (20-25%)
### 5.1 Create charts

#### 5.1.1	Create charts
Pretty basic stuff. I believe it is select data -> insert -> chart (by type)

#### 5.1.2	Create chart sheets
Click on a chart -> Chart Design -> Location -> Move Chart -> to a new sheet. Name it accordingly.

### 5.2 Modify charts

#### 5.2.1	Add data series to charts
chart tools/select data/can change the legends and other bits.
Click on chart -> extend the coloured box that appears around the used (highlighted) cells.

#### 5.2.2	Switch between rows and columns in source data
Chart Tools -> Design switch row/column

#### 5.2.3	Add and modify chart elements
Click the chart -> (top-right of the chart) click the "+" button -> Select chart element

### 5.3 Format charts

#### 5.3.1	Apply chart layouts
Click on chart -> Chart Design -> Chart Layouts -> Quick Layouts.
 
#### 5.3.2	Apply chart styles
Click on chart -> Chart Design -> Chart Styles -> Select whichever from the dropdown menu.

#### 5.3.3	Add alternative text to charts for accessibility
Click object -> Right click -> Edit alt text.

