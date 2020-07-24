# EXCEL_Coursera
  To log the study of the excel courses on Coursera 

## [Excel Skills for Business: intermediate i](https://www.coursera.org/learn/excel-intermediate-1)
#### Week 1
Working with Multiple Worksheets & Workbooks

#### Week 2
Text and Date Functions
- F4: Toggle between relative and absolute reference
- CTRL +: Today()
- MID(text, start_num, num_chars): MID(A2,2,FIND(" ",A2))
- CONCAT(A5:A20)
- Insert a line breaker: A2&CHAR(10)&B2

#### Week 3
Named Ranges
- Zoom out less than 40% to see the name ranges.
- Formular to locate name manager. 
- Use Name box to define name ranges.
- Create from selection to create name ranges for all the columns. 

#### Week 4
Summarising Data
- COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2]…): E.g.: COUNTIFS(Account_Manager, "Connor Betts") or COUNTIFS(Order_Quantity, 50).
- SUMIFS involve extra colums by including criteria ranges
- Sparkines: Sparkline deisgn ribbon
- Advanced charts: Insert -> chart -> define the axis -> select data: legend, horizon and vertical data -> Change chart type(Combo, secondary axis)
- Trendline: linear trend and exponential trend

#### Week 5
Tables
- Sorting: Data -> sort; sort by color by date
- Total row offers a range of calculation.
- Filter: number filters: top 10, above average.
- Automation with table: when add new data, the aggregation will be automatically updated. 
- Subtotal: convert table back to ranges -> data ribbon: subtotal -> different level of summary will be generated. 

#### Week 6
Pivot Tables, Pivot Charts and Slicers 
- Create pivot table: Select the data -> insert pivot table. 2D pivot table gives more detailed break down of the summary.
- Summarize value by, and show values as percentag of (row) total
- Group by date, month, year or quarter
- Add filters to the pivot table
- Double click the entry of pivot table to check the data used in such calculation.
- Select cell in the pivot table -> analysis -> pivot chart
- Select cell in the pivot table -> analysis -> Slices: connect slices to different pivot tables.


## [Excel Skills for Business: intermediate ii](https://www.coursera.org/learn/excel-intermediate-2)
#### Week 1
Data Validation
- Advanced conditional formatting: new rule to highlight certain columns or the entire row
- Circle out invalid data: find and select -> go to special -> data validation
- Create the drop down list

#### Week 2
Conditional Logic
- IF(NOT(A5=1), B5*5%,0)
- IF(Exact(A5, "JOHN"), B5*5%,0)

#### Week 3
Automating Lookups
- VLOOKUP: VLOOKUP(D4,$G$4:$H$7,2)
- INDEX
- CHOOSE
- MATCH

#### Week 4
Formula Auditing and Protection
- Dependent Cell
- #DIV/0!	An Excel error which occurs when you have divided by 0 (zero) or an empty cell.
- #NAME?	An Excel error which occurs when your formula contains unrecognisable text.
- #VALUE!	An Excel error which occurs when a value in your formula is of the wrong data type.
- Evaluate formular: step by step inspect the calculation of a formula

#### Week 5
Data Modelling
- SumProduct
- DataTable: at most two variables. 1, create the formula 2, whatif analysis 3, row or column input: {Table()}
- GoalSeek: arbitrary number of variable with constrains
- Scenario manager: summary of scenario manager will output different scenarios
- Solver

#### Week 6
Recording Macros
- VBA
- Create, edit Macro


## [Excel Skills for Business: Advanced](https://www.coursera.org/learn/excel-advanced)
#### Week 1
- Spreadsheet Design Principles
- Table
- NameRanges

#### Week 2
- Tables and structural referencing
- Array formular
- CountIf: {=SUM((USD_Price<Old_USD_Price)*1)} 


#### Week 3
- Replacing blanks with repeating value
- Date function
- Removing unwanted spaces (trim and clean)
- Removing unwanted character (Substitute, value and CHAR)

#### Week 4
- Financial functions (NPV, PV, FV)
- Rate of return (IRR)
- Depreciation Functions (SLN, SYD, DDB)

#### Week 5
- Relative reference: (INDIRECT, ADDRESS, OFFSET)

#### Week 6
- Dashborad design
- Interactive dashboard

