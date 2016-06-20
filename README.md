# ExcelTable - An Excel Open Office XML SQL Interface

ExcelTable is a pipelined table interface to read an Excel file (.xlsx or .xlsm) as if it were an external table.
It is entirely implemented in PL/SQL using an object type (for the ODCI routines) and a package supporting the core functionalities.

## Bug tracker

Found bugs? I'm sure there are...
Please create an issue here on GitHub at <https://github.com/mbleron/oracle/ExcelTable>.

## Installation

ExcelTable package needs read access to V$PARAMETER view internally to retrieve the value of the `max_string_size` parameter.
Therefore, the owner must be granted the necessary privilege in order to compile and run the program : 
```sql
grant select on sys.v_$parameter to <user>;
```

## Usage

```sql
function getRows (
  p_file   in  blob
, p_sheet  in  varchar2
, p_cols   in  varchar2
, p_range  in  varchar2 default null
) 
return anydataset pipelined
using ExcelTableImpl;
```

* `p_file` : Input Excel file in Office Open XML format (.xlsx or .xlsm).
A helper function `ExcelTable.getFile` is available to directly reference the file from a directory.
* `p_sheet` : Worksheet name
* `p_cols` : Column list (see [specs](#columns-syntax-specification) below)
* `p_range` : Excel-like range expression that defines the table boundaries in the worksheet (see [specs](#range-syntax-specification) below)

#### <a name="colspecs"></a> Columns syntax specification

Column names must be declared using a quoted identifier.

Supported data types are :

* NUMBER – with optional precision and scale specs.

* VARCHAR2 – including CHAR/BYTE semantics. Values larger than the maximum length declared are silently truncated and no error is reported.

* DATE – with optional format mask. The format mask is used if the value is stored as text in the spreadsheet, otherwise the date value is assumed to be stored as date in Excel’s internal serial format.

* CLOB

A special "FOR ORDINALITY" clause (like XMLTABLE or JSON_TABLE’s one) is also available to autogenerate a sequence number.

Each column definition (except for the one qualified with FOR ORDINALITY) may be complemented with an optional "COLUMN" clause to explicitly target a named column in the spreadsheet, instead of relying on the order of the declarations (relative to the range).
Positional and named column definitions cannot be mixed.

For instance :

```
  "RN"    for ordinality
, "COL1"  number
, "COL2"  varchar2(10)
, "COL3"  varchar2(4000)
, "COL4"  date           format 'YYYY-MM-DD'
, "COL5"  number(10,2)
, "COL6"  varchar2(5)
```
or,
```
  "COL1"  number        column 'A'
, "COL2"  varchar2(10)  column 'C'
, "COL3"  clob          column 'D'
```


#### Range syntax specification

There are four ways to specify the table range :

* Range of rows : `'1:100'` – in this case the range of columns implicitly starts at A.
* Range of columns : `'B:E'` – in this case the range of rows implicitly starts at 1.
* Range of cells (top-left to bottom-right) : `'B2:F150'`
* Single cell anchor (top-left cell) : `'C3'`

> If the range is empty, the table implicitly starts at cell A1.


### Examples

Given this sample file : 

* Loading all six columns, starting at cell A2, in order to skip the header :

```
select t.*
from table(
       ExcelTable.getRows(
         ExcelTable.getFile('TMP_DIR','ooxdata3.xlsx')
       , 'DataSource'
       , ' "SRNO"    number
         , "NAME"    varchar2(10)
         , "VAL"     number
         , "DT"      date
         , "SPARE1"  varchar2(6)
         , "SPARE2"  varchar2(6)'
       , 'A2'
       )
     ) t
;
```

* Loading columns B and F only, from rows 2 to 10, with a generated sequence :

```
select t.*
from table(
       ExcelTable.getRows(
         ExcelTable.getFile('TMP_DIR','ooxdata3.xlsx')
       , 'DataSource'
       , q'{
           "R_NUM"   for ordinality
         , "NAME"    varchar2(10) column 'B'
         , "SPARE2"  varchar2(6)  column 'F'
         }'
       , '2:10'
       )
     ) t
;
```



## Author

**Marc Bleron**

## Copyright and license

Copyright 2016 Marc Bleron. Released under MIT license.
