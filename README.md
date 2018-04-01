# ExcelTable - An Oracle SQL Interface for MS Excel Files

ExcelTable is a pipelined table interface to read an Excel file (.xlsx, .xlsm, .xls) as if it were an external table.
It is entirely implemented in PL/SQL using an object type (for the ODCI routines) and a package supporting the core functionalities.

> As of version 1.2, a streaming implementation is available for better scalability on large files. 
> This feature requires the server-side Java VM.

> As of version 1.3, ExcelTable can read password-encrypted files.

> As of version 2.0, ExcelTable can read old 97-2003 Excel files (.xls)

## Bug tracker

Found bugs? I'm sure there are...  
Please create an issue [here](https://github.com/mbleron/ExcelTable/issues).

## Installation

### Database requirement

ExcelTable requires Oracle Database 11\.2\.0\.1 and onwards.
> Note that the interface may work as well on versions 11\.1\.0\.6 and 11\.1\.0\.7, with limited support for CLOB projections, but that scenario has not been tested.

### DBA preliminary tasks

On database versions prior to 11\.2\.0\.4, a temporary XMLType table is used internally.
The owner requires the CREATE TABLE privilege in this case : 
```sql
grant create table to <user>;
```

In order to read encrypted Excel files, the interface requires access to the DBMS_CRYPTO API (see PL/SQL section below).  
The owner must therefore be granted EXECUTE privilege on it : 
```sql
grant execute on sys.dbms_crypto to <user>;
```


### PL/SQL

Create the following objects, in this order : 

> As of version 2.0, the following packages are now **mandatory** dependencies :  
> [XUTL_CDF](https://github.com/mbleron/MSUtilities/tree/master/CDFReader) : CFBF (OLE2) file reader  
> [XUTL_OFFCRYPTO](https://github.com/mbleron/MSUtilities/tree/master/OfficeCrypto) : Office crypto routines  
> XUTL_XLS for reading .xls files

```
@xutl_cdf.pks
@xutl_cdf.pkb
@xutl_offcrypto.pks
@xutl_offcrypto.pkb
```

```
@ExcelTableCell.tps
@ExcelTableCellList.tps
@xutl_xls.pks
@xutl_xls.pkb
@ExcelTableImpl.tps
@ExcelTable.pks
@ExcelTable.pkb
@ExcelTableImpl.tpb
```


### Java

If you want to use the streaming method, some Java classes - packed in a jar file - have to be deployed in the database.  
The jar files to deploy depend on the database version.

* Versions < 11\.2\.0\.4  
Except for version 11\.2\.0\.4 which supports JDK 6, Oracle 11g only supports JDK 5 (Java 1.5).
Load the following jar files in order to use the streaming method : 
  + stax-api-1.0-2.jar  
  + sjsxp-1.0.2.jar  
  + exceldbtools-1.5.jar

```
loadjava -u user/passwd@sid -r -v -jarsasdbobjects java/lib/stax-api-1.0-2.jar
loadjava -u user/passwd@sid -r -v -jarsasdbobjects java/lib/sjsxp-1.0.2.jar
loadjava -u user/passwd@sid -r -v -jarsasdbobjects java/lib/exceldbtools-1.5.jar
```


* Versions >= 11\.2\.0\.4  
The StAX API is included in JDK 6, as well as the Sun Java implementation (SJXSP), so for those versions one only needs to load the following jar file :  
  + exceldbtools-1.6.jar

```
loadjava -u user/passwd@sid -r -v -jarsasdbobjects java/lib/exceldbtools-1.6.jar
```

## Usage

```sql
function getRows (
  p_file     in blob
, p_sheet    in varchar2
, p_cols     in varchar2
, p_range    in varchar2 default null
, p_method   in binary_integer default DOM_READ
, p_password in varchar2 default null
) 
return anydataset pipelined
using ExcelTableImpl;
```

* `p_file` : Input Excel file (.xlsx, .xlsm or .xls format).
A helper function `ExcelTable.getFile` is available to directly reference the file from a directory.
* `p_sheet` : Worksheet name
* `p_cols` : Column list (see [specs](#columns-syntax-specification) below)
* `p_range` : Excel-like range expression that defines the table boundaries in the worksheet (see [specs](#range-syntax-specification) below)
* `p_method` : Read method - `DOM_READ` (0) the default, or `STREAM_READ` (1). The parameter value is ignored when reading from a .xls file.
* `p_password` : Optional - password used to encrypt the Excel document


## 
```sql
procedure setFetchSize (p_nrows in number);
```
Use setFetchSize() to control the number of rows returned by each invocation of the ODCITableFetch method.  
If the number of rows requested by the client is greater than the fetch size, the fetch size is used instead.  
The default fetch size is 100.  


## 
```sql
function getCursor (
  p_file     in blob
, p_sheet    in varchar2
, p_cols     in varchar2
, p_range    in varchar2 default null
, p_method   in binary_integer default DOM_READ
, p_password in varchar2 default null    
)
return sys_refcursor;
```
getCursor() returns a REF cursor allowing the consumer to iterate through the resultset returned by an equivalent getRows() call.  
It may be useful in PL/SQL code where static reference to table function returning ANYDATASET is not supported.  

## 
#### Columns syntax specification

Column names must be declared using a quoted identifier.

Supported data types are :

* NUMBER â€“ with optional precision and scale specs.

* VARCHAR2 â€“ including CHAR/BYTE semantics. Values larger than the maximum length declared are silently truncated and no error is reported.

* DATE â€“ with optional format mask. The format mask is used if the value is stored as text in the spreadsheet, otherwise the date value is assumed to be stored as date in Excelâ€™s internal serial format.

* CLOB

A special "FOR ORDINALITY" clause (like XMLTABLE or JSON_TABLEâ€™s one) is also available to autogenerate a sequence number.

Each column definition (except for the one qualified with FOR ORDINALITY) may be complemented with an optional "COLUMN" clause to explicitly target a named column in the spreadsheet, instead of relying on the order of the declarations (relative to the range).
Positional and named column definitions cannot be mixed.

New in version 1.6 : 
ExcelTable can also extract cell comments and project them as regular columns.  
In order to do so, the column syntax has been extended with a FOR METADATA clause : 

`FOR METADATA (COMMENT)`


Examples :

```
  "RN"    for ordinality
, "COL1"  number
, "COL2"  varchar2(10)
, "COL3"  varchar2(4000)
, "COL4"  date           format 'YYYY-MM-DD'
, "COL5"  number(10,2)
, "COL6"  varchar2(5)
```

```
  "COL1"  number        column 'A'
, "COL2"  varchar2(10)  column 'C'
, "COL3"  clob          column 'D'
```

```
"SPARE2"         varchar2(30)   column 'F'
"SPARE2_COMMENT" varchar2(2000) column 'F' for metadata (comment)
```

## 
#### Range syntax specification

There are four ways to specify the table range :

* Range of rows : `'1:100'` â€“ in this case the range of columns implicitly starts at A.
* Range of columns : `'B:E'` â€“ in this case the range of rows implicitly starts at 1.
* Range of cells (top-left to bottom-right) : `'B2:F150'`
* Single cell anchor (top-left cell) : `'C3'`

> If the range is empty, the table implicitly starts at cell A1.

## 
#### Cryptographic features overview

By default, Office 97-2003 password-protected files use RC4 encryption.  
Latest versions (2007+) based on [ECMA-376](http://www.ecma-international.org/publications/standards/Ecma-376.htm) standard use AES encryption : 

| Office version  | Method  | Encryption | Hash algorithm | Block chaining  
| :-------------- | :-----  | :--------- | :------------- | :-------------
| 97-2003         | RC4     | RC4        | MD5            | -
| 2007            | Standard| AES-128    | SHA-1          | ECB
| 2010            | Agile   | AES-128    | SHA-1          | CBC
| 2013            | Agile   | AES-256    | SHA512         | CBC
| 2016            | Agile   | AES-256    | SHA512         | CBC

Oracle, through DBMS_CRYPTO API, only supports SHA-2 algorithms (SHA256, 384, 512) starting from 12c.  
Therefore, in prior versions, the [OfficeCrypto](https://github.com/mbleron/MSUtilities/tree/master/OfficeCrypto) implementation cannot read Office 2013 (and onwards) documents encrypted with the default options.  

Full specs available on MSDN : [[MS-OFFCRYPTO]](https://msdn.microsoft.com/en-us/library/cc313071)  


## 
### Examples

Given this sample file : [ooxdata3.xlsx](./samples/ooxdata3.xlsx)

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

* Loading column C, starting at row 5, from a password-encrypted workbook ([crypto2016.xlsx](./samples/crypto2016.xlsx)) : 

```
SQL> select *
  2  from table(
  3         ExcelTable.getRows(
  4           ExcelTable.getFile('TMP_DIR','crypto2016.xlsx')
  5         , 'Feuil1'
  6         , '"COL1" number'
  7         , 'C5'
  8         , 0
  9         , p_password => 'AZE'
 10         )
 11       )
 12  ;
 
      COL1
----------
         1
         2
         3
 
```  

* Loading first three columns, row 1 to 91, from a password-encrypted .xls workbook ([ooxdata2c.xls](./samples/ooxdata2c.xls)) : 

```
SQL> select t.srno
  2       , t.name
  3       , t.content
  4       , length(t.content) as content_length
  5  from table(
  6         ExcelTable.getRows(
  7           ExcelTable.getFile('TMP_DIR','ooxdata2c.xls')
  8         , 'DataSource'
  9         , ' "SRNO"    number
 10           , "NAME"    varchar2(10)
 11           , "CONTENT" clob'
 12         , '1:91'
 13         , p_method   => null
 14         , p_password => 'pass123'
 15         )
 16       ) t
 17  ;
 
      SRNO NAME       CONTENT                                             CONTENT_LENGTH
---------- ---------- --------------------------------------------------- --------------
         1 LINE-00001 ABCD                                                             4
         2 LINE-00002 ABC                                                              3
         3 LINE-00003 ABC                                                              3
         4 LINE-00004 ABC                                                              3
         5 LINE-00005 ABC                                                              3
         6 LINE-00006 ABC                                                              3
         7 LINE-00007 ABC                                                              3
         8 LINE-00008 €XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX           8000
         9 LINE-00009 €€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€€           4000
        10 LINE-00010 10LINE-00010                                                    12
        11 LINE-00011 €XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX          12000
        12 LINE-00012 ABC                                                              3
        13 LINE-00013 ABC                                                              3
        14 LINE-00014 ABC                                                              3
...
        89 LINE-00089 ABC                                                              3
        90 LINE-00090 ABC                                                              3
        91 LINE-00091 ABC                                                              3
 
89 rows selected.
 
```  

* Retrieving a REF cursor for query #1 : 

```
SQL> var rc refcursor
SQL>
SQL> begin
  2    :rc :=
  3    ExcelTable.getCursor(
  4      p_file  => ExcelTable.getFile('TMP_DIR','ooxdata3.xlsx')
  5    , p_sheet => 'DataSource'
  6    , p_cols  => '"SRNO" number, "NAME" varchar2(10), "VAL" number, "DT" date, "SPARE1" varchar2(6), "SPARE2" varchar2(6)'
  7    , p_range => 'A2'
  8    );
  9  end;
 10  /

PL/SQL procedure successfully completed.

SQL> print rc

      SRNO NAME              VAL DT        SPARE1 SPARE2
---------- ---------- ---------- --------- ------ ------
         1 LINE-00001 66916.2986 13-OCT-23
         2 LINE-00002 96701.3427 05-SEP-06
         3 LINE-00003 68778.8698 23-JAN-11        OK
         4 LINE-00004  95110.028 03-MAY-07        OK
         5 LINE-00005 62561.5708 04-APR-27
         6 LINE-00006 28677.1166 11-JUL-23        OK
         7 LINE-00007 16141.0202 20-NOV-02
         8 LINE-00008 80362.6256 19-SEP-10
         9 LINE-00009 10384.1973 16-JUL-02
        10 LINE-00010  5266.9097 08-AUG-21
        11 LINE-00011 12513.0679 01-JUL-08
        12 LINE-00012 66596.9707 22-MAR-13
...
        97 LINE-00097 19857.7661 16-FEB-09
        98 LINE-00098 19504.3969 05-DEC-17
        99 LINE-00099 98675.8673 05-JUN-06
       100 LINE-00100 24288.2885 22-JUL-20

100 rows selected.

```  

* Extracting column F (value and cell comment) : 

```
SQL> select t.*
  2  from table(
  3         ExcelTable.getRows(
  4           ExcelTable.getFile('TMP_DIR','ooxdata3.xlsx')
  5         , 'DataSource'
  6         , q'{
  7             "RN"             for ordinality
  8           , "SPARE2"         varchar2(30)   column 'F'
  9           , "SPARE2_COMMENT" varchar2(2000) column 'F' for metadata (comment)
 10           }'
 11         , '2:11'
 12         )
 13       ) t
 14  ;

 RN SPARE2 SPARE2_COMMENT
--- ------ ------------------------------
  1
  2
  3 OK     bleronm:
           This is a comment.

  4 OK
  5
  6 OK
  7
  8
  9        This is
           another comment
           on three lines

 10

10 rows selected.

```


## CHANGELOG
### 2.0 (2018-04-01)
* Added support for Excel 97-2003 files (.xls)

### 1.6.1 (2018-03-17)
* Added large strings support for versions prior 11.2.0.2 

### 1.6 (2017-12-31)

* Added cell comments extraction
* Internal modularization

### 1.5 (2017-07-10)

* Fixed bug related to zip archives created with data descriptors. Now reading CRC-32, compressed and uncompressed sizes directly from Central Directory entries.
* Removed dependency to V$PARAMETER view (thanks [Paul](https://paulzipblog.wordpress.com/) for the suggestion)

### 1.4 (2017-06-11)

* Added getCursor() function
* Fixed NullPointerException when using streaming method and file has no sharedStrings

### 1.3 (2017-05-30)

* Added support for password-encrypted files
* Fixed minor bugs

### 1.2 (2016-10-30)

* Added new streaming read method
* Added setFetchSize() procedure

### 1.1 (2016-06-25)

* Added internal collection and LOB freeing


### 1.0 (2016-05-01)

* Creation



## Copyright and license

Copyright 2016-2018 Marc Bleron. Released under MIT license.
