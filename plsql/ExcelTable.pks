create or replace package ExcelTable is
/* ======================================================================================

  MIT License

  Copyright (c) 2016-2021 Marc Bleron

  Permission is hereby granted, free of charge, to any person obtaining a copy
  of this software and associated documentation files (the "Software"), to deal
  in the Software without restriction, including without limitation the rights
  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  copies of the Software, and to permit persons to whom the Software is
  furnished to do so, subject to the following conditions:

  The above copyright notice and this permission notice shall be included in all
  copies or substantial portions of the Software.

  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  SOFTWARE.

=========================================================================================
    Change history :
    Marc Bleron       2016-05-01     Creation
    Marc Bleron       2016-06-25     Added string_cache.delete on tableClose
                                     Added lob freeing
    Marc Bleron       2016-09-11     Added ECDS_MAX_OFFSET in function Zip_openArchive
    Marc Bleron       2016-10-30     New streaming read method for large files (requires 
                                     Java)
                                     Added setFetchSize() procedure
    Marc Bleron       2017-05-14     Fixed ORA-06531 when file has no sharedStrings
    Marc Bleron       2017-05-28     Added support for password-encrypted workbooks
    Marc Bleron       2017-06-11     Added getCursor() function
    Marc Bleron       2017-07-10     Fixed bug when accessing zip archive created with 
                                     data descriptors. (Un)compressed sizes and CRC-32
                                     are now read from central directory entries
    Marc Bleron       2017-07-14     Support for long identifiers
    Marc Bleron       2017-12-31     Added cell comments extraction
    Marc Bleron       2018-02-04     Added support for .xls files
    Marc Bleron       2018-03-17     Added support for large shared strings in versions 
                                     prior 11.2.0.2
    Marc Bleron       2018-04-21     Added support for .xlsb files
    Marc Bleron       2018-04-24     Unzip utility : detection of stored files (no comp)
    Marc Bleron       2018-05-12     Added support for ODF spreadsheets (.ods)
    Marc Bleron       2018-08-22     new API for DML operations
    Marc Bleron       2018-11-02     Added multi-sheet support
    Marc Bleron       2019-03-31     Added default value feature to DML API
    Marc Bleron       2019-04-02     Added support for XML spreasheetML files
    Marc Bleron       2019-05-12     Fix : requested rows count wrongly decremented for 
                                           empty row
                                     Fix : getCursor() failure with multi-sheet support
    Marc Bleron       2019-05-21     Added flat file support
    Marc Bleron       2019-08-25     Fix : Exception when loading a zero or negative 
                                           value as a NUMBER(p) or NUMBER(p,s)
                                     Fix : DML API : Default value not applied for 
                                           empty cells
    Marc Bleron       2019-09-26     Fix : Error when using FOR ORDINALITY with a 
                                           positional text data source
    Marc Bleron       2019-10-02     Added strict OOXML support
    Marc Bleron       2019-11-03     Added streaming read method for ODF spreadsheets
    Marc Bleron       2020-02-29     Added cellNote attribute to ExcelTableCell
    Marc Bleron       2021-02-12     Fix : wrong value for cells containing an empty 
                                           shared string
====================================================================================== */

  -- Read methods  
  DOM_READ               constant pls_integer := 0;
  STREAM_READ            constant pls_integer := 1;
  STREAM_READ_XDB        constant pls_integer := 2; -- beta feature

  -- Metadata constants
  META_ORDINALITY        constant pls_integer := 0;
  META_COMMENT           constant pls_integer := 2;
  META_SHEET_NAME        constant pls_integer := 8;
  META_SHEET_INDEX       constant pls_integer := 16;

  -- DML operation types
  DML_INSERT             constant pls_integer := 0;
  DML_UPDATE             constant pls_integer := 1;
  DML_MERGE              constant pls_integer := 2;
  DML_DELETE             constant pls_integer := 3;

  subtype DMLContext is binary_integer;

  function getFile (
    p_directory in varchar2
  , p_filename  in varchar2
  ) 
  return blob;
  
  function getTextFile (
    p_directory in varchar2
  , p_filename  in varchar2
  , p_charset   in varchar2 default 'CHAR_CS'
  ) 
  return clob;
  
  procedure setDebug (p_status in boolean);
  procedure setFetchSize (p_nrows in number);
  procedure useSheetPattern (p_state in boolean);
    
  function createDMLContext (
    p_table_name in varchar2    
  )
  return DMLContext;
  
  procedure mapColumn (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in anydata default null
  );

  procedure mapColumnWithDefault (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in varchar2
  );
   
  procedure mapColumnWithDefault (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in number
  );
  
  procedure mapColumnWithDefault (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in date
  );
  
  function loadData (
    p_ctx       in DMLContext
  , p_file      in blob
  , p_sheet     in varchar2
  , p_range     in varchar2 default null
  , p_method    in binary_integer default DOM_READ
  , p_password  in varchar2 default null
  , p_dml_type  in pls_integer default DML_INSERT
  , p_err_log   in varchar2 default null
  )
  return integer;

  function loadData (
    p_ctx       in DMLContext 
  , p_file      in blob
  , p_sheets    in ExcelTableSheetList 
  , p_range     in varchar2 default null
  , p_method    in binary_integer default DOM_READ
  , p_password  in varchar2 default null
  , p_dml_type  in pls_integer default DML_INSERT
  , p_err_log   in varchar2 default null
  )
  return integer;
  
  function loadData (
    p_ctx        in DMLContext 
  , p_file       in clob
  , p_skip       in pls_integer
  , p_line_term  in varchar2
  , p_field_sep  in varchar2 default null
  , p_text_qual  in varchar2 default null
  , p_dml_type   in pls_integer default DML_INSERT
  , p_err_log    in varchar2 default null
  )
  return integer;
  
  /*
  EBNF grammar for the range_expr and column_list expression

    range_expr ::= ( cell_ref [ ":" cell_ref ] | col_ref ":" col_ref | row_ref ":" row_ref )
    cell_ref   ::= col_ref row_ref
    col_ref    ::= { "A".."Z" }
    row_ref    ::= integer
  
    column_list    ::= column_expr { "," column_expr }
    column_expr    ::= ( identifier datatype [ "column" string_literal ] [ for_metadata ]
                       | identifier for_ordinality )
    datatype       ::= ( number_expr | varchar2_expr | date_expr | timestamp_expr | clob_expr )
    number_expr    ::= "number" [ "(" ( integer | "*" ) [ "," integer ] ")" ]
    varchar2_expr  ::= "varchar2" "(" integer [ "char" | "byte" ] ")"
    date_expr      ::= "date" [ "format" string_literal ]
    timestamp_expr ::= "timestamp" [ "(" integer ")" ] [ "format" string_literal ]
    clob_expr      ::= "clob"
    for_ordinality ::= "for" "ordinality"
    for_metadata   ::= "for" "metadata" "(" ( "comment" | "formula" ) ")"
    identifier     ::= "\"" { char } "\""
    string_literal ::= "'" { char } "'"
  
  */
  
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

  function getRows (
    p_file     in blob
  , p_sheets   in ExcelTableSheetList
  , p_cols     in varchar2
  , p_range    in varchar2 default null
  , p_method   in binary_integer default DOM_READ
  , p_password in varchar2 default null
  ) 
  return anydataset pipelined
  using ExcelTableImpl;

  function getRows (
    p_file      in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer
  , p_line_term in varchar2
  , p_field_sep in varchar2 default null
  , p_text_qual in varchar2 default null
  ) 
  return anydataset pipelined
  using ExcelTableImpl;

  function getRawCells (
    p_file         in blob
  , p_sheetFilter  in anydata
  , p_cols         in varchar2
  , p_range        in varchar2 default null
  , p_method       in binary_integer default DOM_READ
  , p_password     in varchar2 default null
  )
  return ExcelTableCellList pipelined;
  
  function getCursor (
    p_file     in blob
  , p_sheet    in varchar2
  , p_cols     in varchar2
  , p_range    in varchar2 default null
  , p_method   in binary_integer default DOM_READ
  , p_password in varchar2 default null    
  )
  return sys_refcursor;

  function getCursor (
    p_file     in blob
  , p_sheets   in ExcelTableSheetList
  , p_cols     in varchar2
  , p_range    in varchar2 default null
  , p_method   in binary_integer default DOM_READ
  , p_password in varchar2 default null    
  )
  return sys_refcursor;

  function getCursor (
    p_file      in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer
  , p_line_term in varchar2
  , p_field_sep in varchar2 default null
  , p_text_qual in varchar2 default null    
  )
  return sys_refcursor;
    
  procedure tableDescribe (
    rtype    out nocopy anytype
  , p_range  in  varchar2
  , p_cols   in  varchar2
  , p_ff     in  boolean default false
  );

  function tablePrepare(
    tf_info  in  sys.ODCITabFuncInfo
  )
  return anytype;

  procedure tableStart (
    p_file         in  blob
  , p_sheetFilter  in  anydata
  , p_range        in  varchar2
  , p_cols         in  varchar2
  , p_method       in  binary_integer
  , p_ctx_id       out binary_integer
  , p_password     in  varchar2
  );

  procedure tableStart (
    p_file       in  clob
  , p_cols       in  varchar2
  , p_skip       in  pls_integer
  , p_line_term  in  varchar2
  , p_field_sep  in  varchar2
  , p_text_qual  in  varchar2
  , p_ctx_id     out binary_integer
  );

  procedure tableFetch(
    rtype   in out nocopy anytype
  , ctx_id  in binary_integer
  , nrows   in number
  , rws     out nocopy anydataset
  );
  
  procedure tableClose(
    p_ctx_id  in binary_integer
  );
  
  function getSheets (
    p_file         in blob
  , p_password     in varchar2 default null
  , p_method       in binary_integer default DOM_READ
  )
  return ExcelTableSheetList pipelined;

  function isReadMethodAvailable (
    p_method in binary_integer
  )
  return boolean;
  
end ExcelTable;
/
