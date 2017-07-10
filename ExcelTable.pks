create or replace package ExcelTable is
/* ======================================================================================

  MIT License

  Copyright (c) 2016,2017 Marc Bleron

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
====================================================================================== */

  -- Read methods  
  DOM_READ               constant binary_integer := 0;
  STREAM_READ            constant binary_integer := 1;

  procedure setFetchSize (p_nrows in number);
  
  /*
  EBNF grammar for the range_expr and column_list expression

    range_expr ::= ( cell_ref [ ":" cell_ref ] | col_ref ":" col_ref | row_ref ":" row_ref )
    cell_ref   ::= col_ref row_ref
    col_ref    ::= { "A".."Z" }
    row_ref    ::= integer
  
    column_list    ::= column_expr { "," column_expr }
    column_expr    ::= ( identifier datatype [ "column" string_literal ] | identifier for_ordinality )
    datatype       ::= ( number_expr | varchar2_expr | date_expr | clob_expr | for_ordinality )
    number_expr    ::= "number" [ "(" ( integer | "*" ) [ "," integer ] ")" ]
    varchar2_expr  ::= "varchar2" "(" integer [ "char" | "byte" ] ")"
    date_expr      ::= "date" [ "format" string_literal ]
    clob_expr      ::= "clob"
    for_ordinality ::= "for" "ordinality"
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
  
  function getCursor (
    p_file     in blob
  , p_sheet    in varchar2
  , p_cols     in varchar2
  , p_range    in varchar2 default null
  , p_method   in binary_integer default DOM_READ
  , p_password in varchar2 default null    
  )
  return sys_refcursor;
    
  procedure tableDescribe (
    rtype    out nocopy anytype
  , p_range  in  varchar2
  , p_cols   in  varchar2
  );

  function tablePrepare(
    tf_info  in  sys.ODCITabFuncInfo
  )
  return anytype;

  procedure tableStart (
    p_file     in  blob
  , p_sheet    in  varchar2
  , p_range    in  varchar2
  , p_cols     in  varchar2
  , p_method   in  binary_integer
  , p_ctx_id   out binary_integer
  , p_password in  varchar2
  );

  procedure tableFetch(
    p_type   in out nocopy anytype
  , p_ctx_id in out nocopy binary_integer
  , p_done   in out nocopy integer
  , nrows    in number
  , rws      out nocopy anydataset
  );
  
  procedure tableClose(
    p_ctx_id  in binary_integer
  );
  
  function getFile (
    p_directory in varchar2
  , p_filename  in varchar2
  ) 
  return blob;
  
end ExcelTable;
/
