create or replace package xutl_xls is
/* ======================================================================================

  MIT License

  Copyright (c) 2018-2019 Marc Bleron

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
    Marc Bleron       2018-02-04     Creation
    Marc Bleron       2018-10-17     New buffered LOB reader
    Marc Bleron       2018-10-28     Multi-sheet support
====================================================================================== */
  
  procedure set_debug (p_mode in boolean);
  
  function new_context (
    p_file      in blob 
  , p_password  in varchar2 default null
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null   
  )
  return pls_integer;
  
  procedure free_context (
    p_ctx_id  in pls_integer 
  );

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer default null
  )
  return ExcelTableCellList;

  function get_sheetList (
    p_ctx_id  in pls_integer 
  )
  return ExcelTableSheetList;

  procedure add_sheets (
    p_ctx_id     in pls_integer
  , p_sheetList  in ExcelTableSheetList
  );

  function get_comments (
    p_ctx_id     in pls_integer
  , p_sheetName  in varchar2
  )
  return ExcelTableCellList;
  
  function getRows (
    p_file      in blob 
  --, p_sheet     in varchar2
  , p_password  in varchar2 default null
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null
  )
  return ExcelTableCellList
  pipelined;
  
  --procedure read_all (p_wb in blob);

end xutl_xls;
/
