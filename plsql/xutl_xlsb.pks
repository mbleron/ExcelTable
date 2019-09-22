create or replace package xutl_xlsb is
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
    Marc Bleron       2018-04-02     Creation
    Marc Bleron       2018-08-23     Bug fix : no row returned if fetch_size (p_nrows)
                                     is less than first row index
    Marc Bleron       2018-09-02     Multi-sheet support
====================================================================================== */
  
  type SheetEntry_T is record (name varchar2(31 char), relId varchar2(255 char));
  type SheetEntries_T is table of SheetEntry_T;
 
  procedure set_debug (p_mode in boolean);
    
  function new_context (
    p_sst_part    in blob
  , p_cols        in varchar2 default null
  , p_firstRow    in pls_integer default null
  , p_lastRow     in pls_integer default null    
  )
  return pls_integer;

  procedure add_sheet (
    p_ctx_id   in pls_integer
  , p_content  in blob
  );
  
  function get_sheetEntries (
    p_workbook  in blob
  )
  return SheetEntries_T;
 
  procedure free_context (
    p_ctx_id  in pls_integer 
  );

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer
  )
  return ExcelTableCellList;

  function get_comments (
    p_comments  in blob
  )
  return ExcelTableCellList;
  
/*  function get_rows (
    p_sheet_part  in blob 
  , p_sst_part    in blob
  , p_cols        in varchar2 default null
  , p_firstRow    in pls_integer default null
  , p_lastRow     in pls_integer default null
  )
  return ExcelTableCellList
  pipelined;*/

end xutl_xlsb;
/
