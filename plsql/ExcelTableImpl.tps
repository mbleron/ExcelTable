create or replace type ExcelTableImpl as object (
/* ======================================================================================

  MIT License

  Copyright (c) 2016-2019 Marc Bleron

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
    Marc Bleron       2016-10-19     Removed doc_id attribute, 
                                     changed ctx_id to integer data type
    Marc Bleron       2017-05-28     Added p_password argument to ODCITableDescribe,
                                     -Prepare and -Start
    Marc Bleron       2017-07-22     Moved 'done' attribute to context cache
    Marc Bleron       2019-03-30     Multi-sheet support
    Marc Bleron       2019-05-21     Flat file support
====================================================================================== */

   atype     anytype
 , ctx_id    integer

 , static function ODCITableDescribe(
     rtype      out anytype
   , p_file     in  blob
   , p_sheet    in  varchar2
   , p_cols     in  varchar2
   , p_range    in  varchar2 default null
   , p_method   in  binary_integer default 0
   , p_password in  varchar2 default null
   ) 
   return number

 , static function ODCITableDescribe(
     rtype      out anytype
   , p_file     in  blob
   , p_sheets   in  ExcelTableSheetList
   , p_cols     in  varchar2
   , p_range    in  varchar2 default null
   , p_method   in  binary_integer default 0
   , p_password in  varchar2 default null
   ) 
   return number

 , static function ODCITablePrepare(
     sctx       out ExcelTableImpl
   , tf_info    in  sys.ODCITabFuncInfo
   , p_file     in  blob
   , p_sheet    in  varchar2
   , p_cols     in  varchar2
   , p_range    in  varchar2 default null
   , p_method   in  binary_integer default 0
   , p_password in  varchar2 default null
   )
   return number

 , static function ODCITablePrepare(
     sctx       out ExcelTableImpl
   , tf_info    in  sys.ODCITabFuncInfo
   , p_file     in  blob
   , p_sheets   in  ExcelTableSheetList
   , p_cols     in  varchar2
   , p_range    in  varchar2 default null
   , p_method   in  binary_integer default 0
   , p_password in  varchar2 default null
   )
   return number

 , static function ODCITableStart(
     sctx       in out ExcelTableImpl
   , p_file     in blob
   , p_sheet    in varchar2
   , p_cols     in varchar2
   , p_range    in varchar2 default null
   , p_method   in binary_integer default 0
   , p_password in varchar2 default null
   ) 
   return number

 , static function ODCITableStart(
     sctx       in out ExcelTableImpl
   , p_file     in blob
   , p_sheets   in  ExcelTableSheetList
   , p_cols     in varchar2
   , p_range    in varchar2 default null
   , p_method   in binary_integer default 0
   , p_password in varchar2 default null
   ) 
   return number

 , member function ODCITableFetch(
     self   in out ExcelTableImpl
   , nrows  in     number
   , rws    out    anydataset
   )
   return number

 , member function ODCITableClose
   return number

   -- Flat file
 , static function ODCITableDescribe(
     rtype        out anytype
   , p_file       in clob
   , p_cols       in varchar2
   , p_skip       in pls_integer
   , p_line_term  in varchar2
   , p_field_sep  in varchar2 default null
   , p_text_qual  in varchar2 default null
   ) 
   return number

 , static function ODCITablePrepare(
     sctx         out ExcelTableImpl
   , tf_info      in  sys.ODCITabFuncInfo
   , p_file       in clob
   , p_cols       in varchar2
   , p_skip       in pls_integer
   , p_line_term  in varchar2
   , p_field_sep  in varchar2 default null
   , p_text_qual  in varchar2 default null
   )
   return number

 , static function ODCITableStart(
     sctx         in out ExcelTableImpl
   , p_file       in clob
   , p_cols       in varchar2
   , p_skip       in pls_integer
   , p_line_term  in varchar2
   , p_field_sep  in varchar2 default null
   , p_text_qual  in varchar2 default null
   ) 
   return number

)
not final
/
