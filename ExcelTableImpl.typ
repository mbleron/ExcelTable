create or replace type ExcelTableImpl as object (
/* ======================================================================================

  MIT License

  Copyright (c) 2016 Marc Bleron

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
====================================================================================== */

   atype     anytype
 , doc_id    raw(13)
 , ctx_id    raw(13)
 , r_num     integer
 , done      integer

 , static function ODCITableDescribe(
     rtype    out anytype
   , p_file   in  blob
   , p_sheet  in  varchar2
   , p_cols   in  varchar2
   , p_range  in  varchar2 default null
   ) 
   return number

 , static function ODCITablePrepare(
     sctx     out ExcelTableImpl
   , tf_info  in  sys.ODCITabFuncInfo
   , p_file   in  blob
   , p_sheet  in  varchar2
   , p_cols   in  varchar2
   , p_range  in  varchar2 default null
   )
   return number

 , static function ODCITableStart(
     sctx     in out ExcelTableImpl
   , p_file   in blob
   , p_sheet  in varchar2
   , p_cols   in varchar2
   , p_range  in varchar2 default null
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

)
/
create or replace type body ExcelTableImpl as

  static function ODCITableDescribe (
    rtype    out anytype
  , p_file   in  blob
  , p_sheet  in  varchar2
  , p_cols   in  varchar2
  , p_range  in  varchar2 default null
  )
  return number
  is
  begin
    
    ExcelTable.tableDescribe(rtype, p_range, p_cols);
    
    return ODCIConst.SUCCESS;
    
  end ODCITableDescribe;
  

  static function ODCITablePrepare (
    sctx     out ExcelTableImpl
  , tf_info  in  sys.ODCITabFuncInfo
  , p_file   in  blob
  , p_sheet  in  varchar2
  , p_cols   in  varchar2
  , p_range  in  varchar2 default null
  )
  return number
  is
  begin
    
    --dbms_output.put_line('ODCITablePrepare');

    sctx := ExcelTableImpl(
              ExcelTable.tablePrepare(tf_info)
            , null
            , null
            , 0
            , 0
            ) ;

    return ODCIConst.SUCCESS;
      
  end ODCITablePrepare;
  

  static function ODCITableStart (
    sctx     in out ExcelTableImpl
  , p_file   in blob
  , p_sheet  in varchar2
  , p_cols   in varchar2
  , p_range  in  varchar2 default null
  )
  return number
  is
  begin

    --dbms_output.put_line('ODCITableStart');
    ExcelTable.tableStart(p_file, p_sheet, p_range, p_cols, sctx.doc_id, sctx.ctx_id);

    return ODCIConst.SUCCESS;
    
  end ODCITableStart;
  

  member function ODCITableFetch (
    self   in out ExcelTableImpl
  , nrows  in     number
  , rws    out    anydataset
  )
  return number
  is
  begin
    
    --dbms_output.put_line('ODCITableFetch : '||nrows);
    
    ExcelTable.tableFetch(
      self.atype 
    , self.ctx_id
    , self.r_num
    , self.done
    , nrows
    , rws
    );
    
    return ODCIConst.SUCCESS;
     
  end ODCITableFetch;
  

  member function ODCITableClose
  return number
  is
  begin
   
    --dbms_output.put_line('ODCITableClose');
    ExcelTable.tableClose(self.doc_id, self.ctx_id);
    
    return ODCIConst.SUCCESS;
    
  end ODCITableClose;

end;
/
