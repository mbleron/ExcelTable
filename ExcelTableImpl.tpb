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
