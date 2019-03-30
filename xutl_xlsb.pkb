create or replace package body xutl_xlsb is
 
  -- Binary Record Types
  BRT_ROWHDR          constant pls_integer := 0; 
  BRT_CELLBLANK       constant pls_integer := 1;
  BRT_CELLRK          constant pls_integer := 2;
  BRT_CELLERROR       constant pls_integer := 3;
  BRT_CELLBOOL        constant pls_integer := 4;
  BRT_CELLREAL        constant pls_integer := 5;
  BRT_CELLST          constant pls_integer := 6; 
  BRT_CELLISST        constant pls_integer := 7;
  BRT_FMLASTRING      constant pls_integer := 8;
  BRT_FMLANUM         constant pls_integer := 9;
  BRT_FMLABOOL        constant pls_integer := 10;
  BRT_FMLAERROR       constant pls_integer := 11;
  BRT_ENDBUNDLESHS    constant pls_integer := 144;
  --BRT_BEGINSHEETDATA  constant pls_integer := 145;
  BRT_ENDSHEETDATA    constant pls_integer := 146;  
  BRT_BUNDLESH        constant pls_integer := 156;
  BRT_BEGINSST        constant pls_integer := 159;
  BRT_BEGINCOMMENTS   constant pls_integer := 628;
  BRT_ENDCOMMENTLIST  constant pls_integer := 634;
  BRT_BEGINCOMMENT    constant pls_integer := 635;
  BRT_COMMENTTEXT     constant pls_integer := 637;
  
  -- Error Types
  FT_ERR_NULL         constant raw(1) := '00';
  FT_ERR_DIV_ZERO     constant raw(1) := '07';
  FT_ERR_VALUE        constant raw(1) := '0F';
  FT_ERR_REF          constant raw(1) := '17';
  FT_ERR_NAME         constant raw(1) := '1D';
  FT_ERR_NUM          constant raw(1) := '24';
  FT_ERR_NA           constant raw(1) := '2A';
  FT_ERR_GETDATA      constant raw(1) := '2B';
  
  -- Boolean Values
  BOOL_FALSE          constant raw(1) := '00';
  BOOL_TRUE           constant raw(1) := '01';
  
  -- String Types
  ST_SIMPLE           constant pls_integer := 0;
  ST_RICHSTR          constant pls_integer := 1;
  
  ERR_EXPECTED_REC    constant varchar2(100) := 'Error at position %d, expecting a [%s] record';
  
  type RecordTypeMap_T is table of varchar2(128) index by pls_integer;
  recordTypeMap  RecordTypeMap_T;
  
  type BitMaskTable_T is varray(8) of raw(1);
  BITMASKTABLE    constant BitMaskTable_T := BitMaskTable_T('01','02','04','08','10','20','40','80');

  type ColumnMap_T is table of varchar2(2) index by pls_integer;
  
  type Range_T is record (
    firstRow  pls_integer
  , lastRow   pls_integer
  , colMap    ColumnMap_T
  );

  type Stream_T is record (
    content    blob
  , sz         integer
  , offset     integer
  , rt         pls_integer
  , hsize      pls_integer
  , rsize      binary_integer
  , rstart     integer
  , available  pls_integer
  );
  
  type String_T is record (
    is_lob    boolean := false
  , strValue  varchar2(32767)
  , lobValue  clob  
  );
  
  type XLRichString_T is record (
    cch       binary_integer
  , fExtStr   boolean
  , fRichStr  boolean
  , byte_len  pls_integer := 0
  , content   String_T
  );
  
  type String_Array_T is table of String_T;
  
  type SST_T is record (
    cstTotal   binary_integer
  , cstUnique  binary_integer
  , strings    String_Array_T
  );
  
  type BrtBundleSh_T is record (
    hsState   raw(4)
  , iTabID    raw(4)
  , strRelID  varchar2(255 char)
  , strName   varchar2(31 char)
  );
  
  type Comment_T is record (
    rw    pls_integer
  , col   pls_integer
  , text  varchar2(32767)
  );
  
  type RK_T is record (
    fX100     boolean
  , fInt      boolean
  , RkNumber  raw(4)
  );
  
  type SheetList_T is table of blob;
  
  type Context_T is record (
    stream      Stream_T
  , sst         SST_T
  , rng         Range_T
  , curr_rw     pls_integer
  , done        boolean
  , sheetList   SheetList_T
  , curr_sheet  pls_integer
  );
  
  type Context_cache_T is table of Context_T index by pls_integer;
  ctx_cache  Context_cache_T;
  
  debug_mode       boolean := false;
  MAX_STRING_SIZE  pls_integer;

  function raw2int (r in raw) return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(r, utl_raw.little_endian);
  end;

  function get_max_string_size 
  return pls_integer 
  is
    l_result  pls_integer;
  begin
    select lengthb(rpad('x',32767,'x')) 
    into l_result
    from dual;
    return l_result;
  end;

  procedure init 
  is
  begin
    MAX_STRING_SIZE := get_max_string_size();
    recordTypeMap(BRT_BEGINSST) := 'BrtBeginSst';
    recordTypeMap(BRT_BEGINCOMMENTS) := 'BrtBeginComments';
    recordTypeMap(BRT_COMMENTTEXT) := 'BrtCommentText';
    --recordTypeMap(raw2int(RT_STRING)) := 'String';
  end;

  procedure set_debug (p_mode in boolean)
  is
  begin
    debug_mode := p_mode;  
  end;
  
  procedure debug (message in varchar2)
  is
  begin
    if debug_mode then
      dbms_output.put_line(message);
    end if;
  end;

  procedure trace_lob (message in varchar2)
  is
    type lob_stats_rec is record (
      cache_lobs     number
    , nocache_lobs   number
    , abstract_lobs  number
    );
    lob_stats  lob_stats_rec;
  begin
    select cache_lobs, nocache_lobs, abstract_lobs
    into lob_stats
    from v$temporary_lobs
    where sid = sys_context('userenv','sid');
    debug('LOB Stats for step : '||message);
    debug('cache_lobs = '||lob_stats.cache_lobs);
    debug('nocache_lobs = '||lob_stats.nocache_lobs);
    debug('abstract_lobs = '||lob_stats.abstract_lobs);
    debug('----------------------------------------------');
  end;

  procedure error (
    errcode in pls_integer
  , message in varchar2
  , arg1    in varchar2 default null
  , arg2    in varchar2 default null
  , arg3    in varchar2 default null
  ) 
  is
  begin
    raise_application_error(errcode, utl_lms.format_message(message, arg1, arg2, arg3));
  end;

  procedure expect (
    stream  in Stream_T
  , rt      in pls_integer
  )
  is
  begin
    if stream.rt != rt then
      error(-20731, ERR_EXPECTED_REC, stream.rstart - stream.hsize, recordTypeMap(rt));
    end if;    
  end;
  
  function is_bit_set (
    byteVal  in raw
  , bitNum   in pls_integer
  )
  return boolean
  is
    bitmask  raw(1) := BITMASKTABLE(bitNum);
  begin
    return ( utl_raw.bit_and(byteVal, bitmask) = bitmask );
  end;

  function read_bytes (
    stream  in out nocopy Stream_T
  , amount  in pls_integer
  )
  return raw
  is
    bytes  raw(32767);
  begin
    bytes := dbms_lob.substr(stream.content, amount, stream.offset);
    stream.offset := stream.offset + amount;
    stream.available := stream.available - amount;
    return bytes;
  end;

  function read_int8 (
    stream  in out nocopy Stream_T
  )
  return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(read_bytes(stream, 1), utl_raw.little_endian);
  end;

  function read_int16 (
    stream  in out nocopy Stream_T
  )
  return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(read_bytes(stream, 2), utl_raw.little_endian);
  end;

  function read_int32 (
    stream  in out nocopy Stream_T
  )
  return binary_integer
  is
  begin
    return utl_raw.cast_to_binary_integer(read_bytes(stream, 4), utl_raw.little_endian);
  end;

  function open_stream (
    wbFile  in blob
  )
  return Stream_T
  is
    stream  Stream_T;
  begin
    stream.content := wbFile;
    stream.sz := dbms_lob.getlength(stream.content);
    stream.offset := 0;
    stream.rsize := 0;
    stream.rstart := 1;
    return stream;
  end;
  
  procedure close_stream (
    stream  in out nocopy Stream_T 
  )
  is
  begin
    dbms_lob.freetemporary(stream.content);
  end;

  procedure next_record (
    stream in out nocopy Stream_T
  ) 
  is
    --INT2_7   constant pls_integer := 128;
    --INT2_14  constant pls_integer := 16384;
    --INT2_21  constant pls_integer := 2097152;
    int8  pls_integer;
    rstart  integer;
  begin
    stream.offset := stream.rstart + stream.rsize;
    rstart := stream.offset;
    
    -- record type
    int8 := read_int8(stream);
    if int8 < 128 then
      stream.rt := int8;
    else
      stream.rt := bitand(int8,127);
      int8 := read_int8(stream);
      stream.rt := stream.rt + bitand(int8,127) * 128;
    end if;
    
    -- record size
    int8 := read_int8(stream); -- byte 1 
    stream.rsize := bitand(int8,127); -- lowest 7 bits
    if int8 >= 128 then  
      int8 := read_int8(stream); -- byte 2
      stream.rsize := stream.rsize + bitand(int8,127) * 128;
      if int8 >= 128 then
        int8 := read_int8(stream); -- byte 3
        stream.rsize := stream.rsize + bitand(int8,127) * 16384;
        if int8 >= 128 then
          int8 := read_int8(stream); -- byte 4
          stream.rsize := stream.rsize + bitand(int8,127) * 2097152;
        end if;
      end if;
    end if;
    
    stream.hsize := stream.offset - rstart;
    stream.available := stream.rsize;
    -- current record start
    stream.rstart := stream.offset;
    debug('RECORD INFO ['||to_char(stream.rstart,'FM09999999')||']['||lpad(stream.rsize,6)||'] '||stream.rt);
  end;

  procedure seek_first (
    stream       in out nocopy Stream_T
  , record_type  in raw  
  )
  is
  begin
    next_record(stream);
    while stream.offset < stream.sz and stream.rt != record_type loop
      next_record(stream);
    end loop;
  end;

  procedure seek (
    stream  in out nocopy Stream_T
  , pos     in integer
  )
  is
  begin
    stream.rstart := pos;
    stream.rsize := 0;
  end;
  
  procedure skip (
    stream  in out nocopy Stream_T
  , amount  in integer
  )
  is
  begin
    stream.offset := stream.offset + amount;
    stream.available := stream.available - amount;
  end;

  -- convert a 0-based column number to base26 string
  function base26encode (colNum in pls_integer) 
  return varchar2
  result_cache
  is
    output  varchar2(3);
    num     pls_integer := colNum + 1;
  begin
    if colNum is not null then
      while num != 0 loop
        output := chr(65 + mod(num-1,26)) || output;
        num := trunc((num-1)/26);
      end loop;
    end if;
    return output;
  end;

  function base26decode (colRef in varchar2)
  return pls_integer
  result_cache
  is
  begin
    return ascii(substr(colRef,-1,1))-65 
         + nvl((ascii(substr(colRef,-2,1))-64)*26, 0)
         + nvl((ascii(substr(colRef,-3,1))-64)*676, 0);
  end;

  function parseColumnList (
    cols  in varchar2
  , sep   in varchar2 default ','
  )
  return ColumnMap_T
  is
    colMap  Columnmap_t;
    i       pls_integer;
    token   varchar2(3);
    p1      pls_integer := 1;
    p2      pls_integer;  
  begin
    if cols is not null then
      loop
        p2 := instr(cols, sep, p1);
        if p2 = 0 then
          token := substr(cols, p1);
        else
          token := substr(cols, p1, p2-p1);    
          p1 := p2 + 1;
        end if;
        i := base26decode(token);
        colMap(i) := token;
        exit when p2 = 0;
      end loop;
    end if;
    return colMap; 
  end;

  procedure next_sheet (
    ctx  in out nocopy Context_T
  )
  is
    has_next  boolean := (ctx.curr_sheet < ctx.sheetList.count);
  begin
    if ctx.curr_sheet > 0 then
      close_stream(ctx.stream);
    end if;
    while has_next loop
      ctx.curr_sheet := ctx.curr_sheet + 1;
      debug('Switching to sheet '||ctx.curr_sheet);
      ctx.stream := open_stream(ctx.sheetList(ctx.curr_sheet));
      exit;
      --has_next := (ctx.curr_sheet < ctx.sheetList.count);
    end loop;
    if not has_next then
      debug('End of sheet list');
      ctx.done := true;
    end if;
  end;

  function read_Bool (stream in out nocopy Stream_T)
  return String_T
  is
    bBool  raw(1);
    str    String_T;
  begin
    bBool := read_bytes(stream, 1);
    str.strValue := case bBool 
                      when BOOL_TRUE then 'TRUE' 
                      when BOOL_FALSE then 'FALSE' 
                    end; 
    return str;
  end;

  function read_Err (stream in out nocopy Stream_T)
  return String_T
  is
    fErr  raw(1);
    str   String_T;
  begin
    fErr := read_bytes(stream, 1);
    str.strValue := 
      case fErr
            when FT_ERR_NULL then '#NULL!'
            when FT_ERR_DIV_ZERO then '#DIV/0!'
            when FT_ERR_VALUE then '#VALUE!'
            when FT_ERR_REF then '#REF!'
            when FT_ERR_NAME then '#NAME?'
            when FT_ERR_NUM then '#NUM!'
            when FT_ERR_NA then '#N/A'
            when FT_ERR_GETDATA then '#GETTING_DATA'
          end;     
     return str;
  end;
  
  function read_RK (stream in out nocopy Stream_T)
  return number
  is
    rk  RK_T;
    nm  number;
  begin
    rk.RkNumber := read_bytes(stream, 4);
    rk.fX100 := is_bit_set(utl_raw.substr(rk.RkNumber,1,1), 1);
    rk.fInt := is_bit_set(utl_raw.substr(rk.RkNumber,1,1), 2);
    
    rk.RkNumber := utl_raw.bit_and(rk.RkNumber, 'FCFFFFFF');
    if rk.fInt then 
      -- convert to int and shift right 2
      nm := to_number(utl_raw.cast_to_binary_integer(rk.RkNumber, utl_raw.little_endian)/4);
    else
      -- pad LSBs with 0 and convert to double
      nm := to_number(utl_raw.cast_to_binary_double(utl_raw.concat('00000000',rk.RkNumber),utl_raw.little_endian));
    end if;
    if rk.fX100 then
      nm := nm/100;
    end if;
    return nm;
  end;
  
  function read_Number (stream in out nocopy Stream_T)
  return number
  is
    Xnum  raw(8);
  begin
    Xnum := read_bytes(stream, 8);
    return to_number(utl_raw.cast_to_binary_double(Xnum, utl_raw.little_endian));
  end;

  function read_SSTItem (
    stream  in out nocopy Stream_T
  , sst  in SST_T 
  )
  return String_T
  is
    idx  pls_integer := read_int32(stream) + 1;
  begin
    return sst.strings(idx);
  end;

  procedure read_XLString (
    stream  in out nocopy Stream_T
  , str     in out nocopy XLRichString_T
  , sttype  in pls_integer default ST_SIMPLE
  )
  is
    raw1    raw(1);
    buf     raw(32764);
    cbuf    varchar2(32764);  
    rem     pls_integer;
    amt     pls_integer;
    csz     pls_integer := 2;
    csname  varchar2(30) := 'AL16UTF16LE';
  begin
    
    if stType = ST_RICHSTR then
      raw1 := read_bytes(stream, 1);
      str.fRichStr := is_bit_set(raw1, 1);
      str.fExtStr := is_bit_set(raw1, 2);
    end if;
    
    str.cch := read_int32(stream);
    
    if str.cch != -1 then
    
      rem := str.cch; -- characters left to read;
      
      while rem != 0 loop
        
        amt := least(8191, rem) * csz; -- byte amount to read      
        buf := read_bytes(stream, amt);
        cbuf := utl_i18n.raw_to_char(buf, csname);
        
        if str.content.is_lob then
        
          str.content.lobValue := str.content.lobValue || cbuf;
        
        else
          
          str.byte_len := str.byte_len + lengthb(cbuf);
          if str.byte_len > MAX_STRING_SIZE then
            -- switch to lob storage
            dbms_lob.createtemporary(str.content.lobValue, true);
            if str.content.strValue is not null then
              dbms_lob.writeappend(str.content.lobValue, length(str.content.strValue), str.content.strValue);
            end if;
            dbms_lob.writeappend(str.content.lobValue, length(cbuf), cbuf);
            str.content.is_lob := true;
            str.content.strValue := null;
          else       
            str.content.strValue := str.content.strValue || cbuf;        
          end if;  
        
        end if;
        
        rem := rem - amt/csz;
        
      end loop;
    
    end if;
        
  end;
  
  function read_XLString (
    stream  in out nocopy Stream_T
  , stType  in pls_integer default 0
  )
  return String_T
  is
    xlstr  XLRichString_T;
  begin
    read_XLString(stream, xlstr, stType);
    return xlstr.content;
  end;
  
  function read_SheetInfo (
    stream  in out nocopy Stream_T
  )
  return BrtBundleSh_T
  is
    sh  BrtBundleSh_T;
  begin
    sh.hsState := read_bytes(stream, 4);
    sh.iTabID := read_bytes(stream, 4);
    sh.strRelID := read_XLString(stream).strValue;
    debug(sh.strRelID);
    sh.strName := read_XLString(stream).strValue;
    debug(sh.strName);
    return sh;
  end;

  function get_sheetEntries (
    p_workbook  in blob
  )
  return SheetEntries_T
  is
    i             pls_integer := 0;
    stream        Stream_T;
    sh            BrtBundleSh_T;
    sheetEntries  SheetEntries_T := SheetEntries_T();
  begin

    stream := open_stream(p_workbook);

    next_record(stream);
    -- read records until BrtEndBundleShs is found
    while stream.rt != BRT_ENDBUNDLESHS loop    
      if stream.rt = BRT_BUNDLESH then
        i := i + 1;
        sheetEntries.extend;
        sh := read_SheetInfo(stream);
        sheetEntries(i).name := sh.strName;
        sheetEntries(i).relId := sh.strRelID;
      end if;   
      next_record(stream);
    end loop;
    
    close_stream(stream);
    
    return sheetEntries;
    
  end;

  function read_CellBlock (
    ctx   in out nocopy Context_T
  , nrows in pls_integer
  )
  return ExcelTableCellList
  is
    rcnt   pls_integer := 0;
    rw     pls_integer := ctx.curr_rw;
    col    pls_integer;
    num    number;
    str    String_T;
    
    cells  ExcelTableCellList := ExcelTableCellList();
    
    procedure read_col is
    begin
      col := read_int32(ctx.stream);
      skip(ctx.stream, 4); -- iStyleRef, fPhShow, reserved
    end;
    
    procedure add_cell(val in anydata) is
    begin
      if ctx.rng.colMap.count = 0 or ctx.rng.colMap.exists(col) then
        cells.extend;
        cells(cells.last) := new ExcelTableCell(rw + 1, base26encode(col), null, val, ctx.curr_sheet);
      end if;
    end;
    
  begin
  
    if rw is not null then
      rcnt := 1;
    end if;
    
    next_record(ctx.stream);
    
    loop
           
      if ctx.stream.rt = BRT_ENDSHEETDATA then
        --ctx.done := true;
        --exit;
        next_sheet(ctx);
        
      elsif ctx.stream.rt = BRT_ROWHDR then
        rw := read_int32(ctx.stream);
        debug('Row '||rw);
        
        if rw > ctx.rng.lastRow then
          debug('End of range');
          --ctx.done := true;
          --exit;
          next_sheet(ctx);
        
        elsif rw >= ctx.rng.firstRow then
          
          rcnt := rcnt + 1;
          if rcnt > nrows then
            debug('Batch size = '||to_char(rcnt-1));
            ctx.curr_rw := rw;
            exit;
          end if;
          
        end if;
        
      elsif rw >= ctx.rng.firstRow then
        
        case ctx.stream.rt
        when BRT_CELLBLANK then
        
          read_col;
        
        when BRT_CELLRK    then
          
          read_col;
          num := read_RK(ctx.stream);
          add_cell(anydata.ConvertNumber(num));
        
        when BRT_CELLERROR then
          
          read_col;
          str := read_Err(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
        
        when BRT_CELLBOOL  then
          
          read_col;
          str := read_Bool(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
        
        when BRT_CELLREAL  then
          
          read_col;
          num := read_Number(ctx.stream);
          add_cell(anydata.ConvertNumber(num));
          
        when BRT_CELLST    then
          
          read_col;
          str := read_XLString(ctx.stream);
          if str.is_lob then
            add_cell(anydata.ConvertClob(str.lobValue));
          else
            add_cell(anydata.ConvertVarchar2(str.strValue));
          end if;
        
        when BRT_CELLISST  then
          
          read_col;
          str := read_SSTItem(ctx.stream, ctx.sst);
          if str.is_lob then
            add_cell(anydata.ConvertClob(str.lobValue));
          else
            add_cell(anydata.ConvertVarchar2(str.strValue));
          end if;          
          
        when BRT_FMLASTRING  then
          
          read_col;
          str := read_XLString(ctx.stream);
          if str.is_lob then
            add_cell(anydata.ConvertClob(str.lobValue));
          else
            add_cell(anydata.ConvertVarchar2(str.strValue));
          end if;
        
        when BRT_FMLANUM  then
        
          read_col;
          num := read_Number(ctx.stream);
          add_cell(anydata.ConvertNumber(num));
        
        when BRT_FMLABOOL  then
        
          read_col;
          str := read_Bool(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
        
        when BRT_FMLAERROR  then
        
          read_col;
          str := read_Err(ctx.stream);
          add_cell(anydata.ConvertVarchar2(str.strValue));
          
        else
          null;
        end case;
      
      end if;
      
      exit when ctx.done;
      
      next_record(ctx.stream);
    
    end loop;
    
    return cells;
    
  end;

  procedure read_SST (
    sst_part in blob
  , sst      in out nocopy SST_T
  )
  is
    stream  Stream_T;
  begin
    stream := open_stream(sst_part);
    
    next_record(stream);
    expect(stream, BRT_BEGINSST);
    sst.cstTotal := read_int32(stream);
    sst.cstUnique := read_int32(stream);
    
    debug('sst.cstTotal = '||sst.cstTotal);
    debug('sst.cstUnique = '||sst.cstUnique);
    
    sst.strings := String_Array_T();
    sst.strings.extend(sst.cstUnique);
    
    for i in 1 .. sst.cstUnique loop
      next_record(stream);
      sst.strings(i) := read_XLString(stream, ST_RICHSTR);
      --debug(sst.strings(i).strValue);
    end loop;
    
    close_stream(stream);
  
  end;

  function read_Comment (
    stream  in out nocopy Stream_T 
  )
  return Comment_T
  is
    cmt  Comment_T;
  begin
    skip(stream, 4); -- iauthor
    cmt.rw := read_int32(stream); -- rwFirst
    skip(stream, 4); -- rwLast
    cmt.col := read_int32(stream); -- colFirst
    skip(stream, 4); -- colLast
    next_record(stream);
    expect(stream, BRT_COMMENTTEXT);
    cmt.text := read_XLString(stream, ST_RICHSTR).strValue;
    --debug(cmt.text);
    return cmt;
  end;

  procedure read_CommentList (
    stream       in out nocopy Stream_T
  , commentList  in out nocopy ExcelTableCellList
  )
  is
    cmt  Comment_T;
    i    pls_integer := 0;
  begin
    
    next_record(stream);
    expect(stream, BRT_BEGINCOMMENTS);
    
    -- read records until BrtEndCommentList is found
    while stream.rt != BRT_ENDCOMMENTLIST loop    
      if stream.rt = BRT_BEGINCOMMENT then
        cmt := read_Comment(stream);
        commentList.extend;
        i := i + 1;
        commentList(i) := ExcelTableCell(
                            cellRow => cmt.rw + 1
                          , cellCol => base26encode(cmt.col)
                          , cellType => null
                          , cellData => anydata.ConvertVarchar2(cmt.text)
                          , sheetIdx => null
                          );
      end if;
      next_record(stream);
    end loop;

  end;
  
  /*  
  procedure read_all (content in blob)
  is
    stream  Stream_T;
    rw      pls_integer;
  begin
    
    stream := open_stream(content);
    next_record(stream);
    
    while stream.rt != BRT_ENDSHEETDATA loop
      
      if stream.rt = BRT_ROWHDR then
        rw := read_int32(stream);
        debug('Row '||rw);
      end if;
      
      next_record(stream);
    
    end loop;
    close_stream(stream);
  
  end;
  */
  
  function new_context (
    p_sst_part  in blob
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null
  )
  return pls_integer
  is
    ctx     Context_T;
    ctx_id  pls_integer; 
  begin
    
    --read_SheetMap(p_wb_part, ctx.sheetMap);
  
    if p_sst_part is not null then
      read_sst(p_sst_part, ctx.sst);
    end if;
  
    ctx.rng.firstRow := nvl(p_firstRow, 1) - 1;
    ctx.rng.lastRow := nvl(p_lastRow, 1048576) - 1;
    ctx.rng.colMap := parseColumnList(p_cols);
    --ctx.stream := open_stream(p_sheet_part);
    ctx.done := false;
    ctx.sheetList := SheetList_T();
    ctx.curr_sheet := 0;
    
    ctx_id := nvl(ctx_cache.last, 0) + 1;
    ctx_cache(ctx_id) := ctx;
    
    return ctx_id;
    
  end;

  procedure add_sheet (
    p_ctx_id   in pls_integer
  , p_content  in blob
  )
  is
    i  pls_integer;
  begin
    ctx_cache(p_ctx_id).sheetList.extend;
    i := ctx_cache(p_ctx_id).sheetList.last;
    ctx_cache(p_ctx_id).sheetList(i) := p_content;
  end;

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer
  )
  return ExcelTableCellList
  is
    cells  ExcelTableCellList;
  begin
    if ctx_cache(p_ctx_id).curr_sheet = 0 then
      next_sheet(ctx_cache(p_ctx_id));
    end if;
    if not ctx_cache(p_ctx_id).done then
      cells := read_CellBlock(ctx_cache(p_ctx_id), p_nrows);
    end if;
    return cells;
  end;

  procedure free_context (
    p_ctx_id  in pls_integer 
  )
  is
  begin
    --close_stream(ctx_cache(p_ctx_id).stream);
    ctx_cache(p_ctx_id).sst.strings := String_Array_T();
    ctx_cache.delete(p_ctx_id);
  end;

  function get_comments (
    p_comments  in blob
  )
  return ExcelTableCellList
  is
    stream       Stream_T;
    commentList  ExcelTableCellList := ExcelTableCellList();
  begin
    stream := open_stream(p_comments);
    read_CommentList(stream, commentList);
    close_stream(stream);
    return commentList;
  end;
  

  /*function get_rows (
    p_sheet_part  in blob 
  , p_sst_part    in blob
  , p_cols        in varchar2 default null
  , p_firstRow    in pls_integer default null
  , p_lastRow     in pls_integer default null
  )
  return ExcelTableCellList
  pipelined
  is
    ctx_id  pls_integer;
    cells   ExcelTableCellList;   
  begin
    
    ctx_id := new_context(p_sheet_part, p_sst_part, p_cols, p_firstRow, p_lastRow);

    while not ctx_cache(ctx_id).done loop
      debug('NEXT BATCH');
      cells := read_CellBlock(ctx_cache(ctx_id), 100);
      for i in 1 .. cells.count loop
        pipe row (cells(i));
      end loop;
    end loop;
    
    free_context(ctx_id);
    
    return;
    
  end;*/

  /*
  function get_sheetRelId (
    p_workbook   in blob
  , p_sheetName  in varchar2
  )
  return varchar2
  is
    stream    Stream_T;
    sheetMap  SheetRelMap_T;
    relId     varchar2(255 char);
  begin
    stream := open_stream(p_workbook);
    read_SheetMap(stream, sheetMap);
    relId := sheetMap(p_sheetName);
    close_stream(stream);
    return relId;
  end;
  */

begin
  
  init();

end xutl_xlsb;
/
