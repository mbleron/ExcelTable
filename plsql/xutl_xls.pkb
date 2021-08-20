create or replace package body xutl_xls is

  RT_BOF           constant raw(2) := '0908';
  RT_FILEPASS      constant raw(2) := '2F00';
  RT_BOUNDSHEET8   constant raw(2) := '8500';
  RT_SST           constant raw(2) := 'FC00';
  RT_INDEX         constant raw(2) := '0B02';
  RT_EOF           constant raw(2) := '0A00';
  RT_DBCELL        constant raw(2) := 'D700';
  RT_CONTINUE      constant raw(2) := '3C00';
  RT_LABELSST      constant raw(2) := 'FD00';
  RT_RK            constant raw(2) := '7E02';
  RT_NUMBER        constant raw(2) := '0302';
  RT_BLANK         constant raw(2) := '0102';
  RT_MULRK         constant raw(2) := 'BD00';
  RT_MULBLANK      constant raw(2) := 'BE00';
  RT_FORMULA       constant raw(2) := '0600';
  RT_STRING        constant raw(2) := '0702';
  RT_OBJ           constant raw(2) := '5D00';
  RT_NOTE          constant raw(2) := '1C00';
  RT_MSODRAWING    constant raw(2) := 'EC00';
  RT_TXO           constant raw(2) := 'B601';
  RT_BOOLERR       constant raw(2) := '0502';
  
  RT_USREXCL       constant raw(2) := '9401';
  RT_FILELOCK      constant raw(2) := '9501';
  RT_INTERFACEHDR  constant raw(2) := 'E100';
  RT_RRDINFO       constant raw(2) := '9601';
  RT_RRDHEAD       constant raw(2) := '3801';
  
  FT_STRING        constant raw(1) := '00';
  FT_BOOLEAN       constant raw(1) := '01';
  FT_ERROR         constant raw(1) := '02';
  FT_BLANK         constant raw(1) := '03';
  
  FT_ERR_NULL      constant raw(1) := '00';
  FT_ERR_DIV_ZERO  constant raw(1) := '07';
  FT_ERR_VALUE     constant raw(1) := '0F';
  FT_ERR_REF       constant raw(1) := '17';
  FT_ERR_NAME      constant raw(1) := '1D';
  FT_ERR_NUM       constant raw(1) := '24';
  FT_ERR_NA        constant raw(1) := '2A';
  FT_ERR_GETDATA   constant raw(1) := '2B';
   
  BOOL_FALSE       constant raw(1) := '00';
  BOOL_TRUE        constant raw(1) := '01';
  
  OT_NOTE          constant binary_integer := 25;
  
  ST_UNISTR        constant pls_integer := 0;
  ST_SHORTUNISTR   constant pls_integer := 1;
  ST_RICHUNISTR    constant pls_integer := 2;
  ST_UNISTR_NOCCH  constant pls_integer := 3;
  
  ERR_NO_PASSWORD   constant varchar2(100) := 'The workbook is encrypted but no password was provided';
  ERR_EXPECTED_REC  constant varchar2(100) := 'Error at position %d, expecting a [%s] record';
  
  type RecordTypeMap_T is table of varchar2(128) index by binary_integer;
  recordTypeMap  RecordTypeMap_T;
  
  type BitMaskTable_T is varray(8) of raw(1);
  BITMASKTABLE    constant BitMaskTable_T := BitMaskTable_T('01','02','04','08','10','20','40','80');

  type ColumnMap_T is table of varchar2(2) index by pls_integer;
  
  type Range_T is record (
    firstRow  pls_integer
  , lastRow   pls_integer
  , colMap    ColumnMap_T
  );

  type Record_T is record (
    rt         raw(2)
  , current    integer
  , sz         pls_integer
  , available  pls_integer
  , has_next   boolean := true
  , next       integer
  );
  
  type Buffer_T is record (
    content    raw(32767)
  , offset     pls_integer := 0
  , sz         pls_integer
  , available  pls_integer := 0
  );

  type Stream_T is record (
    content    blob
  , sz         integer
  , offset     integer
  , rec        Record_T
  , buf        Buffer_T 
  );
  
  type String_T is record (
    is_lob       boolean := false
  , strValue     varchar2(32767)
  , lobValue     clob  
  );
  
  type XLUnicodeRichExtString_T is record (
    cch          binary_integer
  , fHighByte    boolean
  , fExtSt       boolean
  , fRichSt      boolean
  , cRun         binary_integer
  , cbExtRst     binary_integer
  , byte_len     pls_integer := 0
  , content      String_T
  );
  
  type String_Array_T is table of String_T;
  
  type SST_T is record (
    cstTotal     binary_integer
  , cstUnique    binary_integer
  , strings      String_Array_T
  );
  
  type BoundSheet8_T is record (
    lbPlyPos  integer
  , hsState   raw(1)
  , dt        raw(1)
  , stName    varchar2(255 char)
  );
  
  type Int16_Array_T is table of binary_integer index by pls_integer;
  
  type DBCell_T is record (
    pos     integer
  , dbRtrw  binary_integer
  , rgdb    Int16_Array_T
  , prev_pos integer
  );
  
  type DBCellArray_T is table of DBCell_T;
  
  type Index_T is record (
    rwMic   binary_integer
  , rwMac   binary_integer
  , rgibRw  DBCellArray_T
  );
  
  type IndexArray_T is table of Index_T;
  
  type Obj_T is record (
    ft  raw(2)
  , cb  raw(2)
  , ot  binary_integer
  , id  binary_integer
  );
  
  type TxOMap_T is table of varchar2(32767) index by binary_integer;
  
  type NoteSh_T is record (
    rw        binary_integer
  , col       binary_integer
  , idObj     binary_integer
  , stAuthor  varchar2(54 char)
  , txt       varchar2(32767)
  );
  
  type BoundSheetList_T is table of BoundSheet8_T;
  type SheetMap_T is table of pls_integer index by varchar2(512);
  
  type Workbook_T is record (
    sheets    BoundSheetList_T
  , sheetMap  SheetMap_T
  , sst       SST_T
  );
  
  type RK_T is record (
    fX100     boolean
  , fInt      boolean
  , RkNumber  raw(4)
  );
  
  type Formula_T is record (
    byte1   raw(1)
  , byte2   raw(1)
  , byte3   raw(1)
  , byte4   raw(1)
  , byte5   raw(1)
  , byte6   raw(1)
  , fExprO  raw(2)
  );
  
  -- comments
  type CommentMap_T is table of varchar2(32767) index by varchar2(7); -- cell comment indexed by cellref
  type Notes_T is table of CommentMap_T index by pls_integer; -- comment map indexed by sheet index
  
  type Context_T is record (
    stream      Stream_T
  , wb          Workbook_T
  , rng         Range_T
  , blockCount  pls_integer
  , blockNum    pls_integer
  , done        boolean := false
  , shIndices   IndexArray_T
  , currShIdx   pls_integer
  , readNotes   boolean
  , notes       Notes_T 
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
    recordTypeMap(raw2int(RT_MSODRAWING)) := 'MsoDrawing';
    recordTypeMap(raw2int(RT_TXO)) := 'TxO';
    recordTypeMap(raw2int(RT_CONTINUE)) := 'Continue';
    recordTypeMap(raw2int(RT_STRING)) := 'String';
    recordTypeMap(raw2int(RT_DBCELL)) := 'DBCell';
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
  , rt      in raw
  )
  is
  begin
    if stream.rec.rt != rt then
      debug(
        utl_lms.format_message(
          'Unexpected record found at position %d [0x%s] : 0x%s instead of 0x%s (%s)'
        , to_char(stream.rec.current)
        , to_char(stream.rec.current, 'FM0XXXXXXX')
        , rawtohex(stream.rec.rt)
        , rawtohex(rt)
        , recordTypeMap(raw2int(rt))
        )
      );
      error(-20731, ERR_EXPECTED_REC, stream.rec.current, recordTypeMap(raw2int(rt)));
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

  procedure bufferize (
    stream  in out nocopy Stream_T
  )
  is
  begin
    stream.offset := stream.offset + stream.buf.offset - 1;
    stream.buf.sz := 32767;
    dbms_lob.read(stream.content, stream.buf.sz, stream.offset, stream.buf.content);
    stream.buf.available := stream.buf.sz;
    stream.buf.offset := 1;
  end;

  function read_bytes (
    stream  in out nocopy Stream_T
  , amount  in pls_integer
  )
  return raw
  is
    bytes  raw(32767);
  begin
    
    if amount > 0 then
      
      if amount > stream.buf.available then
        bufferize(stream);
      end if;
      
      bytes := utl_raw.substr(stream.buf.content, stream.buf.offset, amount);
      stream.buf.offset := stream.buf.offset + amount;
      stream.buf.available := stream.buf.available - amount;
      stream.rec.available := stream.rec.available - amount;
    
    end if;
    
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
    stream.offset := 1;
    stream.buf.sz := 0;
    stream.rec.next := 1;
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
    HEADER_SIZE  constant pls_integer := 4;
  begin
    
    if stream.rec.next >= stream.offset + stream.buf.sz then
      stream.offset := stream.rec.next;
      --stream.buf.sz := 0;
      stream.buf.offset := 1;
      bufferize(stream);
      stream.rec.current := stream.offset;
    else
      stream.buf.offset := stream.rec.next - stream.offset + 1;
      stream.buf.available := stream.buf.sz - stream.buf.offset + 1;
      stream.rec.current := stream.rec.next;
    end if;
    
    -- record type
    stream.rec.rt := read_bytes(stream, 2);
    -- record size
    stream.rec.sz := read_int16(stream);
    
    stream.rec.available := stream.rec.sz;
    -- next record absolute offset
    stream.rec.next := stream.rec.current + stream.rec.sz + HEADER_SIZE;
    stream.rec.has_next := (stream.rec.next < stream.sz);
    
    --debug(utl_lms.format_message('REC[%s] 0x%s',rawtohex(stream.rec.rt),to_char(stream.rec.current,'FM0XXXXXXX')));
    
  end;

  procedure seek_first (
    stream       in out nocopy Stream_T
  , record_type  in raw
  )
  is
  begin
    while stream.rec.has_next and stream.rec.rt != record_type loop
      next_record(stream);
      exit when stream.rec.rt = RT_EOF;
    end loop;
  end;

  procedure seek (
    stream      in out nocopy Stream_T
  , pos         in integer
  , read_header in boolean default false
  )
  is
  begin
    
    if pos >= stream.offset + stream.buf.sz or pos < stream.offset then
      stream.offset := pos;
      stream.buf.sz := 0;
    else
      stream.buf.offset := pos - stream.offset + 1;
    end if;
    stream.rec.next := pos;
    stream.rec.has_next := (stream.rec.next < stream.sz);
    
    if read_header then
      next_record(stream);
    end if;
    
  end;
  
  procedure skip (
    stream  in out nocopy Stream_T
  , amount  in integer
  )
  is
    rem     integer;
    skipped integer;
  begin
    
    if stream.rec.available >= amount then
      
      if amount > stream.buf.available then
        bufferize(stream);
      end if;
      
      stream.buf.offset := stream.buf.offset + amount;
      stream.buf.available := stream.buf.available - amount;
      
      stream.rec.available := stream.rec.available - amount;
      
    else
      
      rem := amount - stream.rec.available;
      while rem != 0 loop
        next_record(stream);
        expect(stream, RT_CONTINUE);
        skipped := least(rem, stream.rec.available);
        
        stream.buf.offset := stream.buf.offset + skipped;
        stream.buf.available := stream.buf.available - skipped;
        
        stream.rec.available := stream.rec.available - skipped;
        rem := rem - skipped;
      end loop;
    
    end if;

  end;

  function base26encode (colNum in pls_integer)
  return varchar2
  result_cache
  is
  begin
    return chr(nullif(trunc(colNum/26),0)+64) || chr(mod(colNum,26)+65);
  end;

  function base26decode (colRef in varchar2)
  return pls_integer
  result_cache
  is
  begin
    return ascii(substr(colRef,-1,1))-65 + nvl((ascii(substr(colRef,-2,1))-64)*26, 0);
  end;

  function parseColumnList (
    cols  in varchar2
  , sep   in varchar2 default ','
  )
  return ColumnMap_T
  is
    colMap  Columnmap_t;
    i       pls_integer;
    token   varchar2(2);
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
    ctx_id  in pls_integer
  )
  is
    currShIdx       pls_integer := ctx_cache(ctx_id).currShIdx;
    shIndicesCount  pls_integer := ctx_cache(ctx_id).shIndices.count;
    has_next        boolean := (currShIdx < shIndicesCount);
  begin
    while has_next loop
      currShIdx := currShIdx + 1;
      debug('Switching to sheet '||currShIdx);
      ctx_cache(ctx_id).blockNum := 1;
      ctx_cache(ctx_id).blockCount := ctx_cache(ctx_id).shIndices(currShIdx).rgibRw.count;
      exit when ctx_cache(ctx_id).blockCount != 0;
      has_next := (currShIdx < shIndicesCount);
    end loop;
    if not has_next then
      debug('End of sheet list');
      ctx_cache(ctx_id).done := true;
    end if;
    ctx_cache(ctx_id).currShIdx := currShIdx;
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

  function read_LabelSst (
    ctx_id  in pls_integer
  , stream  in out nocopy Stream_T
  )
  return String_T
  is
    idx  pls_integer := read_int32(stream) + 1;
  begin
    return ctx_cache(ctx_id).wb.sst.strings(idx);
  end;

  procedure read_XLString (
    stream  in out nocopy Stream_T
  , str     in out nocopy XLUnicodeRichExtString_T
  , sttype  in pls_integer
  )
  is
    raw1    raw(1);
    buf     raw(8224);
    cbuf    varchar2(32767);  
    rem     pls_integer;
    amt     pls_integer;
    csz     pls_integer;
    csname  varchar2(30);
  begin
    if sttype = ST_SHORTUNISTR then
      str.cch := read_int8(stream);
    elsif sttype != ST_UNISTR_NOCCH then
      str.cch := read_int16(stream);
    end if;
    
    raw1 := read_bytes(stream, 1);
    
    if sttype = ST_RICHUNISTR then 
      str.fExtSt := is_bit_set(raw1, 3);
      str.fRichSt := is_bit_set(raw1, 4);
    
      if str.fRichSt then
        str.cRun := read_int16(stream);
      end if;
      
      if str.fExtSt then
        str.cbExtRst := read_int32(stream);
      end if;
    end if; 
    
    rem := str.cch; -- characters left to read;
    
    loop
      
      str.fHighByte := is_bit_set(raw1, 1);
      
      if str.fHighByte then
        csz := 2;
        csname := 'AL16UTF16LE';
      else
        csz := 1;
        csname := 'WE8ISO8859P1';
      end if;
      
      amt := least(stream.rec.available, rem * csz); -- byte amount to read      
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
            
      exit when rem = 0;
      next_record(stream);
      expect(stream, RT_CONTINUE);
      -- first byte specifies the character compression flag again
      raw1 := read_bytes(stream, 1);
      
    end loop;
    /*
    if stream.available = 0 then
      next_record(stream);
    end if;
    */
    if sttype = ST_RICHUNISTR then
      if str.fRichSt then
        skip(stream, str.cRun * 4); -- size of FormatRun struct = 4
      end if;
      
      if str.fExtSt then
        skip(stream, str.cbExtRst);
      end if;
    end if;
        
  end;
  
  function read_XLString (
    stream  in out nocopy Stream_T
  , sttype  in pls_integer
  )
  return String_T
  is
    xlstr  XLUnicodeRichExtString_T;
  begin
    read_XLString(stream, xlstr, sttype);
    return xlstr.content;
  end;
  
  procedure read_Formula (
    stream in out nocopy Stream_T
  , str    out nocopy String_T
  , num    out nocopy number
  , ftype  out nocopy raw
  )
  is
    frml  Formula_T;
  begin
    frml.byte1 := read_bytes(stream, 1);
    frml.byte2 := read_bytes(stream, 1);
    frml.byte3 := read_bytes(stream, 1);
    frml.byte4 := read_bytes(stream, 1);
    frml.byte5 := read_bytes(stream, 1);
    frml.byte6 := read_bytes(stream, 1);
    frml.fExprO := read_bytes(stream, 2);
    
    if frml.fExprO = hextoraw('FFFF') then
      
      case frml.byte1
      when FT_STRING then
        next_record(stream);
        
        -- optionally, next record might be an Array, Table or ShrFmla record
        if stream.rec.rt != RT_STRING then
          next_record(stream);
        end if;
        
        expect(stream, RT_STRING);
        str := read_XLString(stream, ST_UNISTR);
        
      when FT_BOOLEAN then
        str.strValue := case frml.byte3 
                          when BOOL_TRUE then 'TRUE' 
                          when BOOL_FALSE then 'FALSE' 
                        end;
      when FT_ERROR then
        str.strValue := 
          case frml.byte3
            when FT_ERR_NULL then '#NULL!'
            when FT_ERR_DIV_ZERO then '#DIV/0!'
            when FT_ERR_VALUE then '#VALUE!'
            when FT_ERR_REF then '#REF!'
            when FT_ERR_NAME then '#NAME?'
            when FT_ERR_NUM then '#NUM!'
            when FT_ERR_NA then '#N/A'
          end; 
      when FT_BLANK then
        str.strValue := '<BLANK>';
      end case;
      
      ftype := FT_STRING;
      
    else
      
      num := to_number(
               utl_raw.cast_to_binary_double(
                 utl_raw.concat(frml.byte1,frml.byte2,frml.byte3,frml.byte4,frml.byte5,frml.byte6,frml.fExprO)
               , utl_raw.little_endian
               )
             );
             
      ftype := null;
      
    end if;
    
  end;

  procedure read_BoolErr (
    stream  in out nocopy Stream_T
  , str     in out nocopy String_T
  )
  is
    bBoolErr  raw(1) := read_bytes(stream, 1);
    fError    boolean := (read_bytes(stream, 1) = BOOL_TRUE);
  begin
    if fError then
      str.strValue := 
        case bBoolErr
          when FT_ERR_NULL then '#NULL!'
          when FT_ERR_DIV_ZERO then '#DIV/0!'
          when FT_ERR_VALUE then '#VALUE!'
          when FT_ERR_REF then '#REF!'
          when FT_ERR_NAME then '#NAME?'
          when FT_ERR_NUM then '#NUM!'
          when FT_ERR_NA then '#N/A'
          when FT_ERR_GETDATA then '#GETTING_DATA'
        end;
    else
      str.strValue := 
        case bBoolErr 
          when BOOL_TRUE then 'TRUE' 
          when BOOL_FALSE then 'FALSE' 
        end;
    end if;
  end;
  
  function read_SheetInfo (
    stream  in out nocopy Stream_T
  )
  return BoundSheet8_T
  is
    sheetInfo  BoundSheet8_T;
  begin
    sheetInfo.lbPlyPos := read_int32(stream);
    sheetInfo.hsState := read_bytes(stream, 1);
    sheetInfo.dt := read_bytes(stream, 1);
    sheetInfo.stName := read_XLString(stream, ST_SHORTUNISTR).strValue;
    return sheetInfo;
  end;

  procedure read_DBCell (
    stream  in out nocopy Stream_T
  , db      in out nocopy DBCell_T
  )
  is
    rgdbSize  pls_integer;
  begin
    
    seek(stream, db.pos, read_header => true);
    
    -- issue #15 - check bad DBCell pointer
    if stream.rec.rt != RT_DBCELL then
      debug(utl_lms.format_message('Bad DBCell pointer (0x%s)',to_char(db.pos,'FM0XXXXXXX')));
      -- go to previous DBCell record
      seek(stream, db.prev_pos, read_header => true);
      -- skip it
      next_record(stream);
      -- go to next DBCell record
      seek_first(stream, RT_DBCELL);
      -- save correct pointer
      db.pos := stream.rec.current;
    end if;
    
    db.dbRtrw := read_int32(stream);
    rgdbSize := (stream.rec.sz - 4)/2; -- array size = (record size - 4)/2
    for i in 1 .. rgdbSize loop
      db.rgdb(i) := read_int16(stream);
    end loop;
  end;

  procedure read_CellBlock (
    ctx_id  in pls_integer
  , stream  in out nocopy Stream_T
  , rng     in Range_T
  , cells   in out nocopy ExcelTableCellList
  )
  is
    db     DBCell_T;
    base   integer; -- end of 1st row
    cb     integer; -- start of cell block
    rw     binary_integer;
    col    binary_integer;
    num    number;
    str    String_T;
    cnt    pls_integer;
    ftype  raw(1);

    currShIdx   pls_integer := ctx_cache(ctx_id).currShIdx;
    blockNum    pls_integer := ctx_cache(ctx_id).blockNum;
    blockCount  pls_integer := ctx_cache(ctx_id).blockCount;
    
    function get_note (cell_ref in varchar2) return varchar2 is
    begin
      if ctx_cache(ctx_id).notes.exists(currShIdx) and ctx_cache(ctx_id).notes(currShIdx).exists(cell_ref) then
        return ctx_cache(ctx_id).notes(currShIdx)(cell_ref);
      else
        return null;
      end if;
    end;
    
    procedure add_cell(v in anydata) is
    begin
      if rng.colMap.exists(col) then
        cells.extend;
        cells(cells.last) := new ExcelTableCell(rw + 1, rng.colMap(col), '', v, currShIdx, get_note(rng.colMap(col)||to_char(rw+1)));
      end if;
    end;
    
  begin
    
    db := ctx_cache(ctx_id).shIndices(ctx_cache(ctx_id).currShIdx).rgibRw(blockNum);
    
    if db.dbRtrw != 0 then
      
      base := db.pos - db.dbRtrw + 20;
      cb := base + db.rgdb(1);     
      seek(stream, cb);
      -- read cells until start of DBCell record
      loop
        next_record(stream);
        exit when stream.rec.rt = RT_DBCELL;
        rw := read_int16(stream);
        
        exit when rw > rng.lastRow;
        continue when rw < rng.firstRow;
        
        case stream.rec.rt
        when RT_MULRK then
          
          col := read_int16(stream);
          cnt := (stream.rec.sz - 6)/6;
          for i in 1 .. cnt loop
            skip(stream, 2); -- ixfe
            num := read_RK(stream);
            add_cell(anydata.ConvertNumber(num));
            --debug(utl_lms.format_message('[%d,%d]=%s', rw+1, col+1, rawtohex(stream.rec.rt)));
            col := col + 1;
          end loop;
          
        when RT_MULBLANK then
          
          col := read_int16(stream);
          cnt := (stream.rec.sz - 6)/2;
          for i in 1 .. cnt loop
            skip(stream, 2); -- ixfe
            --debug(utl_lms.format_message('[%d,%d]=%s', rw+1, col+1, rawtohex(stream.rec.rt)));
            col := col + 1;
          end loop;      
        
        else
          col := read_int16(stream);
          skip(stream, 2); -- ixfe
          num := null;
          str := null;
          case stream.rec.rt
          when RT_RK then
            num := read_RK(stream);
            add_cell(anydata.ConvertNumber(num));
          when RT_NUMBER then
            num := read_Number(stream);
            add_cell(anydata.ConvertNumber(num));
          when RT_BLANK then
            num := null;
          when RT_LABELSST then
            str := read_LabelSst(ctx_id, stream);

            if str.is_lob then
              add_cell(anydata.ConvertClob(str.lobValue));           
            else
              add_cell(anydata.ConvertVarchar2(str.strValue));
            end if;
          when RT_FORMULA then
            read_Formula(stream, str, num, ftype);
            if ftype = FT_STRING then
              if str.is_lob then
                add_cell(anydata.ConvertClob(str.lobValue));           
              else
                add_cell(anydata.ConvertVarchar2(str.strValue));
              end if;
            else
              add_cell(anydata.ConvertNumber(num));
            end if;
          when RT_BOOLERR then
            read_BoolErr(stream, str);
            add_cell(anydata.ConvertVarchar2(str.strValue));
          else
            null;
          end case;
          
          --debug(utl_lms.format_message('[%d,%d]=%s', rw+1, col+1, rawtohex(stream.rec.rt)));
        
        end case;
        
      end loop;
    
    end if;

    blockNum := blockNum + 1;
    if blockNum > blockCount then
      next_sheet(ctx_id);
    else  
      ctx_cache(ctx_id).blockNum := blockNum;
    end if;
    
  end;

  procedure read_SST (
    stream  in out nocopy Stream_T
  , sst     in out nocopy SST_T
  )
  is
  begin
    
    sst.cstTotal := read_int32(stream);
    sst.cstUnique := read_int32(stream);
    
    debug('sst.cstTotal = '||sst.cstTotal);
    debug('sst.cstUnique = '||sst.cstUnique);
    
    sst.strings := String_Array_T();
    sst.strings.extend(sst.cstUnique);
    
    for i in 1 .. sst.cstUnique loop
      sst.strings(i) := read_XLString(stream, ST_RICHUNISTR);
      
      if stream.rec.available = 0 then
        next_record(stream);
        -- workaround if cstUnique is not set correctly :
        -- force exiting the loop when something else than a Continue record is encountered
        exit when stream.rec.rt != RT_CONTINUE;
      end if;
      
    end loop;
  
  end;
  
  procedure read_Index (
    ctx       in out nocopy Context_T
  , sheetName in varchar2         
  )
  is
    i              pls_integer;
    cnt            pls_integer;
    pos            integer;
    firstRowBlock  pls_integer;
    lastRowBlock   pls_integer;
    idx            Index_T;
    idx_pos        integer;
  begin
    
    -- move file pointer to BOF of sheet
    i := ctx.wb.sheetMap(sheetName);
    pos := ctx.wb.sheets(i).lbPlyPos + 1;
    seek(ctx.stream, pos);
    
    -- move to Index record
    seek_first(ctx.stream, RT_INDEX);
    -- save offset of Index record
    idx_pos := ctx.stream.rec.current;
  
    skip(ctx.stream, 4); -- reserved
    idx.rwMic := read_int32(ctx.stream);
    idx.rwMac := read_int32(ctx.stream);
    skip(ctx.stream, 4); -- ibXf (ignored)
    
    firstRowBlock := trunc((greatest(ctx.rng.firstRow, idx.rwMic) - idx.rwMic)/32);
    lastRowBlock := trunc((least(ctx.rng.lastRow, idx.rwMac - 1) - idx.rwMic)/32);
     
    -- number of DBCell pointers
    cnt := lastRowBlock - firstRowBlock + 1;
    
    -- skip till first row block pointer
    skip(ctx.stream, firstRowBlock * 4);
    idx.rgibRw := DBCellArray_T(); 
    
    for i in 1 .. cnt loop
      pos := read_int32(ctx.stream) + 1;
      idx.rgibRw.extend;
      idx.rgibRw(i).pos := pos;
      --debug(utl_lms.format_message('Row block %s (0x%s)',to_char(i,'FM0999'),to_char(pos,'FM0XXXXXXX')));
    end loop;
    
    pos := idx_pos; -- initialize with index record offset
    for i in 1 .. cnt loop
      idx.rgibRw(i).prev_pos := pos;
      read_DBCell(ctx.stream, idx.rgibRw(i));
      -- in case a bad DBCell pointer was found
      pos := idx.rgibRw(i).pos;
    end loop;
    
    ctx.shIndices.extend;
    ctx.shIndices(ctx.shIndices.last) := idx;
    
  end;

  function read_TxO (
    stream  in out nocopy Stream_T
  )
  return varchar2
  is
    str      XLUnicodeRichExtString_T;
  begin
    
    skip(stream, 10);
    -- cchText
    str.cch := read_int16(stream);
    -- read text data in the next set of Continue records
    if str.cch != 0 then
      next_record(stream);
      read_XLString(stream, str, ST_UNISTR_NOCCH);
    end if;
    
    return str.content.strValue;
    
  end;

  procedure read_Obj (
    stream  in out nocopy Stream_T
  , txtMap  in out nocopy TxOMap_T
  )
  is
    obj  Obj_T;
  begin
    obj.ft := read_bytes(stream, 2);
    obj.cb := read_bytes(stream, 2);
    obj.ot := read_int16(stream);
    obj.id := read_int16(stream);
    
    debug('Object type = '||obj.ot);
    -- process only Note (comment) object
    if obj.ot = OT_NOTE then
      
      next_record(stream);
      -- next record should be an MsoDrawing
      expect(stream, RT_MSODRAWING);      
      -- TODO : test if the MsoDrawing contains an OfficeArtClientTextbox structure (2.5.195)      
      next_record(stream);
      -- next record should be a TxO
      expect(stream, RT_TXO);
      
      txtMap(obj.id) := read_TxO(stream);
    
    end if;
  end;
  
  function read_NoteSh (
    stream  in out nocopy Stream_T
  )
  return NoteSh_T
  is
    note  NoteSh_T;
  begin
    note.rw := read_int16(stream);
    note.col := read_int16(stream);
    skip(stream, 2);
    note.idObj := read_int16(stream);
    note.stAuthor := read_XLString(stream, ST_UNISTR).strValue;
    return note;
  end;

  function read_Comments (
    ctx        in out nocopy Context_T
  , sheetName  in varchar2
  )
  return CommentMap_T
  is
    i             pls_integer;
    pos           integer;
    found_NoteSh  boolean := false;
    txtMap        TxOMap_T;
    note          NoteSh_T;
    commentMap    CommentMap_T;
  begin
    
    -- move file pointer to BOF of sheet
    debug('sheetName = '||sheetName);
    i := ctx.wb.sheetMap(sheetName);
    pos := ctx.wb.sheets(i).lbPlyPos + 1;
    seek(ctx.stream, pos);
    
    -- move to first Obj record, or EOF if no Obj found
    seek_first(ctx.stream, RT_OBJ);
    
    while ctx.stream.rec.rt != RT_EOF loop
      case ctx.stream.rec.rt
      when RT_OBJ then
        read_Obj(ctx.stream, txtMap);
      when RT_NOTE then
        found_NoteSh := true;
        note := read_NoteSh(ctx.stream);
        commentMap(base26encode(note.col) || to_char(note.rw + 1)) := txtMap(note.idObj);
      else
        exit when found_NoteSh;
      end case;
      next_record(ctx.stream);
    end loop;
    
    return commentMap;

  end;

  procedure decrypt (
    stream    in out nocopy Stream_T
  , password  in varchar2
  )
  is
    DUMMY_HEADER constant raw(4) := hextoraw('00000000');
    BLOCKSIZE    pls_integer;
        
    keyInfo              xutl_offcrypto.rc4_info_t;
    derivedKey           raw(16);
    blockNum             pls_integer := 0;
    blockBuffer          raw(1024);
    decryptedBlock       raw(1024);
    blockOffset          pls_integer;
    
    leftInBlock          pls_integer;
    numBlocks            pls_integer;

    chnkStart            pls_integer;
    chnkLen              pls_integer;   
    rangeStart           pls_integer;
    rangeEnd             pls_integer;
    
    targetOffset         integer := 0;
    dummyHeaderSplitPos  pls_integer;
    
    type chunk_t is record (offset pls_integer, len pls_integer);
    type chunk_list_t is table of chunk_t;
    chunks  chunk_list_t := chunk_list_t();

    procedure add_chunk (p_offset in pls_integer, p_len in pls_integer)
    is
      i  pls_integer;
    begin
      chunks.extend;
      i := chunks.last;
      chunks(i).offset := p_offset;
      chunks(i).len := p_len;
    end;
    
    procedure init_block (num in pls_integer)
    is
    begin
      blockBuffer := null;
      blockOffset := 1;
      leftInBlock := 1024;
      derivedKey := xutl_offcrypto.get_key_binary_rc4(keyInfo, num);
    end;

    procedure write_block_chunk
    is
      rangeSize  pls_integer := rangeEnd - rangeStart + 1;
    begin
      dbms_lob.write(stream.content, rangeSize, targetOffset + rangeStart, utl_raw.substr(decryptedBlock, rangeStart, rangeSize));
    end;
    
    procedure write_block
    is
    begin
      debug('write_block');
      decryptedBlock := dbms_crypto.Decrypt(blockBuffer, dbms_crypto.ENCRYPT_RC4, derivedKey);
      blockSize := utl_raw.length(decryptedBlock);
      -- start overwriting original stream with decrypted block
      rangeStart := 1;
      
      for i in 1 .. chunks.count loop  
      
        chnkStart := chunks(i).offset;
        chnkLen := chunks(i).len;
        
        if chnkStart > rangeStart then
          rangeEnd := chnkStart - 1;
          write_block_chunk;
        end if;
        
        rangeStart := chnkStart + chnkLen;
      
      end loop;
      
      chunks.delete;
      
      if rangeStart > BLOCKSIZE + 1 then
        add_chunk(1, chnkLen - (BLOCKSIZE - chnkStart + 1));
      elsif rangeStart <= BLOCKSIZE then
        rangeEnd := BLOCKSIZE;
        write_block_chunk;
      end if;
          
      targetOffset := targetOffset + BLOCKSIZE;
      
      blockNum := blockNum + 1;
      init_block(blockNum);
    end;
    
    procedure fill_block (chunk in raw)
    is
      chunkSize  pls_integer := nvl(utl_raw.length(chunk),0);
    begin
      blockBuffer := utl_raw.concat(blockBuffer, chunk);
      blockOffset := blockOffset + chunkSize;
      leftInBlock := leftInBlock - chunkSize;
      if leftInBlock = 0 then
        write_block;
      end if;
    end;
    
  begin
    
    debug('start decrypt');
    keyInfo := xutl_offcrypto.get_binary_rc4_info(read_bytes(stream, stream.rec.sz), password);
    -- rewind
    seek(stream, 1);
  
    init_block(blockNum);
    
    while stream.rec.has_next loop
      
      next_record(stream);
    
      -- [MS-XLS] 2.2.10
      -- When obfuscating or encrypting BIFF records in these streams the record type 
      -- and record size components MUST NOT be obfuscated or encrypted. 
      -- In addition the following records MUST NOT be obfuscated or encrypted: 
      -- BOF (section 2.4.21), FilePass (section 2.4.117), UsrExcl (section 2.4.339), 
      -- FileLock (section 2.4.116), InterfaceHdr (section 2.4.146), RRDInfo (section 2.4.227), 
      -- and RRDHead (section 2.4.226). 
      -- Additionally, the lbPlyPos field of the BoundSheet8 record (section 2.4.28) MUST NOT be encrypted.
      case 
      when stream.rec.rt in (
        RT_BOF
      , RT_FILEPASS
      , RT_USREXCL
      , RT_FILELOCK
      , RT_INTERFACEHDR
      , RT_RRDINFO
      , RT_RRDHEAD
      )
      then
        -- whole record
        add_chunk(blockOffset, stream.rec.sz + 4);
      when stream.rec.rt = RT_BOUNDSHEET8 then
        -- record header + lbPlyPos (4 bytes)
        add_chunk(blockOffset, 8);
      else
        -- record header only
        add_chunk(blockOffset, 4);
      end case;
      
      dummyHeaderSplitPos := least(leftInBlock, 4);
      
      fill_block(utl_raw.substr(DUMMY_HEADER, 1, dummyHeaderSplitPos));
      
      if dummyHeaderSplitPos < 4 then
        fill_block(utl_raw.substr(DUMMY_HEADER, dummyHeaderSplitPos+1));
      end if;
    
      if stream.rec.sz <= leftInBlock then
        -- read all record data and append to current block
        fill_block(read_bytes(stream, stream.rec.sz));
        
      else      
        -- read record data that fits in the current block
        fill_block(read_bytes(stream, leftInBlock));       
        -- read record data in chunk of 1024 bytes
        numBlocks := trunc(stream.rec.available/1024);
        for i in 1 .. numBlocks loop
          fill_block(read_bytes(stream, 1024));
        end loop;
        -- read rest of available record data
        if stream.rec.available != 0 then
          fill_block(read_bytes(stream, stream.rec.available));         
        end if;
      
      end if;
      
    end loop;
    
    write_block;
    
    debug('end decrypt');
  
  end;
  
  procedure read_Globals (
    stream    in out nocopy Stream_T
  , wb        in out nocopy Workbook_T
  , password  in varchar2
  )
  is
    i       pls_integer;
    pos     integer;
  begin

    wb.sheets := BoundSheetList_T();
    
    next_record(stream);
    while stream.rec.rt != RT_EOF loop
      
      case stream.rec.rt
      when RT_FILEPASS then
        
        debug('The workbook is encrypted');
        if password is null then
          error(-20730, ERR_NO_PASSWORD);
        else
          -- save pointer to next record
          pos := stream.rec.next;
          decrypt(stream, password);
          -- restore stream
          seek(stream, pos);
        end if;
      
      when RT_BOUNDSHEET8 then
        
        wb.sheets.extend;
        i := wb.sheets.last;
        wb.sheets(i) := read_sheetInfo(stream);
        wb.sheetMap(wb.sheets(i).stName) := i;
      
      when RT_SST then
        
        read_SST(stream, wb.sst);
      
      else
        
        null;
      
      end case;
    
      next_record(stream);
      
    end loop;
    
    for i in 1 .. wb.sheets.count loop
      debug(to_char(wb.sheets(i).lbPlyPos,'FM0XXXXXXX')||' '||wb.sheets(i).stName);
    end loop; 
    
  end;
  
  function new_context (
    p_file      in blob 
  , p_password  in varchar2 default null
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null
  , p_readNotes in boolean default true
  )
  return pls_integer
  is
    ctx     Context_T;
    ctx_id  pls_integer; 
  begin
    ctx.rng.firstRow := nvl(p_firstRow, 1) - 1;
    ctx.rng.lastRow := nvl(p_lastRow, 65536) - 1;
    ctx.rng.colMap := parseColumnList(p_cols);
    
    ctx.stream := open_stream(p_file);
    read_Globals(ctx.stream, ctx.wb, p_password);
    
    ctx.shIndices := IndexArray_T();
    ctx.currShIdx := 0;
    ctx.readNotes := nvl(p_readNotes, true);
    ctx_id := nvl(ctx_cache.last, 0) + 1;
    ctx_cache(ctx_id) := ctx;
    
    return ctx_id;   
  end;
  
  procedure free_context (
    p_ctx_id  in pls_integer 
  )
  is
  begin
    close_stream(ctx_cache(p_ctx_id).stream);
    ctx_cache(p_ctx_id).wb.sst.strings.delete;
    ctx_cache.delete(p_ctx_id);
  end;

  function get_sheetList (
    p_ctx_id  in pls_integer 
  )
  return ExcelTableSheetList
  is
    sheetList  ExcelTableSheetList := ExcelTableSheetList();
    sheets     BoundSheetList_T := ctx_cache(p_ctx_id).wb.sheets;
  begin
    for i in 1 .. sheets.count loop
      sheetList.extend;
      sheetList(i) := sheets(i).stName;
    end loop;
    return sheetList;
  end;

  procedure add_sheets (
    p_ctx_id     in pls_integer
  , p_sheetList  in ExcelTableSheetList
  )
  is
  begin
    for i in 1 .. p_sheetList.count loop
      read_Index(ctx_cache(p_ctx_id), p_sheetList(i));
      if ctx_cache(p_ctx_id).readNotes then
        ctx_cache(p_ctx_id).notes(i) := read_Comments(ctx_cache(p_ctx_id), p_sheetList(i));
      end if;
    end loop;
    next_sheet(p_ctx_id);
  end;

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer default null
  )
  return ExcelTableCellList
  is
    numBlocks  pls_integer := ceil(nvl(p_nrows,1024)/32);
    cells      ExcelTableCellList := ExcelTableCellList();
  begin
    for i in 1 .. numBlocks loop
      if not ctx_cache(p_ctx_id).done then
        read_CellBlock(p_ctx_id, ctx_cache(p_ctx_id).stream, ctx_cache(p_ctx_id).rng, cells);
      end if;
    end loop;
    return cells;
  end;
  
  /*
  function get_comments (
    p_ctx_id     in pls_integer
  , p_sheetName  in varchar2
  )
  return ExcelTableCellList
  is
    comments  ExcelTableCellList := ExcelTableCellList();
  begin
    --read_Comments(ctx_cache(p_ctx_id), p_sheetName, comments);
    return comments;
  end;
  */
  
  function getRows (
    p_file      in blob 
  --, p_sheet     in varchar2
  , p_password  in varchar2 default null
  , p_cols      in varchar2 default null
  , p_firstRow  in pls_integer default null
  , p_lastRow   in pls_integer default null
  )
  return ExcelTableCellList
  pipelined
  is
    ctx_id  pls_integer;
    cells   ExcelTableCellList;   
  begin
    
    ctx_id := new_context(p_file, /*p_sheet, */ p_password, p_cols, p_firstRow, p_lastRow);

    while not ctx_cache(ctx_id).done loop
      cells := ExcelTableCellList();
      read_CellBlock(ctx_id, ctx_cache(ctx_id).stream, ctx_cache(ctx_id).rng, cells);
      for i in 1 .. cells.count loop
        pipe row (cells(i));
      end loop;
    end loop;
    
    free_context(ctx_id);
    
    return;
    
  end;
  
begin
  
  init();

end xutl_xls;
/
