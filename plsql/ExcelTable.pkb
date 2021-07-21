create or replace package body ExcelTable is

  -- Parser Constants
  T_MINUS                constant binary_integer := 16; -- -
  T_LEFT                 constant binary_integer := 28; -- (
  T_RIGHT                constant binary_integer := 29; -- )
  T_COMMA                constant binary_integer := 32; -- ,
  T_STAR                 constant binary_integer := 35; -- *
  T_EOF                  constant binary_integer := -1; -- end-of-file
  T_INT                  constant binary_integer := 40;
  T_NAME                 constant binary_integer := 41;
  T_STRING               constant binary_integer := 42;
  T_IDENT                constant binary_integer := 43;
  T_COLON                constant binary_integer := 46; -- :
  
  -- parser flags
  PARSE_METADATA         constant binary_integer := 1;
  PARSE_COLUMN           constant binary_integer := 2;
  PARSE_POSITION         constant binary_integer := 4;
  PARSE_SIMPLE           constant binary_integer := 8;
  PARSE_DEFAULT          constant binary_integer := PARSE_COLUMN + PARSE_METADATA;
  
  QUOTED_IDENTIFIER      constant binary_integer := 1;
  DIGITS                 constant varchar2(10) := '0123456789';
  LETTERS                constant varchar2(26) := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  
  META_VALUE             constant binary_integer := 1;
  --META_FORMULA           constant binary_integer := 4;
  META_CONSTANT          constant binary_integer := 32;
  
  INVALID_CHARACTER      constant varchar2(100) := 'Invalid character ''%s'' (%d) found at position %d';
  UNEXPECTED_EOF         constant varchar2(100) := 'Unexpected end-of-file';
  UNEXPECTED_SYMBOL      constant varchar2(100) := 'Error at position %d : unexpected symbol ''%s''';
  UNEXPECTED_INSTEAD_OF  constant varchar2(100) := 'Error at position %d : unexpected symbol ''%s'' instead of ''%s''';
  UNSUPPORTED_DATATYPE   constant varchar2(100) := 'Unknown or unsupported data type : %s';
  RANGE_EMPTY_REF        constant varchar2(100) := 'Range error : empty reference';
  RANGE_INVALID_REF      constant varchar2(100) := 'Range error : invalid reference ''%s''';
  RANGE_INVALID_COL      constant varchar2(100) := 'Range error : column out of range ''%s''';
  RANGE_INVALID_ROW      constant varchar2(100) := 'Range error : row out of range ''%d''';
  RANGE_INVALID_EXPR     constant varchar2(100) := 'Range error : invalid range expression ''%s''';
  RANGE_START_ROW_ERR    constant varchar2(100) := 'Range error : start row (%d) must be lower or equal than end row (%d)';
  RANGE_START_COL_ERR    constant varchar2(100) := 'Range error : start column (''%s'') must be lower or equal than end column (''%s'')';
  RANGE_EMPTY_COL_REF    constant varchar2(100) := 'Range error : missing column reference in ''%s''';
  RANGE_EMPTY_ROW_REF    constant varchar2(100) := 'Range error : missing row reference in ''%s'''; 
  SINGLETON_CLAUSE       constant varchar2(100) := 'At most one ''%s'' clause is allowed';
  MIXED_COLUMN_DEF       constant varchar2(100) := 'Cannot mix positional and named column definitions';
  MIXED_FIELD_DEF        constant varchar2(100) := 'Cannot mix positional and non-positional field definitions';
  EMPTY_COL_REF          constant varchar2(100) := 'Missing column reference for ''%s''';
  INVALID_COL_REF        constant varchar2(100) := 'Invalid column reference ''%s''';
  INVALID_COL            constant varchar2(100) := 'Column out of range ''%s''';
  DUPLICATE_COL_REF      constant varchar2(100) := 'Duplicate column reference ''%s''';
  DUPLICATE_COL_NAME     constant varchar2(100) := 'Duplicate column name ''%s''';
  NO_PASSWORD            constant varchar2(100) := 'The document is encrypted but no password was provided';
  DML_UNKNOWN_TYPE       constant varchar2(100) := 'Unknown DML statement type';
  DML_NO_KEY             constant varchar2(100) := 'No key column specified';
  INVALID_FILTER_TYPE    constant varchar2(100) := 'Invalid sheet filter type ''%s''';
  INVALID_DOCUMENT       constant varchar2(100) := 'Input file does not appear to be a valid Office document';
  INVALID_READ_METHOD    constant varchar2(100) := 'Invalid read method : %d';
  --UNIMPLEMENTED_FEAT     constant varchar2(100) := 'Unimplemented feature : %s';

  FILE_XLSX              constant pls_integer := 0;
  FILE_XLSB              constant pls_integer := 1;
  FILE_XLS               constant pls_integer := 2;
  FILE_ODS               constant pls_integer := 3;
  FILE_FF                constant pls_integer := 4;
  FILE_XSS               constant pls_integer := 5;
  
  -- OOX Constants
  -- moved to prop_map to implement OOXML strict:
  /*
  RS_OFFICEDOC           constant varchar2(100) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  RS_WORKSHEET           constant varchar2(100) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
  RS_COMMENTS            constant varchar2(100) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
  */
  CT_XL_BINARY_FILE      constant varchar2(100) := 'application/vnd.ms-excel.sheet.binary.macroEnabled.main';
  CT_SHAREDSTRINGS_BIN   constant varchar2(100) := 'application/vnd.ms-excel.sharedStrings';
  CT_SHAREDSTRINGS       constant varchar2(100) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';

  -- moved to prop_map:
  --SML_NSMAP              constant varchar2(100) := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  
  -- ODF constants
  MIMETYPE_ODS           constant varchar2(100) := 'application/vnd.oasis.opendocument.spreadsheet';
  ODF_OFFICE_NSMAP       constant varchar2(100) := 'xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"';
  ODF_TABLE_NSMAP        constant varchar2(100) := 'xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0"';
  ODF_TEXT_NSMAP         constant varchar2(100) := 'xmlns:x="urn:oasis:names:tc:opendocument:xmlns:text:1.0"';
  ODF_OFFICE_TABLE_NSMAP constant varchar2(150) := ODF_OFFICE_NSMAP || ', ' || ODF_TABLE_NSMAP;
  ODF_OFFICE_TEXT_NSMAP  constant varchar2(150) := ODF_OFFICE_NSMAP || ', ' || ODF_TEXT_NSMAP;
  -- XMLSS
  XSS_DFLT_NSMAP         constant varchar2(100) := 'xmlns="urn:schemas-microsoft-com:office:spreadsheet"';
  XSS_SS_NSMAP           constant varchar2(100) := 'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"';
  
  TMP_TABLE_QNAME        constant varchar2(261) := dbms_assert.enquote_name(sys_context('userenv','current_schema'),false)||'."EXCELTABLE$TMP"';
  QITD_BINARY_HEADER     constant raw(3) := hextoraw('A939FF');
  QITD_BASE64_HEADER     constant varchar2(4) := utl_raw.cast_to_varchar2(utl_encode.base64_encode(QITD_BINARY_HEADER));
    
  -- DB Constants
  DB_CSID                constant pls_integer := nls_charset_id('CHAR_CS');
  DB_CHARSET             constant varchar2(30) := nls_charset_name(DB_CSID);
  DB_VERSION             varchar2(15);
  MAX_CHAR_SIZE          pls_integer;
  LOB_CHUNK_SIZE         pls_integer;
  MAX_STRING_SIZE        pls_integer;
  VC2_MAXSIZE            pls_integer;
  MAX_IDENT_LENGTH       pls_integer;

  MAX_COLUMN_NUMBER      pls_integer;
  MAX_ROW_NUMBER         pls_integer;

  value_out_of_range     exception;
  pragma exception_init (value_out_of_range, -1438);

  buffer_too_small       exception;
  pragma exception_init (buffer_too_small, -24331);

  -- Internal structure definitions
  type metadata_t is record (
    typecode       pls_integer
  , prec           pls_integer
  , scale          pls_integer
  , len            pls_integer
  , csid           pls_integer
  , csfrm          pls_integer
  , attr_elt_type  anytype
  , aname          varchar2(128)
  , schema_name    varchar2(128)
  , type_name      varchar2(128)
  , version        varchar2(30)
  , numelems       pls_integer
  -- extra fields
  , len_in_char    pls_integer
  , col_ref        varchar2(3)
  , max_value      integer
  );
  
  type QI_cell_ref_t is record (c varchar2(3), cn pls_integer, r pls_integer); 
  type QI_range_t is record (start_ref QI_cell_ref_t, end_ref QI_cell_ref_t);
  type QI_position_t is record (start_offset pls_integer, end_offset pls_integer);
  
  type QI_column_t is record (
    metadata       metadata_t
  , format         varchar2(30)
  , for_ordinality boolean default false
  , cell_meta      binary_integer
  , is_key         boolean default false
  , has_default    boolean default false
  , default_value  anydata default null
  , is_positional  boolean default false
  , position       QI_position_t
  );
  
  type QI_column_list_t is table of QI_column_t;
  type QI_column_set_t is table of pls_integer index by varchar2(128);
  type QI_column_ref_set_t is table of binary_integer index by varchar2(3);
  
  type QI_definition_t is record (
    range          QI_range_t
  , cols           QI_column_list_t
  , colSet         QI_column_set_t
  , refSet         QI_column_ref_set_t
  , hasOrdinal     boolean default false
  , hasComment     boolean default false
  --, hasFormula     boolean default false
  , hasSheetName   boolean default false
  , hasSheetIndex  boolean default false
  , isPositional   boolean default false
  );

  type token_map_t is table of varchar2(30) index by binary_integer;
  type token_t is record (type binary_integer, strval varchar2(4000), intval binary_integer, pos binary_integer);
  type tokenizer_t is record (expr varchar2(4000), pos binary_integer, options binary_integer);

  type t_entry is record (offset integer, csize integer, ucsize integer, crc32 raw(4));
  type t_entries is table of t_entry index by varchar2(260);
  type t_archive is record (entries t_entries, content blob);
  
  -- open xml structures
  type t_sheetEntry is record (idx pls_integer, path varchar2(260));
  type t_sheet is record (idx pls_integer, name varchar2(128), path varchar2(260), content blob, comments blob);
  type t_sheetMap is table of t_sheetEntry index by varchar2(128);
  type t_sheets is table of t_sheet;
  
  type t_workbook is record (
    path            varchar2(260)
  , content         xmltype
  , content_binary  blob
  , rels            xmltype
  , sheetmap        t_sheetMap
  );
  
  type t_exceldoc is record (
    content_map xmltype
  , wb          t_workbook
  , is_xlsb     boolean
  , is_strict   boolean default false
  );
  
  -- Transitional and strict OOXML properties
  type t_prop is record (
    trans_value  varchar2(256)
  , strict_value varchar2(256)
  );
  type t_prop_map is table of t_prop index by varchar2(256);
  
  -- string cache
  type t_string_rec is record (strval varchar2(32767), lobval clob);
  type t_strings is table of t_string_rec;
  
  -- comments
  type t_commentMap is table of varchar2(4000) index by varchar2(10);
  type t_comments is table of t_commentMap index by pls_integer;
  
  -- target table info
  type t_table_info is record (schema_name varchar2(128), table_name varchar2(128), dblink varchar2(128));

  type t_cell_info is record (
    cellRef   varchar2(10)
  , cellRow   pls_integer
  , cellCol   varchar2(3)
  , cellType  varchar2(10)
  , cellValue varchar2(32767)
  , sheetIdx  pls_integer
  );
  
  type t_dom_reader is record (
    doc         dbms_xmldom.DOMDocument
  , xpath       varchar2(128)
  , rlist       dbms_xmldom.DOMNodeList
  , rlist_size  pls_integer
  , rlist_idx   pls_integer
  , ns_map      varchar2(256)
  );
  
  type t_xdb_reader is record (table_name varchar2(128), c integer, cell_info t_cell_info);
  
  -- local context cache
  type t_context is record (
    file_type    pls_integer
  , read_method  pls_integer
  , def_cache    QI_definition_t
  , string_cache t_strings
  , comments     t_comments
  , done         boolean default false
  , r_num        binary_integer
  , src_row      pls_integer
  , row_repeat   pls_integer
  , tmp_row      ExcelTableCellList
  , dom_reader   t_dom_reader
  , xdb_reader   t_xdb_reader
  , extern_key   integer
  , ws_content   blob
  , table_info   t_table_info
  , sheets       t_sheets
  , curr_sheet   pls_integer
  , is_strict    boolean
  );
  
  type t_ctx_cache is table of t_context index by binary_integer;
  
  prop_map               t_prop_map;
  token_map              token_map_t;
  tokenizer              tokenizer_t;
  ctx_cache              t_ctx_cache;
  nls_date_format        varchar2(64);
  nls_timestamp_format   varchar2(64);
  nls_numeric_char       varchar2(2);
  fetch_size             binary_integer := 100;
  tmp_table_exists       boolean := false;
  debug_status           boolean := false;
  
  -- if set to true, p_sheet argument is interpreted as a regex pattern
  sheet_pattern_enabled  boolean := false;
  
  procedure error (
    p_message in varchar2
  , p_arg1 in varchar2 default null
  , p_arg2 in varchar2 default null
  , p_arg3 in varchar2 default null
  , p_code in number default -20722
  ) 
  is
  begin
    raise_application_error(p_code, utl_lms.format_message(p_message, p_arg1, p_arg2, p_arg3));
  end;

  
  procedure setDebug (p_status in boolean)
  is
  begin
    debug_status := p_status;
  end;


  procedure debug (message in varchar2)
  is
  begin
    if debug_status then
      dbms_output.put_line('[exceltable] '||message);
    end if;
  end;

  /*
  procedure debug2 (message in varchar2)
  is
    pragma autonomous_transaction;
  begin
    insert into my_log (ts, txt) values (systimestamp, message);
    commit;
  end;
  */
  

  /* =============================================================================================
   Retrieve the maximum number of byte(s) per character for the database character set.
   This piece of information can be extracted from the charset name in Oracle naming convention : 
   WE8ISO8859P15 --> 8 bits --> 1 byte
   AL32UTF8 --> 32 bits --> 4 bytes
   Only exceptions are UTF8 and UTFE (3 and 4 bytes respectively)
   
   As of 12.1, this can be done directly via utl_i18n.get_max_character_size() function.
  ============================================================================================= */ 
  function get_max_char_size (p_charset in varchar2)
  return pls_integer
  is
  begin
    return $IF DBMS_DB_VERSION.VER_LE_11
           $THEN case p_charset
                   when 'UTF8' then 3
                   when 'UTFE' then 4
                   else ceil(to_number(regexp_substr(p_charset, '\d+'))/8)
                 end;
           $ELSE utl_i18n.get_max_character_size(p_charset);
           $END    
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


  function get_max_ident_length (p_compatible in varchar2)
  return pls_integer 
  is
    l_result   pls_integer;
    l_version  pls_integer := to_number(regexp_substr(p_compatible, '\d+', 1, 1));
    l_release  pls_integer := to_number(regexp_substr(p_compatible, '\d+', 1, 2));
  begin
    if l_version > 12 or (l_version = 12 and l_release >= 2) then
      l_result := 128;
    else
      l_result := 30;
    end if; 
    return l_result;
  end;  


  procedure set_sheet_limits (
    column_number  in pls_integer
  , rw_number      in pls_integer
  )
  is
  begin
    MAX_COLUMN_NUMBER := column_number;
    MAX_ROW_NUMBER := rw_number;
  end;
  

  procedure init_state is
    l_compatibility  DB_VERSION%type;
  begin
    dbms_utility.db_version(DB_VERSION, l_compatibility);
    MAX_CHAR_SIZE := get_max_char_size(DB_CHARSET);
    LOB_CHUNK_SIZE := trunc(32767 / MAX_CHAR_SIZE);
    MAX_STRING_SIZE := get_max_string_size();
    VC2_MAXSIZE := trunc(MAX_STRING_SIZE / MAX_CHAR_SIZE);
    MAX_IDENT_LENGTH := get_max_ident_length(l_compatibility);
    set_sheet_limits(16384, 1048576);
  
    token_map(T_NAME)   := '<name>';
    token_map(T_INT)    := '<integer>';
    token_map(T_IDENT)  := '<identifier>';
    token_map(T_STRING) := '<string literal>';
    token_map(T_EOF)    := '<eof>';
    token_map(T_COMMA)  := ',';
    token_map(T_LEFT)   := '(';
    token_map(T_RIGHT)  := ')';
    token_map(T_COLON)  := ':';
    
    prop_map('RS_OFFICEDOC').trans_value := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
    prop_map('RS_OFFICEDOC').strict_value := 'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument';
    prop_map('RS_WORKSHEET').trans_value := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
    prop_map('RS_WORKSHEET').strict_value := 'http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet';
    prop_map('RS_COMMENTS').trans_value := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
    prop_map('RS_COMMENTS').strict_value := 'http://purl.oclc.org/ooxml/officeDocument/relationships/comments';
    prop_map('DEFAULT_NS').trans_value := 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
    prop_map('DEFAULT_NS').strict_value := 'http://purl.oclc.org/ooxml/spreadsheetml/main';
    prop_map('SML_NSMAP').trans_value := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
    prop_map('SML_NSMAP').strict_value := 'xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"';
    
  end;
  
  
  function get_prop (
    prop_name in varchar2
  , is_strict in boolean
  )
  return varchar2
  is
  begin
    if is_strict then
      return prop_map(prop_name).strict_value;
    else
      return prop_map(prop_name).trans_value;
    end if;
  end;

  
  function is_opt_set (options in binary_integer, opt in binary_integer)
  return boolean
  is
  begin
    return ( bitand(options, opt) = opt );
  end;
  

  -- convert a base26-encoded number to decimal
  function base26decode (p_str in varchar2) 
  return pls_integer 
  result_cache
  is
    l_result  pls_integer;
    l_base    pls_integer := 1;
  begin
    if p_str is not null then
      l_result := 0;
      for i in 1 .. length(p_str) loop
        l_result := l_result + (ascii(substr(p_str,-i,1)) - 64) * l_base;
        l_base := l_base * 26;
      end loop;
    end if;
    return l_result;
  end;
  

  -- convert a number to base26 string
  function base26encode (p_num in pls_integer) 
  return varchar2
  result_cache
  is
    l_result  varchar2(3);
    l_num     pls_integer := p_num;
  begin
    if p_num is not null then
      while l_num != 0 loop
        l_result := chr(65 + mod(l_num-1,26)) || l_result;
        l_num := trunc((l_num-1)/26);
      end loop;
    end if;
    return l_result;
  end;
  
  
  function get_column_list (
    def_cache in QI_definition_t
  )
  return varchar2
  is
    column_list  varchar2(4000);
    col_ref      varchar2(3);
  begin
    col_ref := def_cache.refSet.first;
    while col_ref is not null loop
      column_list := column_list || ',' || col_ref;
      col_ref := def_cache.refSet.next(col_ref);
    end loop;
    return substr(column_list, 2);
  end;


  function get_position_list (
    cols in QI_column_list_t
  )
  return varchar2
  is
    position_list  varchar2(32767);
    position_sep   varchar2(30) := token_map(T_COLON);
    
    function make_range(i in pls_integer) return varchar2 is
    begin
      return cols(i).position.start_offset || position_sep || cols(i).position.end_offset;
    end;
    
  begin
    for i in 1 .. cols.count loop
      if not cols(i).for_ordinality then
        if position_list is not null then
          position_list := position_list || ',';
        end if;
        position_list := position_list || make_range(i);
      end if;
    end loop;
    return position_list;
  end;
  

  function parse_column_list (
    cols  in varchar2
  , sep   in varchar2 default ','
  )
  return QI_column_ref_set_t
  is
    rs      QI_column_ref_set_t;
    token   varchar2(30);
    p1      pls_integer := 1;
    p2      pls_integer;
    p3      pls_integer;
    
    c1      pls_integer;
    c2      pls_integer;
    
    function validate (item in varchar2, pos in pls_integer) return varchar2 is
    begin
      if item is null then
        error('Missing column reference at position %d', pos);
      elsif not regexp_like(item, '^[A-Z]+$') then
        error('Invalid column reference ''%s'' at position %d', item, pos);
      end if;
      return item;
    end;
    
  begin

    loop
      
      p2 := instr(cols, sep, p1);
      if p2 = 0 then
        token := substr(cols, p1);
      else
        token := substr(cols, p1, p2-p1);
      end if;
      
      -- is range?
      p3 := instr(token, '-');
      if p3 != 0 then
        c1 := base26decode(validate(substr(token, 1, p3-1), p1));
        c2 := base26decode(validate(substr(token, p3+1), p1+p3));
        if c2 < c1 then
          error('Invalid range expression at position %d', p1);
        end if;
        for i in c1 .. c2 loop
          rs(base26encode(i)) := META_VALUE + META_COMMENT;
        end loop;
      else
        token := validate(token, p1);
        rs(token) := META_VALUE + META_COMMENT;
      end if;
      
      exit when p2 = 0;
      p1 := p2 + 1;
      
    end loop;
    
    return rs;
  
  end;
  

  function get_free_ctx
  return binary_integer 
  is
  begin
    return nvl(ctx_cache.last, 0) + 1;
  end;
  
  
  function get_compiled_ctx (input in varchar2)
  return pls_integer
  is
    bytes  raw(32767) := utl_encode.base64_decode(utl_raw.cast_to_raw(input));
  begin
    return utl_raw.cast_to_binary_integer(utl_raw.substr(bytes, 4, 4));
  end;


  function is_compiled_ctx (input in varchar2)
  return boolean
  is
  begin
    return ( substr(input, 1, 4) = QITD_BASE64_HEADER );
  end;
  

  function parse_tdef (input in varchar2)
  return QI_definition_t
  is
    bytes   raw(32767) := utl_encode.base64_decode(utl_raw.cast_to_raw(input));
    sz      pls_integer := utl_raw.length(bytes);
    pos     pls_integer := 1;
    col     QI_column_t;
    tdef    QI_definition_t;

    function has_next return boolean is
    begin
      return (pos <= sz);
    end;
    function read (amount in pls_integer) return raw is
      output  raw(128);
    begin
      output := utl_raw.substr(bytes, pos, amount);
      pos := pos + amount;
      return output; 
    end;
    procedure skip (amount in pls_integer) is
    begin
      pos := pos + amount;
    end;
    function read_uint8 return pls_integer is
    begin
      return utl_raw.cast_to_binary_integer(read(1));
    end;
    function read_int8 return pls_integer is
      output  pls_integer := read_uint8;
    begin
      return case when output < 128 then output else output - 256 end;
    end;
    function read_uint16 return pls_integer is
    begin
      return utl_raw.cast_to_binary_integer(read(2));
    end;
    function read_string return varchar2 is
      len     pls_integer := read_uint8;
      output  varchar2(128);
    begin
      if len != 0 then
        output := utl_raw.cast_to_varchar2(read(len));
      end if;
      return output;
    end;

  begin
    
    tdef.cols := QI_column_list_t();
    skip(7); -- header + ctx_id
  
    while has_next() loop
      col := null;
      col.metadata.aname := read_string();
      col.metadata.typecode := read_uint8();
      case col.metadata.typecode
      when dbms_types.TYPECODE_VARCHAR2 then
        col.metadata.len := read_uint16();
      when dbms_types.TYPECODE_NUMBER then
        col.metadata.prec := nullif(read_uint8(),0);
        col.metadata.scale := nullif(read_int8(),-128);
      when dbms_types.TYPECODE_DATE then
        col.format := read_string();
      when dbms_types.TYPECODE_TIMESTAMP then
        col.metadata.scale := read_uint8();
        col.format := read_string();
      when dbms_types.TYPECODE_CLOB then
        col.metadata.csid := DB_CSID;
        col.metadata.csfrm := 1;
      else
        null;
      end case;
      
      tdef.cols.extend;
      tdef.cols(tdef.cols.last) := col;
      
    end loop;
    
    return tdef;
  
  end;


  function compile_tdef (
    ctx_id  in pls_integer
  )
  return varchar2
  is
    tdef      QI_definition_t := ctx_cache(ctx_id).def_cache;
    output    raw(32767);
    metadata  metadata_t;      
    procedure push (bytes in raw) is
    begin
      output := utl_raw.concat(output, bytes);
    end;
    procedure push_int8 (num in pls_integer) is
    begin
      push(utl_raw.substr(utl_raw.cast_from_binary_integer(num),-1));
    end;
    procedure push_int16 (num in pls_integer) is
    begin
      push(utl_raw.substr(utl_raw.cast_from_binary_integer(num),-2));
    end;
    procedure push_int32 (num in pls_integer) is
    begin
      push(utl_raw.cast_from_binary_integer(num));
    end;
    procedure push_string (str in varchar2) is
      binstr  raw(128) := utl_raw.cast_to_raw(str);
    begin
      push_int8(nvl(utl_raw.length(binstr),0));
      push(binstr);
    end;
  begin
    push(QITD_BINARY_HEADER);
    push_int32(ctx_id);
    for i in 1 .. tdef.cols.count loop
      metadata := tdef.cols(i).metadata;
      push_string(metadata.aname);
      push_int8(metadata.typecode);
      case metadata.typecode
      when dbms_types.TYPECODE_VARCHAR2 then
        push_int16(metadata.len);
      when dbms_types.TYPECODE_NUMBER then
        push_int8(nvl(metadata.prec,0));
        push_int8(nvl(metadata.scale,-128));
      when dbms_types.TYPECODE_DATE then
        push_string(tdef.cols(i).format);
      when dbms_types.TYPECODE_TIMESTAMP then
        push_int8(metadata.scale);
        push_string(tdef.cols(i).format);
      else
        null;
      end case;
    end loop;
    return utl_raw.cast_to_varchar2(utl_encode.base64_encode(output));
  end;
  
  
  procedure create_tmp_table
  is
    pragma autonomous_transaction;
    table_exists  exception;
    pragma exception_init(table_exists, -955);
    stmt   varchar2(2000) := 'CREATE GLOBAL TEMPORARY TABLE $$TAB (ID INTEGER, DATA XMLTYPE, XID1 INTEGER) ON COMMIT PRESERVE ROWS XMLTYPE DATA STORE AS BINARY XML (CACHE)';
  begin
    if not tmp_table_exists then
      stmt := replace(stmt, '$$TAB', dbms_assert.qualified_sql_name(TMP_TABLE_QNAME));
      execute immediate stmt;
      tmp_table_exists := true;
    end if;
  exception 
    when table_exists then
      tmp_table_exists := true;
  end;
  
  
  procedure load_tmp_table (
    ctx_id       in pls_integer
  , xml_content  in xmltype
  , xid1         in pls_integer default null
  )
  is
    stmt  varchar2(256) := 'INSERT INTO '||TMP_TABLE_QNAME||' (ID,DATA,XID1) VALUES (:1,:2,:3)';
  begin
    execute immediate stmt using ctx_id, xml_content, xid1;
  end;


  procedure delete_tmp_table (ctx_id in pls_integer)
  is
    stmt  varchar2(256) := 'DELETE '||TMP_TABLE_QNAME||' WHERE ID = :1';
  begin
    execute immediate stmt using ctx_id;
  end;
    

  function is_opc_package (p_file in blob)
  return boolean
  is
  begin
    return ( dbms_lob.substr(p_file, 4) = hextoraw('504B0304') );
  end;


  function is_xmlss (p_file in blob)
  return boolean
  is
    PI_START_PATTERN constant raw(18) := utl_raw.cast_to_raw('<?mso-application ');
    PI_STOP_PATTERN  constant raw(2) := utl_raw.cast_to_raw('?>');
    pi_start    integer;
    pi_stop     integer;
    pi_content  varchar2(128);
    output      boolean := false;
  begin
    pi_start := dbms_lob.instr(p_file, PI_START_PATTERN);
    if pi_start != 0 then
      pi_start := pi_start + utl_raw.length(PI_START_PATTERN); 
      pi_stop := dbms_lob.instr(p_file, PI_STOP_PATTERN, pi_start);
      pi_content := trim(utl_i18n.raw_to_char(dbms_lob.substr(p_file, pi_stop - pi_start, pi_start),'AL32UTF8'));
      output := ( pi_content = 'progid="Excel.Sheet"' );
    end if;
    return output;   
  end;
  

  function blob2xml (
    p_content  in blob
  )
  return xmltype
  is
  begin
    return xmltype(p_content, nls_charset_id('AL32UTF8'));
  end;

   
  -- ----------------------------------------------------------------------------------------------
  -- Open a zip archive and read entries from central directory segment
  -- References :
  -- https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
  -- https://en.wikipedia.org/wiki/Zip_%28file_format%29
  function Zip_openArchive (p_zip in blob)
  return t_archive
  is
  
    -- max offset of End of Central Directory Signature, from the end of the archive
    -- = ECD length (21) + max length of comment field (65535) - 1
    ECDS_MAX_OFFSET  constant binary_integer := 65555; 

    ecds             binary_integer; -- End of central directory signature
    oscd             binary_integer; -- Offset of start of central directory, relative to start of archive
    tncdr            binary_integer; -- Total number of central directory records
    fnl              binary_integer; -- File name length
    efl              binary_integer; -- Extra field length
    fcl              binary_integer; -- File comment length
    fn               varchar2(260);  -- File name
    gpb              raw(2);         -- General Purpose Bits
    enc              varchar2(30);
    cdrPtr           binary_integer := 0;
    entry            t_entry;
    my_archive       t_archive;

  begin

    ecds := dbms_lob.instr(p_zip, hextoraw('504B0506'), greatest(1, dbms_lob.getlength(p_zip) - ECDS_MAX_OFFSET));
    oscd := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 4, ecds+16), utl_raw.little_endian)+1;
    tncdr := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 2, ecds+10), utl_raw.little_endian);
    cdrPtr := oscd;
    
    for i in 1..tncdr loop

      gpb := dbms_lob.substr(p_zip, 2, cdrPtr+8);
      if utl_raw.bit_and(gpb, hextoraw('0008')) = hextoraw('0008') then
        enc := 'AL32UTF8';
      else
        enc := 'WE8PC850';
      end if;

      fnl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 2, cdrPtr+28), utl_raw.little_endian);
      efl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 2, cdrPtr+30), utl_raw.little_endian);
      fcl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 2, cdrPtr+32), utl_raw.little_endian);
      fn  := utl_i18n.raw_to_char(dbms_lob.substr(p_zip, fnl, cdrPtr+46), enc);

      if substr(fn, -1) != '/' then
        
        entry.offset := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 4, cdrPtr+42), utl_raw.little_endian) + 1;
        entry.crc32 := dbms_lob.substr(p_zip, 4, cdrPtr+16);
        entry.csize := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 4, cdrPtr+20), utl_raw.little_endian);
        entry.ucsize := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 4, cdrPtr+24), utl_raw.little_endian);
        my_archive.entries(fn) := entry;
        
      end if;
      -- next entry in central directory
      cdrPtr := cdrPtr + 46 + fnl + efl + fcl;

    end loop;

    my_archive.content := p_zip;

    return my_archive;

  end;

  function Zip_hasEntry (
    archive   in t_archive
  , entryName in varchar2 
  )
  return boolean
  is
  begin
    return archive.entries.exists(entryName);
  end;

  -- ----------------------------------------------------------------------------------------------
  -- Get a zip entry by its name
  -- MB 20180424 - added STORED method
  function Zip_getEntry (
    p_archive   in t_archive
  , p_entryname in varchar2
  )
  return blob
  is
    tmp      blob; 
    content  blob;
    cm       binary_integer; -- Compression method
    fnl      binary_integer; -- File name length
    efl      binary_integer; -- Extra field length
    lfh      binary_integer; -- Local file header
    entry    t_entry;
  begin
    
    if p_archive.entries.exists(p_entryname) then
       
      entry := p_archive.entries(p_entryname);
      lfh := entry.offset;
      cm := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_archive.content, 2, lfh+8), utl_raw.little_endian);
      fnl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_archive.content, 2, lfh+26), utl_raw.little_endian);
      efl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_archive.content, 2, lfh+28), utl_raw.little_endian);
      
      case cm
      when 8 then -- DEFLATE
        
        dbms_lob.createtemporary(tmp, true);
        -- gzip magic header + flags
        dbms_lob.writeappend(tmp, 10, hextoraw('1F8B08000000000000FF'));
        dbms_lob.copy(tmp, p_archive.content, entry.csize, 11, lfh + 30 + fnl + efl);
        dbms_lob.append(tmp, entry.crc32); -- CRC32
        dbms_lob.append(tmp, utl_raw.cast_from_binary_integer(entry.ucsize, utl_raw.little_endian)); -- uncompressed size
        
        dbms_lob.createtemporary(content, true);
        utl_compress.lz_uncompress(tmp, content);
        dbms_lob.freetemporary(tmp);
        
      when 0 then -- STORED
        
        dbms_lob.createtemporary(content, true);
        dbms_lob.copy(content, p_archive.content, entry.csize, 1, lfh + 30 + fnl + efl);
      
      else
        raise_application_error(-20724, utl_lms.format_message('Zip_getEntry: unsupported compression method (%d)', cm));
      end case;
      
    end if;
    
    return content;
    
  end;
  
  -- ----------------------------------------------------------------------------------------------
  -- Get a zip entry as XMLType
  --  assuming the part has been encoded in UTF-8, as Excel does natively
  function Zip_getXML (
    p_archive   in t_archive
  , p_partname  in varchar2
  )
  return xmltype
  is
  begin
    return blob2xml(Zip_getEntry(p_archive, p_partname));
  end;


  function get_nls_param (p_name in varchar2) 
  return varchar2 
  is
    l_result  nls_session_parameters.value%type;
  begin
    select value
    into l_result
    from nls_session_parameters
    where parameter = p_name ;
    return l_result;
  exception
    when no_data_found then
      return null;
  end;


  procedure set_nls_cache is
  begin
    nls_date_format := get_nls_param('NLS_DATE_FORMAT');
    nls_numeric_char := get_nls_param('NLS_NUMERIC_CHARACTERS');
    nls_timestamp_format := get_nls_param('NLS_TIMESTAMP_FORMAT');
  end;


  function get_date_format 
  return varchar2 
  is
  begin
    return nls_date_format;  
  end;

  function get_timestamp_format 
  return varchar2 
  is
  begin
    return nls_timestamp_format;  
  end;
 
  function get_decimal_sep 
  return varchar2 
  is
  begin
    return substr(nls_numeric_char, 1, 1);  
  end;


  function resolve_table (qualified_name in varchar2)
  return t_table_info
  is
    info             t_table_info;
    l_part2          varchar2(128);
    l_part1_type     pls_integer;
    l_object_number  number;
  begin
    dbms_utility.name_resolve(
      name          => qualified_name
    , context       => 0 -- table or view
    , schema        => info.schema_name
    , part1         => info.table_name
    , part2         => l_part2
    , dblink        => info.dblink
    , part1_type    => l_part1_type
    , object_number => l_object_number
    );
    return info;
  end; 


  function get_string (ctx_id in binary_integer, idx in binary_integer) 
  return t_string_rec 
  is
  begin
    return ctx_cache(ctx_id).string_cache(idx+1);
  end;

  
  function get_string_val (ctx_id in binary_integer, idx in binary_integer) 
  return varchar2 
  is
    rec  t_string_rec := ctx_cache(ctx_id).string_cache(idx+1);
  begin
    if rec.strval is not null then
      return rec.strval;
    else
      return dbms_lob.substr(rec.lobval, LOB_CHUNK_SIZE);
    end if;
  end;


  function get_clob_val (ctx_id in binary_integer, idx in binary_integer) 
  return clob is
    rec  t_string_rec := ctx_cache(ctx_id).string_cache(idx+1);
  begin
    if rec.strval is not null then
      return to_clob(rec.strval);
    else
      return rec.lobval;
    end if;
  end;


  function get_date_val (
    p_value  in number 
  )
  return date
  is
    l_date    date;
  begin
    -- Excel bug workaround : date 1900-02-29 doesn't exist yet Excel stores it at serial #60
    -- The following skips it and converts to Oracle date correctly
    if p_value > 60 then
      l_date := date '1899-12-30' + p_value;
    elsif p_value < 60 then
      l_date := date '1899-12-31' + p_value;
    end if;
    return l_date;
  end;


  function get_date_val (
    p_value   in varchar2
  , p_format  in varchar2
  )
  return date
  is
  begin
    return to_date(p_value, nvl(p_format, get_date_format));
  end;


  function get_tstamp_val (
    p_value  in number 
  )
  return timestamp
  is
    l_ts  timestamp;
  begin
    -- Excel bug workaround : date 1900-02-29 doesn't exist yet Excel stores it at serial #60
    -- The following skips it and converts to Oracle date correctly
    if p_value > 60 then
      l_ts := timestamp '1899-12-30 00:00:00' + numtodsinterval(p_value, 'DAY');
    elsif p_value < 60 then
      l_ts := timestamp '1899-12-31 00:00:00' + numtodsinterval(p_value, 'DAY');
    end if;
    return l_ts;
  end;


  function get_tstamp_val (
    p_value   in varchar2
  , p_format  in varchar2
  )
  return timestamp
  is
  begin
    return to_timestamp(p_value, nvl(p_format, get_timestamp_format));
  end;


  function get_tstamp_val_iso8601 (
    input  in varchar2
  , scale  in pls_integer default 3
  )
  return timestamp_unconstrained
  is
    DATETIME_FMT  constant varchar2(23) := 'YYYY-MM-DD"T"HH24:MI:SS';
    sepIndex      pls_integer := instr(input,'.',-1);
    output        timestamp_unconstrained;
  begin
    if sepIndex = 0 then
      output := to_timestamp(input, DATETIME_FMT);
    else
      output := to_timestamp(substr(input,1,sepIndex-1), DATETIME_FMT)
           + numtodsinterval(round(to_number(replace(substr(input, sepIndex),'.',get_decimal_sep)), scale), 'second');
    end if;
    return output; 
  end;


  function get_comment (
    ctx_id    in binary_integer
  , sheet_id  in pls_integer
  , cell_ref   in varchar2
  )
  return varchar2
  is
  begin
    if ctx_cache(ctx_id).comments.exists(sheet_id) and ctx_cache(ctx_id).comments(sheet_id).exists(cell_ref) then
      return ctx_cache(ctx_id).comments(sheet_id)(cell_ref);
    else
      return null;
    end if;
  end; 


  procedure readclob (
    p_node    in dbms_xmldom.DOMNode
  , p_content in out nocopy clob
  )
  is
    istream   utl_characterinputstream;
    tmp       varchar2(32767);
    buf       varchar2(32767);
    nread     integer := LOB_CHUNK_SIZE;
    reslen    integer := 0;
    residue   raw(3);
  begin
    
    if p_content is null then
      dbms_lob.createtemporary(p_content, true);
    end if;
    
    istream := dbms_xmldom.getNodeValueAsCharacterStream(p_node);
    
    loop
      
      istream.read(tmp, nread);
      exit when nread = 0;
      -- bug workaround - XMLCharacterInputStream reads bytes instead of characters
      -- (for details, see : https://community.oracle.com/thread/3934929)
      -- concatenate the residue to the current chunk :
      if reslen != 0 then
        tmp := utl_raw.cast_to_varchar2(utl_raw.concat(residue, utl_raw.cast_to_raw(tmp)));
      end if;
      -- get a character-complete string
      buf := substrc(tmp, 1);
      dbms_lob.writeappend(p_content, length(buf), buf);
      -- length of the residue
      reslen := lengthb(tmp) - lengthb(buf);
      if reslen != 0 then
        residue := utl_raw.substr(utl_raw.cast_to_raw(tmp), -reslen);
      end if;
      
    end loop;
    
    istream.close();
    
  end;


  procedure readclob (
    p_nlist   in dbms_xmldom.DOMNodeList
  , p_content in out nocopy clob
  )
  is
  begin    
    for i in 0 .. dbms_xmldom.getLength(p_nlist)-1 loop
      readclob(dbms_xmldom.item(p_nlist, i), p_content);
    end loop;
    dbms_xmldom.freeNodeList(p_nlist);
  end;


  function string_join (
    nlist  in dbms_xmldom.DOMNodeList
  , sep    in varchar2 default chr(10)
  )
  return t_string_rec
  is
  
    sep_len  constant pls_integer := length(sep);
    len      pls_integer;
    str      t_string_rec;
    tmp      varchar2(32767);
    node     dbms_xmldom.DOMNode;
    is_lob   boolean := false;
    
  begin
    
    for i in 0 .. dbms_xmldom.getLength(nlist) - 1 loop
      
      node := dbms_xmldom.item(nlist, i);
      
      begin
        
        dbms_xslprocessor.valueOf(node, '.', tmp);
        --tmp := dbms_xslprocessor.valueOf(node, '.');
          
        if is_lob then
          
          if i > 0 then
            dbms_lob.writeappend(str.lobval, sep_len, sep);
          end if;
          dbms_lob.writeappend(str.lobval, length(tmp), tmp);
          
        else
          
          len := lengthb(str.strval) + lengthb(tmp);
          if i > 0 then
            len := len + sep_len;
          end if;
            
          if len > 32767 then
            -- switch to CLOB storage 
            is_lob := true;
            dbms_lob.createtemporary(str.lobval, true);
            str.lobval := str.strval;
            -- line feed?
            if i > 0 then
              dbms_lob.writeappend(str.lobval, sep_len, sep);
            end if;
            dbms_lob.writeappend(str.lobval, length(tmp), tmp);
          else
            
            if i > 0 then
              str.strval := str.strval || sep;
            end if;
            str.strval := str.strval || tmp;
            
          end if;
            
        end if;    
            
      exception
        when value_error or buffer_too_small then
          
          if not is_lob then
            -- switch to CLOB storage 
            is_lob := true;
            dbms_lob.createtemporary(str.lobval, true);
            str.lobval := str.strval;
          end if;
          -- line feed?
          if i > 0 then
            dbms_lob.writeappend(str.lobval, sep_len, sep);
          end if;
          readclob(dbms_xslprocessor.selectNodes(node, './/text()'), str.lobval);
          
      end;
      
      dbms_xmldom.freeNode(node);
      
    end loop;
    
    return str;
    
  end;
  
  
  function next_token
  return token_t
  is
    token     token_t;
    pos       simple_integer := 0;
    c         varchar2(1 char);
    str       varchar2(4000);
          
    procedure set_token (
      p_type in binary_integer
    , p_strval in varchar2
    , p_pos in binary_integer default null
    ) 
    is
    begin
      token.type := p_type;
      token.strval := case when p_type = T_EOF then token_map(T_EOF) else p_strval end;
      token.pos := nvl(p_pos, tokenizer.pos);
      if p_type = T_INT then
        token.intval := to_number(p_strval);
      end if;
    end;
          
    function getc return varchar2 is
    begin
      tokenizer.pos := tokenizer.pos + 1;
      return substr(tokenizer.expr, tokenizer.pos, 1);
    end;
    
    function look_ahead return varchar2 is
    begin
      return substr(tokenizer.expr, tokenizer.pos + 1, 1); 
    end;
      
  begin
     
    c := getc;
    -- strip whitespaces
    while c in (chr(9), chr(10), chr(13), chr(32)) loop
      c := getc;
    end loop;    
      
    if c is null then
      set_token(T_EOF, c);
    else 
      case c
      when ',' then
        set_token(T_COMMA, c);          
      when '*' then
        set_token(T_STAR, c);
      when '(' then
        set_token(T_LEFT, c);
      when ')' then
        set_token(T_RIGHT, c);
      when ':' then
        set_token(T_COLON, c);
      -- string literal
      when '''' then 
        str := null;
        pos := tokenizer.pos;
        c := getc;
        while c is not null loop
          if c = '''' then 
            if look_ahead() = c then
              c := getc;
            else
              exit;
            end if;
          end if;
          str := str || c;
          c := getc;
        end loop;
        if c is null then
          error(UNEXPECTED_EOF);
        else
          set_token(T_STRING, str, pos);
        end if;

      when '"' then             
        str := null;
        pos := tokenizer.pos;
        c := getc;
        while c is not null loop
          exit when c = '"';
          str := str || c;
          c := getc;
        end loop;
        if c is null then
          error(UNEXPECTED_EOF);
        elsif tokenizer.options = QUOTED_IDENTIFIER then
          set_token(T_IDENT, str, pos);
        else
          set_token(T_NAME, str, pos);
        end if; 
                  
      when '-' then
        set_token(T_MINUS, c);
      
      else
        
        case  
        -- digits
        when c between '0' and '9' then
          str := c;
          pos := tokenizer.pos;
          c := getc;
          while c between '0' and '9' loop
            str := str || c;
            c := getc;
          end loop;
          set_token(T_INT, str, pos);         
          tokenizer.pos := tokenizer.pos - 1;
        -- string
        when c between 'A' and 'Z' 
          or c between 'a' and 'z'
        then       
          str := c;
          pos := tokenizer.pos;
          c := getc;
          while c between 'A' and 'Z'
             or c between 'a' and 'z'
             or c between '0' and '9'
             or c in ('_','$','#') 
          loop
            str := str || c;
            c := getc;
          end loop;                           
          set_token(T_NAME, str, pos);
          tokenizer.pos := tokenizer.pos - 1;
          
        else     
          error(INVALID_CHARACTER, c, ascii(c), tokenizer.pos);    
        end case;
        
      end case;
      
    end if;
    
    return token;
  
  end;


  function property_check (
    input_vector  in raw 
  , property      in raw
  )
  return boolean
  is
  begin
    return utl_raw.bit_and(input_vector, property) = property;
  end;
  

  procedure validate_column (
    p_col_ref in varchar2
  , p_name in varchar2
  , p_range in out nocopy QI_range_t
  )
  is
    l_coln  pls_integer;
  begin
    if p_col_ref is null then
      error(EMPTY_COL_REF, p_name);
    elsif rtrim(p_col_ref, LETTERS) is not null then
      error(INVALID_COL_REF, p_col_ref);
    end if;
    l_coln := base26decode(p_col_ref);
    if l_coln not between nvl(p_range.start_ref.cn, 1) and nvl(p_range.end_ref.cn, MAX_COLUMN_NUMBER) then
      error(INVALID_COL, p_col_ref);
    end if;   
  end;


  procedure validate_columns (
    tdef in out nocopy QI_definition_t
  )
  is
  
    type cell_meta_check_t is record (
      for_ordinality  boolean := false
    , sheet_index     boolean := false
    , sheet_name      boolean := false
    , cnt             pls_integer := 0
    );
  
    start_col         pls_integer;
    end_col           pls_integer;
    pos               pls_integer;
    col_cnt           pls_integer;
    col_ref_cnt       pls_integer := 0;
    col_ref           varchar2(3);
    col_name          varchar2(128);
    cell_meta         binary_integer;
    cell_meta_ref     binary_integer;
    meta_check        cell_meta_check_t;
    
    pos_ref_cnt       pls_integer := 0;
    
  begin
    
    col_cnt := tdef.cols.count;
    start_col := nvl(tdef.range.start_ref.cn, 1);
    end_col := nvl(tdef.range.end_ref.cn, col_cnt);
    pos := 0;
    
    for i in 1 .. col_cnt loop
      
      cell_meta := tdef.cols(i).cell_meta;
      
      if tdef.cols(i).for_ordinality then
        if meta_check.for_ordinality then
          error(SINGLETON_CLAUSE, 'FOR ORDINALITY');
        else
          meta_check.for_ordinality := true;
          meta_check.cnt := meta_check.cnt + 1;
        end if;
        
      elsif cell_meta = META_SHEET_NAME then
        if meta_check.sheet_name then
          error(SINGLETON_CLAUSE, 'FOR METADATA (SHEET_NAME)');
        else
          meta_check.sheet_name := true;
          meta_check.cnt := meta_check.cnt + 1;
        end if;

      elsif cell_meta = META_SHEET_INDEX then
        if meta_check.sheet_index then
          error(SINGLETON_CLAUSE, 'FOR METADATA (SHEET_INDEX)');
        else
          meta_check.sheet_index := true;
          meta_check.cnt := meta_check.cnt + 1;
        end if;
        
      elsif tdef.cols(i).metadata.col_ref is not null then
        col_ref_cnt := col_ref_cnt + 1;
        validate_column(tdef.cols(i).metadata.col_ref, tdef.cols(i).metadata.aname, tdef.range);
      
      elsif pos < end_col and cell_meta != META_CONSTANT then 
        tdef.cols(i).metadata.col_ref := base26encode(start_col + pos);
        pos := pos + 1;
        
      elsif cell_meta = META_CONSTANT then
        meta_check.cnt := meta_check.cnt + 1;
        
      else
        error(INVALID_COL, tdef.cols(i).metadata.aname);
      end if;
      
      if tdef.cols(i).is_positional then
        pos_ref_cnt := pos_ref_cnt + 1;
      end if;
      
      -- check for duplicate column names
      col_name := tdef.cols(i).metadata.aname;
      if tdef.colSet.exists(col_name) then
        error(DUPLICATE_COL_NAME, col_name);
      else
        tdef.colSet(col_name) := i;
      end if;
      
      -- check for duplicate column references
      if not ( tdef.cols(i).for_ordinality or cell_meta in (META_SHEET_NAME, META_SHEET_INDEX) ) then
        col_ref := tdef.cols(i).metadata.col_ref;
        if col_ref is not null then
          if tdef.refSet.exists(col_ref) then       
            cell_meta_ref := tdef.refSet(col_ref);                  
            --if bitand(cell_meta_ref, cell_meta) = 0  then
            if not is_opt_set(cell_meta_ref, cell_meta) then
              tdef.refSet(col_ref) := cell_meta_ref + cell_meta;
            else
              error(DUPLICATE_COL_REF, col_ref);     
            end if;        
          else
            tdef.refSet(col_ref) := cell_meta;
          end if;
        end if;
      end if;
    
    end loop;
    
    -- skip meta columns
    col_cnt := col_cnt - meta_check.cnt;
    
    -- check for mixed column definitions
    if col_ref_cnt != 0 and col_ref_cnt != col_cnt then
      error(MIXED_COLUMN_DEF);
    end if;
    
    -- check for mixed field definitions 
    if pos_ref_cnt != 0 then
      if pos_ref_cnt != col_cnt then
        error(MIXED_FIELD_DEF);
      else
        tdef.isPositional := true;
      end if;
    end if;
  
  end;


  function filterSheetList (
    documentSheetList  in ExcelTableSheetList
  , sheetFilter        in anydata
  )
  return t_sheets
  is
    sheetFilterType    varchar2(257) := sheetFilter.GetTypeName();
    sheetName          varchar2(128);
    sheetPattern       varchar2(128);
    inputSheetList     ExcelTableSheetList;
    dummy              pls_integer;
    i                  pls_integer := 0;
  
    filteredSheets     t_sheets := t_sheets();
    
  begin
    
    case 
    when sheetFilterType = 'SYS.VARCHAR2' then
      
      sheetPattern := sheetFilter.AccessVarchar2();
      
      for idx in 1 .. documentSheetList.count loop
        sheetName := documentSheetList(idx);
        if sheet_pattern_enabled and regexp_like(sheetName, sheetPattern)
           or not(sheet_pattern_enabled) and sheetName = sheetPattern 
        then
          filteredSheets.extend;
          i := i + 1;
          filteredSheets(i).idx := idx;
          filteredSheets(i).name := sheetName;
        end if;
      end loop; 
      
    when sheetFilterType like '%.EXCELTABLESHEETLIST' then
      
      dummy := sheetFilter.GetCollection(inputSheetList);
      
      for idx in 1 .. documentSheetList.count loop
        sheetName := documentSheetList(idx);
        if sheetName member of inputSheetList then
          filteredSheets.extend;
          i := i + 1;
          filteredSheets(i).idx := idx;
          filteredSheets(i).name := sheetName;
        end if;
      end loop;
      
    else
      raise_application_error(-20723, utl_lms.format_message(INVALID_FILTER_TYPE, sheetFilterType));
      
    end case;
   
    return filteredSheets;

  end;


  function QI_parseRange (p_expr in varchar2) 
  return QI_range_t
  is    
    l_pos    pls_integer;
    l_range  QI_range_t;
    
    procedure process_ref (p_expr in varchar2, p_ref in out nocopy QI_cell_ref_t) is
      l_col   varchar2(32);
      l_row   varchar2(32);
      l_coln  pls_integer;
      l_rnum  pls_integer;   
    begin
      if p_expr is null then
        error(RANGE_EMPTY_REF);
      end if;
      l_col := rtrim(p_expr, DIGITS);
      l_row := ltrim(p_expr, LETTERS);
      if rtrim(l_row, DIGITS) is not null or rtrim(l_col, LETTERS) is not null then
        error(RANGE_INVALID_REF, p_expr);
      end if;
      l_coln := base26decode(l_col);
      -- validate column reference
      if l_coln > MAX_COLUMN_NUMBER then
        error(RANGE_INVALID_COL, l_col);
      end if;
      l_rnum := to_number(l_row);
      if l_rnum not between 1 and MAX_ROW_NUMBER then
        error(RANGE_INVALID_ROW, l_rnum);
      end if;
      p_ref.r := l_rnum;
      p_ref.c := l_col; 
      p_ref.cn := l_coln;
    end;
    
  begin
    
    if p_expr is not null then
      
      l_pos := instr(p_expr, ':');
      if l_pos != 0 then
        process_ref(substr(p_expr, 1, l_pos-1), l_range.start_ref);
        process_ref(substr(p_expr, l_pos+1), l_range.end_ref);
        -- validate range :
        if l_range.start_ref.c is not null and l_range.end_ref.c is null 
          or l_range.start_ref.c is null and l_range.end_ref.c is not null 
          or l_range.start_ref.r is not null and l_range.end_ref.r is null 
          or l_range.start_ref.r is null and l_range.end_ref.r is not null
        then
          error(RANGE_INVALID_EXPR, p_expr);
        elsif l_range.start_ref.r > l_range.end_ref.r then
          error(RANGE_START_ROW_ERR, l_range.start_ref.r, l_range.end_ref.r);
        elsif base26decode(l_range.start_ref.c) > base26decode(l_range.end_ref.c) then
          error(RANGE_START_COL_ERR, l_range.start_ref.c, l_range.end_ref.c);
        end if;
                
      else
        process_ref(p_expr, l_range.start_ref);
        -- validate single cell reference
        if l_range.start_ref.c is null then
          error(RANGE_EMPTY_COL_REF, p_expr);
        elsif l_range.start_ref.r is null then
          error(RANGE_EMPTY_ROW_REF, p_expr);
        end if;
      end if;
    
    end if;
    
    return l_range;
        
  end;


  function QI_parseTable (
    p_range   in varchar2
  , p_cols    in varchar2
  , p_options in binary_integer default PARSE_DEFAULT
  ) 
  return QI_definition_t
  is
  
    length_too_long          exception;
    pragma exception_init (length_too_long, -910);
    identifier_too_long      exception;
    pragma exception_init (identifier_too_long, -972);
    zero_length_column       exception;
    pragma exception_init (zero_length_column, -1723);
    precision_out_of_range   exception;
    pragma exception_init (precision_out_of_range, -1727);
    scale_out_of_range       exception;
    pragma exception_init (scale_out_of_range, -1728);
    zero_length_identifier   exception;
    pragma exception_init (zero_length_identifier, -1741);
  
    token                token_t;
    tdef                 QI_definition_t;
    use_char_semantics   boolean := ( get_nls_param('NLS_LENGTH_SEMANTICS') = 'CHAR' );
    
    fl_parse_metadata    boolean := is_opt_set(p_options, PARSE_METADATA);
    fl_parse_column      boolean := is_opt_set(p_options, PARSE_COLUMN);
    fl_parse_position    boolean := is_opt_set(p_options, PARSE_POSITION);

    function accept (t in binary_integer, s in varchar2 default null, i in boolean default true) 
    return boolean 
    is
    begin
      if token.type = t 
         and ( s is null
               or token.strval = s 
               or (upper(token.strval) = upper(s) and i) ) 
      then
        token := next_token();
        return true;
      else
        return false;
      end if;
    end;
       
    procedure expect (t in binary_integer, s in varchar2 default null, i in boolean default true) is
    begin
      if not accept(t, s, i) then
        -- Error at position %d : unexpected symbol '%s' instead of '%s'
        error(UNEXPECTED_INSTEAD_OF, token.pos, token.strval, token_map(t));
      end if;
    end;
    
    -- column_expr    ::= identifier datatype
    procedure column_expr (p_columns in out nocopy QI_column_list_t)
    is
      intval  binary_integer;
      strval  varchar2(4000);
      pos     binary_integer;
      col     QI_column_t;
    begin
      strval := token.strval;
      expect(T_IDENT);
      if strval is null then
        raise zero_length_identifier;
      elsif lengthb(strval) > MAX_IDENT_LENGTH then
        raise identifier_too_long;
      else 
        col.metadata.aname := strval;
      end if;
      if accept(T_NAME, 'NUMBER') then
        col.metadata.typecode := dbms_types.TYPECODE_NUMBER;
        if accept(T_LEFT) then
          if accept(T_STAR) then
            null;
          else
            intval := token.intval;
            expect(T_INT);
            col.metadata.prec := intval;
            col.metadata.scale := 0;
          end if;
          if accept(T_COMMA) then
            if accept(T_MINUS) then
              intval := - token.intval;
            else
              intval := token.intval;
            end if;
            expect(T_INT);
            col.metadata.scale := intval;
          end if;      
          expect(T_RIGHT);
        end if;
        -- check precision and scale against allowed ranges
        if col.metadata.prec not between 1 and 38 then
          raise precision_out_of_range;
        elsif col.metadata.scale not between -84 and 127 then
          raise scale_out_of_range;
        end if;
        -- max value for constrained numbers
        col.metadata.max_value := 10**(col.metadata.prec - col.metadata.scale);
        
      elsif accept(T_NAME, 'VARCHAR2') then
        col.metadata.typecode := dbms_types.TYPECODE_VARCHAR2;
        col.metadata.csid := DB_CSID;
        col.metadata.csfrm := 1;
        expect(T_LEFT);
        intval := token.intval;
        expect(T_INT);
        if accept(T_NAME, 'BYTE') then
          use_char_semantics := false;
        elsif accept(T_NAME, 'CHAR') then
          use_char_semantics := true;
        end if;
        expect(T_RIGHT);
        -- check declared length against VARCHAR2 limit
        if intval = 0 then
          raise zero_length_column;
        elsif intval < 0 or intval > MAX_STRING_SIZE then
          raise length_too_long;
        elsif use_char_semantics then
          col.metadata.len_in_char := intval;
          intval := least(intval * MAX_CHAR_SIZE, MAX_STRING_SIZE);
        end if;
        col.metadata.len := intval;
      
      elsif accept(T_NAME, 'DATE') then
        col.metadata.typecode := dbms_types.TYPECODE_DATE;
        if accept(T_NAME, 'FORMAT') then
          strval := token.strval;
          expect(T_STRING);
          col.format := strval;
        end if;
        
      elsif accept(T_NAME, 'TIMESTAMP') then
        col.metadata.typecode := dbms_types.TYPECODE_TIMESTAMP;
        col.metadata.prec := null;
        if accept(T_LEFT) then
          intval := token.intval;
          expect(T_INT);
          col.metadata.scale := intval;
          expect(T_RIGHT);
        else
          -- default fractional seconds precision
          col.metadata.scale := 6;
        end if;        
        if accept(T_NAME, 'FORMAT') then
          strval := token.strval;
          expect(T_STRING);
          col.format := strval;
        end if;        
        
      elsif accept(T_NAME, 'CLOB') then
        col.metadata.typecode := dbms_types.TYPECODE_CLOB;
        col.metadata.csid := DB_CSID;
        col.metadata.csfrm := 1;
                
      elsif accept(T_NAME, 'FOR') then
        expect(T_NAME, 'ORDINALITY');
        col.metadata.typecode := dbms_types.TYPECODE_NUMBER;
        col.for_ordinality := true;
        tdef.hasOrdinal := true;
      else
        error(UNSUPPORTED_DATATYPE, token.strval);
      end if;
      
      pos := token.pos;
      strval := token.strval;
      -- column reference
      if accept(T_NAME, 'COLUMN') then
        if not(fl_parse_column) or col.for_ordinality then
          error(UNEXPECTED_SYMBOL, pos, strval);
        end if;
        strval := token.strval;
        expect(T_STRING);
        if strval is not null then
          col.metadata.col_ref := strval;
        else
          error(EMPTY_COL_REF, col.metadata.aname);
        end if;   
      --end if;
      
      -- position
      elsif accept(T_NAME, 'POSITION') then
        if not(fl_parse_position) or col.for_ordinality then
          error(UNEXPECTED_SYMBOL, pos, strval);
        end if;
        expect(T_LEFT);
        intval := token.intval;
        expect(T_INT);
        col.position.start_offset := intval;
        expect(T_COLON);
        intval := token.intval;
        expect(T_INT);
        col.position.end_offset := intval;
        expect(T_RIGHT);
        col.is_positional := true;
      end if;
      
      -- cell metadata
      if accept(T_NAME, 'FOR') then
        if not(fl_parse_metadata) or col.for_ordinality then
          error(UNEXPECTED_SYMBOL, pos, strval);
        end if;
        expect(T_NAME, 'METADATA');
        expect(T_LEFT);
        if accept(T_NAME, 'COMMENT') then
          col.cell_meta := META_COMMENT;
          tdef.hasComment := true;
        elsif accept(T_NAME, 'SHEET_NAME') then
          col.cell_meta := META_SHEET_NAME;
          tdef.hasSheetName := true;
        elsif accept(T_NAME, 'SHEET_INDEX') then
          col.cell_meta := META_SHEET_INDEX;
          tdef.hasSheetIndex := true;
        else
          error(UNEXPECTED_SYMBOL, token.pos, token.strval);
        end if;
        expect(T_RIGHT);
      else
        col.cell_meta := META_VALUE;
      end if;
      
      p_columns.extend;
      p_columns(p_columns.last) := col;
    
    end;

    -- column_list    ::= column_expr { "," column_expr }
    procedure column_list (p_columns in out nocopy QI_column_list_t)
    is
    begin
      column_expr(p_columns);
      while token.type = T_COMMA loop
        token := next_token();
        column_expr(p_columns);
      end loop;
    end;
    
    procedure table_expr 
    is
    begin
      tdef.cols := QI_column_list_t();
      column_list(tdef.cols);
    end;
    
  begin

    tdef.range := QI_parseRange(p_range);

    if p_options = PARSE_SIMPLE then
      tdef.refSet := parse_column_list(p_cols);
      -- force reading comments
      tdef.hasComment := true;
    else

      tokenizer.expr := p_cols;
      tokenizer.pos := 0;
      tokenizer.options := QUOTED_IDENTIFIER;

      token := next_token();
      table_expr;
      expect(T_EOF);
      
      validate_columns(tdef);
    
    end if;
    
    return tdef;
   
  end;


  function QI_initContext (
    p_range          in varchar2
  , p_cols           in varchar2
  , p_method         in binary_integer
  , p_parse_options  in binary_integer default PARSE_DEFAULT
  )
  return binary_integer
  is
    ctx_id  binary_integer := get_free_ctx();
  begin
    ctx_cache(ctx_id).read_method := p_method;
    ctx_cache(ctx_id).r_num := 0;
    ctx_cache(ctx_id).curr_sheet := 0;
    ctx_cache(ctx_id).def_cache := QI_parseTable(p_range, p_cols, p_parse_options);    

    return ctx_id;
  end;
    

  -- Java streaming methods wrappers
  function StAX_createContext(
    ctxType    in varchar2
  , cols       in varchar2
  , firstRow   in number
  , lastRow    in number
  , vc2MaxSize in number
  )
  return number
  as language java 
  name 'db.office.spreadsheet.ReadContext.initialize(java.lang.String, java.lang.String, int, int, int) return int';


  function StAX_getSheetList(key in number)
  return ExcelTableSheetList
  as language java
  name 'db.office.spreadsheet.ReadContext.getSheetList(int) return java.sql.Array';


  procedure StAX_setContent(key in number, content in blob)
  as language java 
  name 'db.office.spreadsheet.ReadContext.setContent(int, java.sql.Blob)';
    

  procedure StAX_setSharedStrings(key in number, sharedStrings in blob)
  as language java 
  name 'db.office.spreadsheet.ReadContext.setSharedStrings(int, java.sql.Blob)';


  procedure StAX_addSheet(key in number, idx in number, sheet in blob, comments in blob)
  as language java 
  name 'db.office.spreadsheet.ReadContext.addSheet(int, int, java.sql.Blob, java.sql.Blob)';

  
  function StAX_iterateContext(key in number, nrows in number) 
  return ExcelTableCellList
  as language java 
  name 'db.office.spreadsheet.ReadContext.iterate(int, int) return java.sql.Array';
 
 
  procedure StAX_closeContext(key in number)
  as language java 
  name 'db.office.spreadsheet.ReadContext.terminate(int)';


  procedure XDB_createReader (
    ctx_id       in pls_integer
  , reader       in out nocopy t_xdb_reader
  , sheets       in t_sheets
  , start_row    in pls_integer
  , end_row      in pls_integer
  )
  is
    pragma autonomous_transaction;
    
    xq_expr  varchar2(128) := '/worksheet/sheetData/row';
    info     t_cell_info;
    res      integer;
    
    query     varchar2(2000) := q'{
select t.xid1, x.r, x.t, decode(x.t, 'inlineStr', x.s, x.v) as v
from $$TAB t
   , xmltable(
       xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
     , '$$XQ'
       passing t.data
       columns r  varchar2(10)  path '@r'
             , t  varchar2(10)  path '@t'
             , v  varchar2($$0) path 'v'
             , s  varchar2($$0) path 'is'
     ) x
where t.id = :1
}';
    
  begin
    
    create_tmp_table;
    
    for i in 1 .. sheets.count loop
      load_tmp_table(ctx_id, blob2xml(sheets(i).content), i);
    end loop;
    
    commit;
    
    if start_row is not null then
      xq_expr := xq_expr || '[@r>=' || start_row || ']';
    end if;
    if end_row is not null then
      xq_expr := xq_expr || '[@r<=' || end_row || ']';
    end if;
    xq_expr := xq_expr || '/c';
      
    query := replace(query, '$$TAB', TMP_TABLE_QNAME);
    query := replace(query, '$$XQ', xq_expr);
    query := replace(query, '$$0', MAX_STRING_SIZE);
    
    debug(query);
    
    reader.c := dbms_sql.open_cursor;
    dbms_sql.parse(reader.c, query, dbms_sql.native);
    dbms_sql.bind_variable(reader.c, '1', ctx_id);
    dbms_sql.define_column(reader.c, 1, info.sheetIdx);
    dbms_sql.define_column(reader.c, 2, info.cellRef, 10);
    dbms_sql.define_column(reader.c, 3, info.cellType, 10);
    dbms_sql.define_column(reader.c, 4, info.cellValue, MAX_STRING_SIZE);
    res := dbms_sql.execute(reader.c);
    
  end;
  
  
  procedure XDB_closeReader (
    ctx_id  in pls_integer
  , reader  in out nocopy t_xdb_reader
  )
  is
    pragma autonomous_transaction;
  begin
    debug('Table close XDB');
    dbms_sql.close_cursor(reader.c);
    delete_tmp_table(ctx_id);
    commit;
  end;
  

  function OX_hasContentType (
    p_doc          in t_exceldoc
  , p_content_type in varchar2
  )
  return boolean
  is
    cnt  pls_integer;
  begin

    select count(*)
    into cnt
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/content-types')
         , '/Types/*[@ContentType=$type]'
           passing p_doc.content_map
                 , p_content_type as "type"
         ) x ;

    return (cnt > 0);
    
  end;  


  function OX_getPathByType (
    p_doc          in t_exceldoc
  , p_content_type in varchar2
  )
  return varchar2
  is
    partPath varchar2(256);
  begin

    select ltrim(x.partname, '/')
    into partPath
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/content-types')
         , '/Types/Override[@ContentType=$type]'
           passing p_doc.content_map
                 , p_content_type as "type"
           columns partname varchar2(256) path '@PartName'
         ) x ;

    return partPath;
    
  exception
    when no_data_found then
      return null;
  end;

/*
  function OX_getWorkbookPath (
    archive  in t_archive
  )
  return varchar2
  is  
    l_path   varchar2(260);
    l_rels   xmltype := Zip_getXML(archive, '_rels/.rels');
  begin
   
    select x.partname
    into l_path
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships')
         , '/Relationships/Relationship[@Type=$type]/@Target'
           passing l_rels
                 , RS_OFFICEDOC as "type"
           columns partname varchar2(256) path '.'
         ) x ;

    return l_path;

  end;
*/  
  
  procedure OX_setWorkbookInfo (
    archive in t_archive
  , doc     in out nocopy t_exceldoc 
  )
  is
    l_rels     xmltype := Zip_getXML(archive, '_rels/.rels');
    l_reltype  varchar2(256);
  begin
    select x.partname
         , x.reltype
    into doc.wb.path
       , l_reltype
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships')
         , '/Relationships/Relationship[fn:ends-with(@Type,$type)]'
           passing l_rels
                 , '/relationships/officeDocument' as "type"
           columns partname varchar2(256) path '@Target'
                 , reltype  varchar2(256) path '@Type'
         ) x ;
         
     doc.is_strict := ( l_reltype like 'http://purl.oclc.org/ooxml/%' );
  end;
  

  function OX_readSheets (
    doc  in out nocopy t_exceldoc
  )
  return ExcelTableSheetList
  is
  
    type relMap_t is table of varchar2(260) index by varchar2(1024);
  
    cursor c_rels (rel_type in varchar2) is
    select x.relId, x.sheetPath
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships')
         , 'for $rel in $rels/Relationships/Relationship
            where $rel/@Type = $relType
            return element r { 
              element rId { data($rel/@Id) }
            , element path { resolve-uri($rel/@Target, $path) }
            }'
           passing doc.wb.rels as "rels"
                 --, RS_WORKSHEET as "rsType"
                 , doc.wb.path as "path"
                 , rel_type as "relType"
            columns relId      varchar2(1024) path 'rId'
                  , sheetPath  varchar2(260)  path 'path'
         ) x
    ;
    
    
    cursor c_sheets is
    select x.sheetIdx, x.relId, x.sheetName
    from xmltable(
           xmlnamespaces(
             default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
           , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' as "r"
           )
         , '/workbook/sheets/sheet'
           passing doc.wb.content
           columns sheetIdx   for ordinality
                 , relId      varchar2(1024) path '@r:id'
                 , sheetName  varchar2(128)  path '@name'
         ) x
    ;

    cursor c_sheets_strict is
    select x.sheetIdx, x.relId, x.sheetName
    from xmltable(
           xmlnamespaces(
             default 'http://purl.oclc.org/ooxml/spreadsheetml/main'
           , 'http://purl.oclc.org/ooxml/officeDocument/relationships' as "r"
           )
         , '/workbook/sheets/sheet'
           passing doc.wb.content
           columns sheetIdx   for ordinality
                 , relId      varchar2(1024) path '@r:id'
                 , sheetName  varchar2(128)  path '@name'
         ) x
    ;
    
    relMap            relMap_t;
    sheetEntries_bin  xutl_xlsb.SheetEntries_T;
    sheet             t_sheetEntry;
    sheetList         ExcelTableSheetList := ExcelTableSheetList();

  begin
    
    for r in c_rels(get_prop('RS_WORKSHEET',doc.is_strict)) loop
      relMap(r.relId) := r.sheetPath;
    end loop;
    
    sheetList.extend(relMap.count);
    
    if doc.is_xlsb then
      
      sheetEntries_bin := xutl_xlsb.get_sheetEntries(doc.wb.content_binary);
      for i in 1 .. sheetEntries_bin.count loop
        sheet.idx := i;
        sheet.path := relMap(sheetEntries_bin(i).relId);
        doc.wb.sheetmap(sheetEntries_bin(i).name) := sheet;
        sheetList(i) := sheetEntries_bin(i).name;
      end loop;
    
    else
    
      if doc.is_strict then

        for r in c_sheets_strict loop
          sheet.idx := r.sheetIdx;
          sheet.path := relMap(r.relId);
          doc.wb.sheetmap(r.sheetName) := sheet;
          sheetList(r.sheetIdx) := r.sheetName;
        end loop;
    
      else
        
        for r in c_sheets loop
          sheet.idx := r.sheetIdx;
          sheet.path := relMap(r.relId);
          doc.wb.sheetmap(r.sheetName) := sheet;
          sheetList(r.sheetIdx) := r.sheetName;
        end loop;
      
      end if;
    
    end if;
    
    return sheetList;
  
  end;


  procedure readStringsFromBinXML (
    ctx_id        in pls_integer
  , query         in varchar2
  , xml_content   in xmltype 
  )
  is
    pragma autonomous_transaction;
    l_query varchar2(2000);
  begin
    create_tmp_table;
    load_tmp_table(ctx_id, xml_content);
    l_query := replace(query, '$$XML', '(select data from '||TMP_TABLE_QNAME||' where id = :1)');   
    execute immediate l_query bulk collect into ctx_cache(ctx_id).string_cache using ctx_id;
    delete_tmp_table(ctx_id);
    commit;
  end;


  procedure readStringsDOM (
    sharedStrings  in xmltype
  , string_cache   in out nocopy t_strings
  , is_strict      in boolean
  )
  is
    SML_NSMAP    constant varchar2(256) := get_prop('SML_NSMAP',is_strict);
    domDoc       dbms_xmldom.DOMDocument;
    docNode      dbms_xmldom.DOMNode;
    nlist        dbms_xmldom.DOMNodeList;
    node         dbms_xmldom.DOMNode;
    uniqueCount  pls_integer;
    charBuf      varchar2(32767);
    
  begin

    domDoc := dbms_xmldom.newDOMDocument(sharedStrings);
    docNode := dbms_xmldom.makeNode(domDoc);
    uniqueCount := dbms_xslprocessor.valueOf(docNode, '/sst/@uniqueCount', SML_NSMAP);
    string_cache := t_strings();
    string_cache.extend(uniqueCount);
    
    nlist := dbms_xslprocessor.selectNodes(docNode, '/sst/si', SML_NSMAP);
    
    for i in 0 .. dbms_xmldom.getLength(nlist) - 1 loop
      node := dbms_xmldom.item(nlist, i);
      begin
        dbms_xslprocessor.valueOf(node, '.', charBuf, SML_NSMAP);
        string_cache(i+1).strval := charBuf;
      exception
        when value_error then
          readclob(dbms_xslprocessor.selectNodes(node, '//t/text()', SML_NSMAP), string_cache(i+1).lobval);
      end;
      dbms_xmldom.freeNode(node);
    end loop;
    
    dbms_xmldom.freeNodeList(nlist);
    dbms_xmldom.freeDocument(domDoc);
  
  end;


  procedure OX_loadStringCache (
    archive  in t_archive
  , doc      in t_exceldoc
  , ctx_id   in binary_integer
  ) 
  is
    l_path     varchar2(260) := OX_getPathByType(doc, CT_SHAREDSTRINGS);
    l_xml      xmltype;
    
    l_query    varchar2(2000) := 
    q'{select $$HINT x.strval, x.lobval 
       from xmltable(
              xmlnamespaces(default '$$NS')
            , '/sst/si'
              passing $$XML
              columns strval  varchar2($$0) path '.[string-length() le $$1]'
                    , lobval  clob          path '.[string-length() gt $$1]') x}';

  begin
    
    if l_path is not null then
      
      l_xml := Zip_getXML(archive, l_path);
      l_query := replace(l_query, '$$NS', get_prop('DEFAULT_NS',doc.is_strict));
      l_query := replace(l_query, '$$0', MAX_STRING_SIZE);
      l_query := replace(l_query, '$$1', VC2_MAXSIZE);
      

      /* =======================================================================================
       From 11.2.0.4 and onwards, the new XQuery VM allows very efficient
       evaluation over transient XMLType instances.
       For prior versions, we'll first insert the XML document into a temp XMLType table using 
       Binary XML storage. The temp table is created on-the-fly, not a good practice but a lot
       faster than the alternative using DOM.
       Version 11.2.0.1 has limited support for CLOB, so using DOM to extract large text nodes.
      ======================================================================================= */
      if dbms_db_version.version >= 12 or DB_VERSION like '11.2.0.4%' then
        
        l_query := replace(l_query, '$$HINT', '/*+ no_xml_query_rewrite */');
        l_query := replace(l_query, '$$XML', ':1');
        execute immediate l_query 
        bulk collect into ctx_cache(ctx_id).string_cache
        using l_xml;
      
      elsif DB_VERSION like '11.2.0.2%' or DB_VERSION like '11.2.0.3%' then
        
        l_query := replace(l_query, '$$HINT', null);
        readStringsFromBinXML(ctx_id, l_query, l_xml);
      
      else
        
        readStringsDOM(l_xml, ctx_cache(ctx_id).string_cache, doc.is_strict);
        
      end if;
      
    end if;
    
  end;


  function OX_getSheetComments (
    archive  in t_archive
  , sheet    in t_sheet
  , doc      in t_exceldoc
  )
  return blob
  is
    l_comments_path    varchar2(256);
    l_sheet_rels_path  varchar2(256);
    l_comments_part    blob;
    l_sheet_rels       xmltype;
    
    cursor c_part (rel_type in varchar2) is
    select x.partname
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships')
         , 'for $r in /Relationships/Relationship
            where $r/@Type = $relType
            return resolve-uri($r/@Target, $path)'
           passing l_sheet_rels
                 --, RS_COMMENTS as "relType"
                 , rel_type as "relType"
                 , sheet.path as "path"
           columns partname varchar2(256) path '.'
         ) x ;
    
  begin
    
    l_sheet_rels_path := regexp_replace(sheet.path, '(.*)/(.*)$', '\1/_rels/\2.rels');
    
    if Zip_hasEntry(archive, l_sheet_rels_path) then
    
      l_sheet_rels := Zip_getXML(archive, l_sheet_rels_path);

      -- get path of the comments part
      open c_part(get_prop('RS_COMMENTS',doc.is_strict));
      fetch c_part into l_comments_path;
      close c_part;
      
      if l_comments_path is not null then
        l_comments_part := Zip_getEntry(archive, l_comments_path);
      end if;
          
    end if;
    
    return l_comments_part;
  
  end;


  function OX_readComments (
    content  in blob
  , doc      in t_exceldoc
  )
  return t_commentMap
  is
    l_comments_part   xmltype := blob2xml(content);
    l_comments        t_commentMap;
         
    cursor c_comments is
    select x.cell_ref, x.cell_cmt
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
         , '/comments/commentList/comment'
           passing l_comments_part
           columns cell_ref varchar2(10)   path '@ref'
                 , cell_cmt varchar2(4000) path 'text'
         ) x
    ;

    cursor c_comments_strict is
    select x.cell_ref, x.cell_cmt
    from xmltable(
           xmlnamespaces(default 'http://purl.oclc.org/ooxml/spreadsheetml/main')
         , '/comments/commentList/comment'
           passing l_comments_part
           columns cell_ref varchar2(10)   path '@ref'
                 , cell_cmt varchar2(4000) path 'text'
         ) x
    ;
    
  begin
     
    if doc.is_strict then
      for r in c_comments_strict loop
        l_comments(r.cell_ref) := r.cell_cmt;
      end loop;      
    else    
      for r in c_comments loop
        l_comments(r.cell_ref) := r.cell_cmt;
      end loop;      
    end if;
    
    return l_comments;
  
  end;


  procedure OX_openWorkbook (
    archive in t_archive
  , doc     in out nocopy t_exceldoc
  ) 
  is
  begin
    
    doc.content_map := Zip_getXML(archive, '[Content_Types].xml');
    -- Excel Binary File (.xlsb)?
    doc.is_xlsb := OX_hasContentType(doc, CT_XL_BINARY_FILE);
    --doc.workbook.path := OX_getWorkbookPath(archive);
    OX_setWorkbookInfo(archive, doc);
    
    if doc.is_xlsb then
      doc.wb.content_binary := Zip_getEntry(archive, doc.wb.path);
    else
      doc.wb.content := Zip_getXML(archive, doc.wb.path);
    end if;
    
    doc.wb.rels := Zip_getXML(archive, regexp_replace(doc.wb.path, '(.*)/(.*)$', '\1/_rels/\2.rels'));
  
  end;
  
  
  procedure OX_initDOMReader (
    reader         in out nocopy t_dom_reader
  , sheet_content  in xmltype
  , start_row      in pls_integer default null
  , end_row        in pls_integer default null
  )
  is
  begin
    
    reader.doc := dbms_xmldom.newDOMDocument(sheet_content);

    if reader.xpath is null then
      reader.xpath := '/worksheet/sheetData/row';
      if start_row is not null then
        reader.xpath := reader.xpath || '[@r>=' || start_row || ']';
      end if;
      if end_row is not null then
        reader.xpath := reader.xpath || '[@r<=' || end_row || ']';
      end if;
    end if;
        
    reader.rlist := dbms_xslprocessor.selectNodes(
                      n         => dbms_xmldom.makeNode(reader.doc)
                    , pattern   => reader.xpath
                    , namespace => reader.ns_map
                    );
    
    reader.rlist_idx := 0;
    if dbms_xmldom.isNull(reader.rlist) then
      reader.rlist_size := 0;
    else
      reader.rlist_size := dbms_xmldom.getLength(reader.rlist);
    end if;
    
  end;


  procedure OX_openNextSheet (
    ctx  in out nocopy t_context
  )
  is
    has_next  boolean := (ctx.curr_sheet < ctx.sheets.count);
  begin
    
    if ctx.curr_sheet > 0 then
      -- free open resources
      dbms_xmldom.freeNodeList(ctx.dom_reader.rlist);
      dbms_xmldom.freeDocument(ctx.dom_reader.doc);
    else
      -- set namespace map
      ctx.dom_reader.ns_map := get_prop('SML_NSMAP',ctx.is_strict);
    end if;
    
    while has_next loop
      ctx.curr_sheet := ctx.curr_sheet + 1;
      OX_initDOMReader(
        ctx.dom_reader
      , blob2xml(ctx.sheets(ctx.curr_sheet).content)
      , ctx.def_cache.range.start_ref.r
      , ctx.def_cache.range.end_ref.r
      );
      
      exit when ctx.dom_reader.rlist_size != 0;
      has_next := (ctx.curr_sheet < ctx.sheets.count);
    end loop;
    
    if not has_next then
      ctx.done := true;
    end if;
    
  end;


  procedure OX_openWorksheet (
    archive  in  t_archive
  , sheetFilter  in  anydata
  , ctx_id   in  binary_integer
  )
  is
  
    l_xldoc       t_exceldoc;
    l_sst         blob;
    l_sheets      t_sheets;
    l_key         number;  
    l_read_method binary_integer := ctx_cache(ctx_id).read_method;
    l_tab_def     QI_definition_t := ctx_cache(ctx_id).def_cache;
    l_start_row   pls_integer := l_tab_def.range.start_ref.r;
    l_end_row     pls_integer := l_tab_def.range.end_ref.r;
    
    documentSheetList     ExcelTableSheetList;
    
  begin
    
    OX_openWorkbook(archive, l_xldoc);
    documentSheetList := OX_readSheets(l_xldoc);
    l_sheets := filterSheetList(documentSheetList, sheetFilter);
    
    for i in 1 .. l_sheets.count loop
      l_sheets(i).path := l_xldoc.wb.sheetMap(l_sheets(i).name).path;
      l_sheets(i).content := Zip_getEntry(archive, l_sheets(i).path);
    end loop;
    
    -- comments
    if l_tab_def.hasComment then
      for i in 1 .. l_sheets.count loop
        l_sheets(i).comments := OX_getSheetComments(archive, l_sheets(i), l_xldoc);
        
        if l_sheets(i).comments is not null and not(l_read_method = STREAM_READ or l_xldoc.is_xlsb) then
          ctx_cache(ctx_id).comments(i) := OX_readComments(l_sheets(i).comments, l_xldoc);
          dbms_lob.freetemporary(l_sheets(i).comments);
          l_sheets(i).comments := null;
        end if;
        
      end loop;
    end if;
    
    ctx_cache(ctx_id).sheets := l_sheets;
    
    if l_xldoc.is_xlsb then
      
      l_sst := Zip_getEntry(archive, OX_getPathByType(l_xldoc, CT_SHAREDSTRINGS_BIN));
      
      l_key := xutl_xlsb.new_context(
        l_sst
      , get_column_list(l_tab_def)
      , l_start_row                
      , l_end_row
      );
      
      if dbms_lob.istemporary(l_sst) = 1 then
        dbms_lob.freetemporary(l_sst);
      end if;
      
      for i in 1 .. l_sheets.count loop
        xutl_xlsb.add_sheet(l_key, l_sheets(i).content, l_sheets(i).comments);
      end loop;      
      
      ctx_cache(ctx_id).extern_key := l_key;
      ctx_cache(ctx_id).file_type := FILE_XLSB;   
      
    else
    
      ctx_cache(ctx_id).file_type := FILE_XLSX;
      ctx_cache(ctx_id).is_strict := l_xldoc.is_strict;
    
      case l_read_method 
      when DOM_READ then
        
        OX_loadStringCache(archive, l_xldoc, ctx_id);
        OX_openNextSheet(ctx_cache(ctx_id));
        
      when STREAM_READ then
        
        l_sst := Zip_getEntry(archive, OX_getPathByType(l_xldoc, CT_SHAREDSTRINGS));
        l_key := StAX_createContext(
                   'OOX'
                 , get_column_list(l_tab_def)
                 , nvl(l_start_row, 1)
                 , nvl(l_end_row, -1)
                 , 32767 --MAX_STRING_SIZE
                 );
                 
        StAX_setSharedStrings(l_key, l_sst);
        
        for i in 1 .. l_sheets.count loop
           StAX_addSheet(l_key, i, l_sheets(i).content, l_sheets(i).comments);
        end loop;
        
        if l_sst is not null then
          dbms_lob.freetemporary(l_sst);
        end if;
        
        ctx_cache(ctx_id).extern_key := l_key;
      
      when STREAM_READ_XDB then
        
        OX_loadStringCache (archive, l_xldoc, ctx_id);
        XDB_createReader(ctx_id, ctx_cache(ctx_id).xdb_reader, l_sheets, l_start_row, l_end_row);
        
      else
        -- invalid read method specified
        raise_application_error(-20725, utl_lms.format_message(INVALID_READ_METHOD, l_read_method));
      
      end case;
    
    end if;
    
  end;


  procedure BF_openWorksheet (
    p_file     in  blob
  , p_password in varchar2
  , p_sheet    in  anydata
  , p_ctx_id   in  binary_integer
  )
  is
  
    l_key           number;
    l_tab_def       QI_definition_t := ctx_cache(p_ctx_id).def_cache;
    l_start_row     pls_integer := l_tab_def.range.start_ref.r;
    l_end_row       pls_integer := l_tab_def.range.end_ref.r;
    --l_comment_list  ExcelTableCellList;
    --l_comment_map   t_commentMap;
    --l_comments      t_comments;
    sheetList       ExcelTableSheetList := ExcelTableSheetList();
    filteredSheets  t_sheets;
    
  begin
    
    ctx_cache(p_ctx_id).file_type := FILE_XLS;
    
    l_key := xutl_xls.new_context(
      p_file       => p_file
    , p_password   => p_password
    , p_cols       => get_column_list(l_tab_def)
    , p_firstRow   => l_start_row
    , p_lastRow    => l_end_row
    , p_readNotes  => l_tab_def.hasComment
    );
    
    ctx_cache(p_ctx_id).extern_key := l_key;
    
    filteredSheets := filterSheetList(xutl_xls.get_sheetList(l_key), p_sheet);
    ctx_cache(p_ctx_id).sheets := filteredSheets;
    
    sheetList.extend(filteredSheets.count);
    for i in 1 .. filteredSheets.count loop
      sheetList(i) := filteredSheets(i).name;
    end loop;
    
    xutl_xls.add_sheets(l_key, sheetList);
    
    /*
    if l_tab_def.hasComment then
      for i in 1 .. filteredSheets.count loop
        l_comment_list := xutl_xls.get_comments(l_key, filteredSheets(i).name);
        l_comment_map.delete;
        for j in 1 .. l_comment_list.count loop
          l_comment_map(l_comment_list(j).cellCol || to_char(l_comment_list(j).cellRow)) := l_comment_list(j).cellData.accessVarchar2();
        end loop;
        l_comments(i) := l_comment_map;
      end loop;
      ctx_cache(p_ctx_id).comments := l_comments;
    end if;
    */
    
  end;


  function ODS_getEncryptData (
    manifest   in xmltype
  , entryName  in varchar2
  )
  return xmltype
  is
    output xmltype;
  begin

    select column_value
    into output
    from xmltable(
      xmlnamespaces(
        'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0' as "m"
      , default 'urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'
      )
    , '/manifest/file-entry[@m:full-path=$entryName]/encryption-data'
      passing manifest
            , entryName as "entryName"
    );
    
    return output;
    
  exception
    when no_data_found then
      return null;
  end;


  function ODS_getSheetList (
    content_part  in xmltype
  )
  return ExcelTableSheetList
  is
    sheetList  ExcelTableSheetList;
  begin
    select sheetName
    bulk collect into sheetList
    from xmltable(
           xmlnamespaces(
             'urn:oasis:names:tc:opendocument:xmlns:office:1.0' as "o"
           , 'urn:oasis:names:tc:opendocument:xmlns:table:1.0' as "t"
           )
         , '/o:document-content/o:body/o:spreadsheet/t:table'
           passing content_part
           columns sheetIdx  for ordinality
                 , sheetName varchar2(128) path '@t:name'
         )
    ;    
    return sheetList;
  end;


  procedure ODS_initDOMReader (
    reader    in out nocopy t_dom_reader
  , sheetIdx  in pls_integer
  )
  is
  begin

    reader.xpath := '/o:document-content/o:body/o:spreadsheet/t:table[$SHEETIDX]/t:table-row';
    reader.xpath := replace(reader.xpath, '$SHEETIDX', to_char(sheetIdx));
        
    reader.rlist := dbms_xslprocessor.selectNodes(
                      n         => dbms_xmldom.makeNode(reader.doc)
                    , pattern   => reader.xpath
                    , namespace => ODF_OFFICE_TABLE_NSMAP
                    );
    
    reader.rlist_idx := 0;
    if dbms_xmldom.isNull(reader.rlist) then
      reader.rlist_size := 0;
    else
      reader.rlist_size := dbms_xmldom.getLength(reader.rlist);
    end if;
    
  end;


  procedure ODS_openNextSheet (
    ctx  in out nocopy t_context
  )
  is
    has_next  boolean := (ctx.curr_sheet < ctx.sheets.count);
  begin
    if ctx.curr_sheet > 0 then
      -- free open resources
      dbms_xmldom.freeNodeList(ctx.dom_reader.rlist);
      --dbms_xmldom.freeDocument(ctx.dom_reader.doc);
    end if;
    
    while has_next loop
      ctx.curr_sheet := ctx.curr_sheet + 1;
      ctx.src_row := 0;
      ctx.row_repeat := 0;
      ctx.tmp_row := null;
      ODS_initDOMReader(ctx.dom_reader, ctx.sheets(ctx.curr_sheet).idx);     
      exit when ctx.dom_reader.rlist_size != 0;
      has_next := (ctx.curr_sheet < ctx.sheets.count);
    end loop;
    
    if not has_next then
      ctx.done := true;
    end if;
    
  end;


  procedure ODS_openContent (
    archive      in t_archive
  , sheetFilter  in anydata
  , ctx_id       in binary_integer
  , password     in varchar2
  )
  is
  
    CONTENT_PART_NAME  constant varchar2(256) := 'content.xml';
    l_content     blob;
    l_content_xml xmltype;
    l_manifest    xmltype;
    l_enc_data    xmltype;
    
    l_sheets      t_sheets;
    l_key         number;
    l_read_method binary_integer := ctx_cache(ctx_id).read_method;
    l_tab_def     QI_definition_t := ctx_cache(ctx_id).def_cache;
    l_start_row   pls_integer := l_tab_def.range.start_ref.r;
    l_end_row     pls_integer := l_tab_def.range.end_ref.r;
      
  begin
    
    ctx_cache(ctx_id).file_type := FILE_ODS;
    l_manifest := Zip_getXML(archive, 'META-INF/manifest.xml');
    l_enc_data := ODS_getEncryptData(l_manifest, CONTENT_PART_NAME);
    
    -- is the content encrypted?
    if l_enc_data is null then
      --l_content := Zip_getXML(archive, CONTENT_PART_NAME);
      l_content := Zip_getEntry(archive, CONTENT_PART_NAME);
    else
      if password is null then
        raise_application_error(-20730, NO_PASSWORD);
      end if;
      --l_content := blob2xml(xutl_offcrypto.get_part_ODF(Zip_getEntry(archive, CONTENT_PART_NAME), l_enc_data, password));
      l_content := xutl_offcrypto.get_part_ODF(Zip_getEntry(archive, CONTENT_PART_NAME), l_enc_data, password);
    end if;
    
    case l_read_method
    when DOM_READ then
      l_content_xml := blob2xml(l_content);
      ctx_cache(ctx_id).sheets := filterSheetList(ODS_getSheetList(l_content_xml), sheetFilter);
      ctx_cache(ctx_id).dom_reader.doc := dbms_xmldom.newDOMDocument(l_content_xml);
      ODS_openNextSheet(ctx_cache(ctx_id));
      
    when STREAM_READ then
      l_key := StAX_createContext(
                 'ODF'
               , get_column_list(l_tab_def)
               , nvl(l_start_row, 1)
               , nvl(l_end_row, -1)
               , 32767 --MAX_STRING_SIZE
               );
      StAX_setContent(l_key, l_content);
      -- get sheets and filter
      l_sheets := filterSheetList(StAX_getSheetList(l_key), sheetFilter);
      ctx_cache(ctx_id).sheets := l_sheets;
      -- add sheets
      for i in 1 .. l_sheets.count loop
        --dbms_output.put_line(l_sheets(i).name);
        StAX_addSheet(l_key, i, null, null);
      end loop;
            
      ctx_cache(ctx_id).extern_key := l_key;
      dbms_lob.freetemporary(l_content);
      
    else
      -- invalid read method specified
      raise_application_error(-20725, utl_lms.format_message(INVALID_READ_METHOD, l_read_method));
    end case;
    
  end;


  function XSS_getSheetList (
    xml_content in xmltype
  )
  return ExcelTableSheetList
  is
    sheetList  ExcelTableSheetList;
  begin
    select sheetName
    bulk collect into sheetList
    from xmltable(
           xmlnamespaces(
             default 'urn:schemas-microsoft-com:office:spreadsheet'
           , 'urn:schemas-microsoft-com:office:spreadsheet' as "ss"
           )
         , '/Workbook/Worksheet'
           passing xml_content
           columns sheetIdx  for ordinality
                 , sheetName varchar2(128) path '@ss:Name'
         )
    ;    
    return sheetList;
  end;


  procedure XSS_initDOMReader (
    reader    in out nocopy t_dom_reader
  , sheetIdx  in pls_integer
  )
  is
  begin

    reader.xpath := '/Workbook/Worksheet[$SHEETIDX]/Table/Row';
    reader.xpath := replace(reader.xpath, '$SHEETIDX', to_char(sheetIdx));
        
    reader.rlist := dbms_xslprocessor.selectNodes(
                      n         => dbms_xmldom.makeNode(reader.doc)
                    , pattern   => reader.xpath
                    , namespace => XSS_DFLT_NSMAP
                    );
    
    reader.rlist_idx := 0;
    if dbms_xmldom.isNull(reader.rlist) then
      reader.rlist_size := 0;
    else
      reader.rlist_size := dbms_xmldom.getLength(reader.rlist);
    end if;
    
  end;


  procedure XSS_openNextSheet (
    ctx  in out nocopy t_context
  )
  is
    has_next  boolean := (ctx.curr_sheet < ctx.sheets.count);
  begin
    if ctx.curr_sheet > 0 then
      -- free open resources
      dbms_xmldom.freeNodeList(ctx.dom_reader.rlist);
    end if;
    
    while has_next loop
      ctx.curr_sheet := ctx.curr_sheet + 1;
      ctx.src_row := 0;
      ctx.row_repeat := 0;
      ctx.tmp_row := null;
      XSS_initDOMReader(ctx.dom_reader, ctx.sheets(ctx.curr_sheet).idx);     
      exit when ctx.dom_reader.rlist_size != 0;
      has_next := (ctx.curr_sheet < ctx.sheets.count);
    end loop;
    
    if not has_next then
      ctx.done := true;
    end if;
    
  end;
  

  procedure XSS_openContent (
    content      in blob
  , sheetFilter  in anydata
  , ctx_id       in binary_integer
  )
  is
    l_content     xmltype := blob2xml(content);
  begin
    
    ctx_cache(ctx_id).file_type := FILE_XSS;
    ctx_cache(ctx_id).sheets := filterSheetList(XSS_getSheetList(l_content), sheetFilter); 
    ctx_cache(ctx_id).dom_reader.doc := dbms_xmldom.newDOMDocument(l_content);
    XSS_openNextSheet(ctx_cache(ctx_id));
  
  end;


  procedure openSpreadsheet (
    p_file      in blob
  , p_password  in varchar2
  , p_sheet     in anydata
  , p_ctx_id    in binary_integer
  )
  is
  
    cdf       xutl_cdf.cdf_handle;
    opc       blob;
    archive   t_archive;
    mimetype  varchar2(100);
  
  begin
    
    if is_opc_package(p_file) then
      
      archive := Zip_openArchive(p_file);
      if Zip_hasEntry(archive, 'mimetype') then
        mimetype := utl_raw.cast_to_varchar2(dbms_lob.substr(Zip_getEntry(archive, 'mimetype')));
        if mimetype = MIMETYPE_ODS then
          -- open ODS
          ODS_openContent(archive, p_sheet, p_ctx_id, p_password);
        else
          -- unsupported document format
          raise_application_error(-20720, INVALID_DOCUMENT);
        end if;
      else
        -- process as an Open Office XML document
        OX_openWorksheet(archive, p_sheet, p_ctx_id);
      end if;
        
    elsif xutl_cdf.is_cdf(p_file) then
      
      cdf := xutl_cdf.open_file(p_file);
      if xutl_cdf.stream_exists(cdf, '/Workbook') then
        BF_openWorksheet(xutl_cdf.get_stream(cdf, '/Workbook'), p_password, p_sheet, p_ctx_id);
        xutl_cdf.close_file(cdf);
      elsif xutl_cdf.stream_exists(cdf, '/EncryptedPackage') then
        if p_password is null then
          raise_application_error(-20730, NO_PASSWORD);
        end if;
        opc := xutl_offcrypto.get_package(cdf, p_password);
        archive := Zip_openArchive(opc);
        OX_openWorksheet(archive, p_sheet, p_ctx_id);
      else
        xutl_cdf.close_file(cdf);
        raise_application_error(-20720, INVALID_DOCUMENT);
      end if;
      
    elsif is_xmlss(p_file) then
      --error(UNIMPLEMENTED_FEAT, 'XML Spreadsheet 2003', p_code => -20720);
      XSS_openContent(p_file, p_sheet, p_ctx_id);
    else
      raise_application_error(-20720, INVALID_DOCUMENT);
    end if;
    
  end;


  procedure openFlatFile (
    p_file       in clob
  , p_field_sep  in varchar2
  , p_line_term  in varchar2
  , p_text_qual  in varchar2
  , p_ctx_id     in binary_integer
  )
  is
    l_key      number;  
    l_tab_def  QI_definition_t := ctx_cache(p_ctx_id).def_cache;
    l_skip     pls_integer := nvl(l_tab_def.range.start_ref.r, 1) - 1;
    --l_end_row    pls_integer := l_tab_def.range.end_ref.r;
  begin
    
    if l_tab_def.isPositional then
      l_key := xutl_flatfile.new_context(p_file, get_position_list(l_tab_def.cols), l_skip, xutl_flatfile.TYPE_POSITIONAL);
    else
      l_key := xutl_flatfile.new_context(p_file, get_column_list(l_tab_def), l_skip, xutl_flatfile.TYPE_DELIMITED);
    end if;
    
    xutl_flatfile.set_file_descriptor(l_key, p_field_sep, p_line_term, p_text_qual);
    ctx_cache(p_ctx_id).extern_key := l_key;
    ctx_cache(p_ctx_id).file_type := FILE_FF;
    
  end;
  

  function getCells_DOM (
    ctx_id  in pls_integer
  , nrows   in number 
  )
  return ExcelTableCellList
  is
  
    SML_NSMAP   constant varchar2(256) := ctx_cache(ctx_id).dom_reader.ns_map;
    cells       ExcelTableCellList := ExcelTableCellList();
    cell        ExcelTableCell := ExcelTableCell(null, null, null, null, ctx_cache(ctx_id).curr_sheet, null);
    l_nrows     integer := nrows;
    refset      QI_column_ref_set_t := ctx_cache(ctx_id).def_cache.refSet; 
    info        t_cell_info; 
    row_node    dbms_xmldom.DOMNode;
    cell_nodes  dbms_xmldom.DOMNodeList;
    cell_node   dbms_xmldom.DOMNode;
    str         t_string_rec;
    empty_row   boolean;
    hasComment  boolean := ctx_cache(ctx_id).def_cache.hasComment;
    
  begin

    loop
        
      empty_row := true;
      row_node := dbms_xmldom.item(ctx_cache(ctx_id).dom_reader.rlist, ctx_cache(ctx_id).dom_reader.rlist_idx);
      ctx_cache(ctx_id).dom_reader.rlist_idx := ctx_cache(ctx_id).dom_reader.rlist_idx + 1;
      
      cell_nodes := dbms_xslprocessor.selectNodes(row_node, 'c', SML_NSMAP);
      
      for i in 0 .. dbms_xmldom.getLength(cell_nodes) - 1 loop
        
        cell_node := dbms_xmldom.item(cell_nodes, i);
        info.cellRef := dbms_xslprocessor.valueOf(cell_node, '@r');
        info.cellCol := rtrim(info.cellRef, DIGITS);
        info.cellRow := ltrim(info.cellRef, LETTERS);
        str := null;
        cell.cellData := null;
        
        if refSet.exists(info.cellCol) then

          info.cellValue := null;
          -- read cell value element as VARCHAR2 (fallback to CLOB if too long)
          begin
            info.cellValue := dbms_xslprocessor.valueOf(cell_node, 'v', SML_NSMAP);
          exception
            when value_error or buffer_too_small then
              readclob(dbms_xslprocessor.selectSingleNode(cell_node, 'v/text()', SML_NSMAP), str.lobval);
          end;
                    
          info.cellType := dbms_xslprocessor.valueOf(cell_node, '@t');

          if info.cellType is null or info.cellType = 'n' then
            if info.cellValue is not null then
              cell.cellData := anydata.ConvertNumber(to_number(replace(info.cellValue,'.',get_decimal_sep)));
            end if;
            
          elsif info.cellType = 's' then
          
            str := get_string(ctx_id, info.cellValue);
          
          elsif info.cellType = 'd' then
          
            cell.cellData := anydata.ConvertTimestamp(get_tstamp_val_iso8601(info.cellValue));
          
          elsif info.cellType = 'inlineStr' then
          
            -- read inline string value
            begin
              str.strval := dbms_xslprocessor.valueOf(cell_node, 'is', SML_NSMAP);
            exception
              when value_error or buffer_too_small then
                readclob(dbms_xslprocessor.selectNodes(cell_node, 'is//t/text()', SML_NSMAP), str.lobval);
            end;
            
          elsif info.cellType = 'b' then
          
            str.strval := case when info.cellValue = '1' then 'TRUE' else 'FALSE' end;
                
          else
            str.strval := info.cellValue;
          end if;
          
          cell.cellRow := info.cellRow;
          cell.cellCol := info.cellCol;
          cell.cellType := info.cellType;
          
          if str.strval is not null then
            cell.cellData := anydata.ConvertVarchar2(str.strval);
          elsif str.lobval is not null then
            cell.cellData := anydata.ConvertClob(str.lobval);
          end if;
          
          if hasComment and is_opt_set(refSet(info.cellCol), META_COMMENT) then
            cell.cellNote := get_comment(ctx_id, cell.sheetIdx, info.cellRef);
          end if;
          
          if cell.cellData is not null then
            cells.extend;
            cells(cells.last) := cell;
            empty_row := false;
          end if;
          
        end if;
        
        dbms_xmldom.freeNode(cell_node);
        
      end loop;
      
      dbms_xmldom.freeNodeList(cell_nodes);
      dbms_xmldom.freeNode(row_node);      
        
      -- no more row to read?
      if ctx_cache(ctx_id).dom_reader.rlist_idx = ctx_cache(ctx_id).dom_reader.rlist_size then
        OX_openNextSheet(ctx_cache(ctx_id));       
        cell.sheetIdx := ctx_cache(ctx_id).curr_sheet;
      end if;
      
      if not empty_row then
        l_nrows := l_nrows - 1;
      end if;
              
      exit when ctx_cache(ctx_id).done or l_nrows = 0;

    end loop;
    
    return cells;
    
  end;


  function getCells_ODS (
    ctx_id  in pls_integer
  , nrows   in number 
  )
  return ExcelTableCellList
  is
    cells        ExcelTableCellList := ExcelTableCellList();  
    tmp_row      ExcelTableCellList := ctx_cache(ctx_id).tmp_row;
    cell         ExcelTableCell := ExcelTableCell(null, null, null, null, null, null);
    l_nrows      integer := nrows;
    row_node     dbms_xmldom.DOMNode;
    row_repeat   pls_integer := ctx_cache(ctx_id).row_repeat;
    row_cnt      pls_integer;
    row_idx      pls_integer;
    row_num      pls_integer;
    src_row      pls_integer;
    
    l_start_row  pls_integer := nvl(ctx_cache(ctx_id).def_cache.range.start_ref.r, 1);
    l_end_row    pls_integer := ctx_cache(ctx_id).def_cache.range.end_ref.r;
    refset       QI_column_ref_set_t := ctx_cache(ctx_id).def_cache.refSet;

    str          t_string_rec;
    info         t_cell_info;
    node_cnt     pls_integer;
    node_idx     pls_integer;
    cell_nodes   dbms_xmldom.DOMNodeList;
    cell_node    dbms_xmldom.DOMNode;
    cell_repeat  pls_integer;
    cell_idx     pls_integer;
    cell_cnt     pls_integer;
    
    empty_row    boolean;
    /*
    procedure read_comment (
      cell_node  in dbms_xmldom.DOMNode
    , sheet_idx  in pls_integer
    , cell_ref   in varchar2
    )
    is
      nodeList  dbms_xmldom.DOMNodeList;
    begin
      nodeList := dbms_xslprocessor.selectNodes(cell_node, 'o:annotation/x:p', ODF_OFFICE_TEXT_NSMAP);
      if not dbms_xmldom.isNull(nodeList) then
        ctx_cache(ctx_id).comments(sheet_idx)(cell_ref) := string_join(nodeList).strval;
      end if;
    end;
    */
    function get_annotation (
      cell_node  in dbms_xmldom.DOMNode
    )
    return varchar2
    is
      nodeList    dbms_xmldom.DOMNodeList;
      annotation  t_string_rec;
    begin
      nodeList := dbms_xslprocessor.selectNodes(cell_node, 'o:annotation/x:p', ODF_OFFICE_TEXT_NSMAP);
      if not dbms_xmldom.isNull(nodeList) then
        annotation := string_join(nodeList);
        dbms_xmldom.freeNodeList(nodeList);
      end if;
      return annotation.strval;
    end;

  begin

    row_cnt := ctx_cache(ctx_id).dom_reader.rlist_size;
    row_idx := ctx_cache(ctx_id).dom_reader.rlist_idx;
    src_row := ctx_cache(ctx_id).src_row;
    cell.sheetIdx := ctx_cache(ctx_id).curr_sheet;
      
    loop
          
      src_row := src_row + 1;
      cell.cellRow := src_row;
        
      if row_repeat = 0 then  
      
        row_node := dbms_xmldom.item(ctx_cache(ctx_id).dom_reader.rlist, row_idx);
        row_idx := row_idx + 1;
        row_repeat := nvl(dbms_xslprocessor.valueOf(row_node, '@t:number-rows-repeated', ODF_TABLE_NSMAP), 1) - 1;
           
        -- read cell nodes if within the requested range
        if src_row >= l_start_row - row_repeat then

          tmp_row := ExcelTableCellList();
          empty_row := true;
          cell_nodes := dbms_xslprocessor.selectNodes(row_node, 't:table-cell', ODF_TABLE_NSMAP);
          node_cnt := dbms_xmldom.getLength(cell_nodes);
          node_idx := 0;
          cell_idx := 1;
          cell_repeat := 0;
            
          while node_idx < node_cnt loop
            
            --info.cellCol := base26encode(cell_idx);
            cell.cellCol := base26encode(cell_idx);   
            
            if cell_repeat = 0 then
                
              str := null;
                  
              cell_node := dbms_xmldom.item(cell_nodes, node_idx);
              node_idx := node_idx + 1;  
              cell_repeat := nvl(dbms_xslprocessor.valueOf(cell_node, '@t:number-columns-repeated', ODF_TABLE_NSMAP), 1) - 1;
                
              -- read cell value type
              info.cellType := dbms_xslprocessor.valueOf(cell_node, '@o:value-type', ODF_OFFICE_NSMAP);
                
              case info.cellType
              when 'float' then
                str.strval := dbms_xslprocessor.valueOf(cell_node, '@o:value', ODF_OFFICE_NSMAP);
                cell.cellData := anydata.ConvertNumber(to_number(replace(str.strval,'.',get_decimal_sep)));
                  
              when 'string' then
                begin
                  str.strval := dbms_xslprocessor.valueOf(cell_node, '@o:string-value', ODF_OFFICE_NSMAP);
                exception
                  when value_error or buffer_too_small then
                    readclob(
                      dbms_xslprocessor.selectSingleNode(cell_node, '@o:string-value', ODF_OFFICE_NSMAP)
                    , str.lobval
                    );
                end;

                if str.strval is null and str.lobval is null then
                  str := string_join(dbms_xslprocessor.selectNodes(cell_node, 'x:p', ODF_TEXT_NSMAP));
                end if;
                  
                if str.lobval is not null then
                  cell.cellData := anydata.ConvertClob(str.lobval);
                else
                  --cell.cellData := anydata.ConvertVarchar2(str.strval);
                  if lengthb(str.strval) <= MAX_STRING_SIZE then
                    cell.cellData := anydata.ConvertVarchar2(str.strval);
                  else
                    cell.cellData := anydata.ConvertClob(to_clob(str.strval));
                  end if;
                end if;
                
              when 'date' then
                str.strval := dbms_xslprocessor.valueOf(cell_node, '@o:date-value', ODF_OFFICE_NSMAP);
                cell.cellData := anydata.ConvertTimestamp(to_timestamp(str.strval,'YYYY-MM-DD"T"HH24:MI:SS.FF9'));
                
              when 'percentage' then
                str.strval := dbms_xslprocessor.valueOf(cell_node, '@o:value', ODF_OFFICE_NSMAP);
                cell.cellData := anydata.ConvertNumber(to_number(replace(str.strval,'.',get_decimal_sep)));

              when 'boolean' then
                str.strval := dbms_xslprocessor.valueOf(cell_node, '@o:boolean-value', ODF_OFFICE_NSMAP);
                cell.cellData := anydata.ConvertVarchar2(upper(str.strval));
              else
                --TODO : time-value?
                cell.cellData := null;
              end case;
                
              -- need comment?
              if refset.exists(cell.cellCol) then
                if is_opt_set(refSet(cell.cellCol), META_COMMENT) then
                  --read_comment(cell_node, cell.sheetIdx, cell.cellCol || cell.cellRow);
                  cell.cellNote := get_annotation(cell_node);
                end if;
              end if;
              
              dbms_xmldom.freeNode(cell_node);              
              
            else
                
              cell_repeat := cell_repeat - 1;
                
            end if;
            
            if cell.cellData is not null then
              -- if cell is part of a repeating row, save it in tmp_row
              if row_repeat > 0 then
                tmp_row.extend;
                tmp_row(tmp_row.last) := cell;
              end if;          
              
              if src_row >= l_start_row and refset.exists(cell.cellCol) then
                cells.extend;
                cells(cells.last) := cell;
                empty_row := false;
              end if;
            end if;
            
            cell_idx := cell_idx + 1;
              
          end loop;
            
          dbms_xmldom.freeNodeList(cell_nodes);
          dbms_xmldom.freeNode(row_node);
          
          if not empty_row then
            l_nrows := l_nrows - 1;
          elsif row_repeat > 0 then
            -- ignore empty repeating rows
            src_row := src_row + row_repeat;
            row_repeat := 0;
          end if;
                   
        end if;
          
      else
          
        if src_row >= l_start_row then
          
          -- copy cells from tmp_row
          cell_cnt := tmp_row.count;
          
          if cell_cnt != 0 then
            cells.extend(cell_cnt);
                        
            for i in 1 .. cell_cnt loop
              cell := tmp_row(i);
              cell.cellRow := src_row;
              cells(cells.last - cell_cnt + i) := cell;
            end loop;
            
            l_nrows := l_nrows - 1;
            
          end if;
          
        end if;
         
        row_repeat := row_repeat - 1;
        
      end if;
        
      -- no more row to read?
      if row_idx = row_cnt or src_row = l_end_row then 
        --ctx_cache(ctx_id).done := true;
        ODS_openNextSheet(ctx_cache(ctx_id));

        row_cnt := ctx_cache(ctx_id).dom_reader.rlist_size;
        row_idx := ctx_cache(ctx_id).dom_reader.rlist_idx;
        src_row := ctx_cache(ctx_id).src_row;
        cell.sheetIdx := ctx_cache(ctx_id).curr_sheet;
        row_repeat := 0;
        tmp_row := null;

      end if;
        

              
      exit when ctx_cache(ctx_id).done or l_nrows = 0;

    end loop;
      
    ctx_cache(ctx_id).dom_reader.rlist_idx := row_idx;
    ctx_cache(ctx_id).r_num := row_num;
    ctx_cache(ctx_id).src_row := src_row;
    -- save row repetition info across fetches
    ctx_cache(ctx_id).row_repeat := row_repeat;
    ctx_cache(ctx_id).tmp_row := tmp_row;
      
    return cells;
    
  end;


  function getCells_XSS (
    ctx_id  in pls_integer
  , nrows   in number 
  )
  return ExcelTableCellList
  is
    cells        ExcelTableCellList := ExcelTableCellList();  
    cell         ExcelTableCell := ExcelTableCell(null, null, null, null, null, null);
    l_nrows      integer := nrows;
    row_node     dbms_xmldom.DOMNode;
    row_cnt      pls_integer;
    row_idx      pls_integer;
    row_num      pls_integer;
    src_row      pls_integer;
    
    l_start_row  pls_integer := nvl(ctx_cache(ctx_id).def_cache.range.start_ref.r, 1);
    l_end_row    pls_integer := ctx_cache(ctx_id).def_cache.range.end_ref.r;
    refset       QI_column_ref_set_t := ctx_cache(ctx_id).def_cache.refSet;

    str          t_string_rec;
    info         t_cell_info;
    node_cnt     pls_integer;
    node_idx     pls_integer;
    cell_nodes   dbms_xmldom.DOMNodeList;
    cell_node    dbms_xmldom.DOMNode;
    cell_idx     pls_integer;
    data_node    dbms_xmldom.DOMNode;
    
    empty_row    boolean;
    /*
    procedure read_comment (
      cell_node  in dbms_xmldom.DOMNode
    , sheet_idx  in pls_integer
    , cell_ref   in varchar2
    )
    is
      str  t_string_rec;
    begin
      str.strval := dbms_xslprocessor.valueOf(cell_node, 'Comment/Data', XSS_DFLT_NSMAP);
      if str.strval is not null then
        ctx_cache(ctx_id).comments(sheet_idx)(cell_ref) := str.strval;
      end if;
    end;
    */
    function get_annotation (
      cell_node  in dbms_xmldom.DOMNode
    )
    return varchar2
    is
      annotation  t_string_rec;
    begin
      annotation.strval := dbms_xslprocessor.valueOf(cell_node, 'Comment/Data', XSS_DFLT_NSMAP);
      return annotation.strval;
    end;

  begin

    row_cnt := ctx_cache(ctx_id).dom_reader.rlist_size;
    row_idx := ctx_cache(ctx_id).dom_reader.rlist_idx;
    src_row := ctx_cache(ctx_id).src_row;
    cell.sheetIdx := ctx_cache(ctx_id).curr_sheet;
      
    loop
          
      src_row := src_row + 1;
      cell.cellRow := src_row;
      
      row_node := dbms_xmldom.item(ctx_cache(ctx_id).dom_reader.rlist, row_idx);
      row_idx := row_idx + 1;
      
      -- read row index
      src_row := nvl(dbms_xslprocessor.valueOf(row_node, '@ss:Index', XSS_SS_NSMAP), src_row);
           
      -- read cell nodes if within the requested range
      if src_row >= l_start_row then
        
        empty_row := true;
        cell_nodes := dbms_xslprocessor.selectNodes(row_node, 'Cell', XSS_DFLT_NSMAP);
        node_cnt := dbms_xmldom.getLength(cell_nodes);
        node_idx := 0;
        cell_idx := 1;
            
        while node_idx < node_cnt loop
          
          str := null;
                  
          cell_node := dbms_xmldom.item(cell_nodes, node_idx);
          node_idx := node_idx + 1;
          
          -- read cell index
          cell_idx := nvl(dbms_xslprocessor.valueOf(cell_node, '@ss:Index', XSS_SS_NSMAP), cell_idx);
          cell.cellCol := base26encode(cell_idx);
          
          -- read cell data
          cell.cellData := null;
          data_node := dbms_xslprocessor.selectSingleNode(cell_node, 'Data', XSS_DFLT_NSMAP);
          
          if not dbms_xmldom.isNull(data_node) then
            -- type
            info.cellType := dbms_xslprocessor.valueOf(data_node, '@ss:Type', XSS_SS_NSMAP);
            -- value      
            case info.cellType
            when 'Number' then
              str.strval := dbms_xslprocessor.valueOf(data_node, '.');
              cell.cellData := anydata.ConvertNumber(to_number(replace(str.strval,'.',get_decimal_sep)));
                    
            when 'String' then
              begin
                str.strval := dbms_xslprocessor.valueOf(data_node, '.');
              exception
                when value_error or buffer_too_small then
                  readclob(
                    dbms_xslprocessor.selectNodes(data_node, './/text()')
                  , str.lobval
                  );
              end;
                    
              if str.lobval is not null then
                cell.cellData := anydata.ConvertClob(str.lobval);
              else
                --cell.cellData := anydata.ConvertVarchar2(str.strval);
                if lengthb(str.strval) <= MAX_STRING_SIZE then
                  cell.cellData := anydata.ConvertVarchar2(str.strval);
                else
                  cell.cellData := anydata.ConvertClob(to_clob(str.strval));
                end if;
              end if;
                  
            when 'DateTime' then
              str.strval := dbms_xslprocessor.valueOf(data_node, '.');
              cell.cellData := anydata.ConvertTimestamp(to_timestamp(str.strval,'YYYY-MM-DD"T"HH24:MI:SS.FF9'));
              
            when 'Boolean' then
              str.strval := dbms_xslprocessor.valueOf(data_node, '.');
              cell.cellData := anydata.ConvertVarchar2(case when str.strval = '1' then 'TRUE' else 'FALSE' end);
              
            when 'Error' then
              cell.cellData := anydata.ConvertVarchar2(dbms_xslprocessor.valueOf(data_node, '.'));
                  
            else
              null;
              
            end case;
            
            dbms_xmldom.freeNode(data_node);
          
          end if;
                
          -- need comment?
          if refset.exists(cell.cellCol) then
            if is_opt_set(refSet(cell.cellCol), META_COMMENT) then
              --read_comment(cell_node, cell.sheetIdx, cell.cellCol || cell.cellRow);
              cell.cellNote := get_annotation(cell_node);
            end if;
          end if;
              
          dbms_xmldom.freeNode(cell_node);
            
          if cell.cellData is not null then             
            if refset.exists(cell.cellCol) then
              cells.extend;
              cells(cells.last) := cell;
              empty_row := false;
            end if;
          end if;
            
          cell_idx := cell_idx + 1;
              
        end loop;
            
        dbms_xmldom.freeNodeList(cell_nodes);
        dbms_xmldom.freeNode(row_node);
          
        if not empty_row then
          l_nrows := l_nrows - 1;
        end if;
                   
      end if;
        
      -- no more row to read?
      if row_idx = row_cnt or src_row = l_end_row then 
        --ctx_cache(ctx_id).done := true;
        XSS_openNextSheet(ctx_cache(ctx_id));

        row_cnt := ctx_cache(ctx_id).dom_reader.rlist_size;
        row_idx := ctx_cache(ctx_id).dom_reader.rlist_idx;
        src_row := ctx_cache(ctx_id).src_row;
        cell.sheetIdx := ctx_cache(ctx_id).curr_sheet;

      end if;
      
      exit when ctx_cache(ctx_id).done or l_nrows = 0;

    end loop;
      
    ctx_cache(ctx_id).dom_reader.rlist_idx := row_idx;
    ctx_cache(ctx_id).r_num := row_num;
    ctx_cache(ctx_id).src_row := src_row;
      
    return cells;
    
  end;
  

  function getCells_XDB (
    ctx_id   in binary_integer
  , nrows    in number
  )
  return ExcelTableCellList
  is
    
    l_nrows  pls_integer := nrows;
    c        integer := ctx_cache(ctx_id).xdb_reader.c;
    res      integer;
    info     t_cell_info := ctx_cache(ctx_id).xdb_reader.cell_info;
    str      t_string_rec;
    cell     ExcelTableCell := ExcelTableCell(null, null, null, null, null, null);
    cells    ExcelTableCellList := ExcelTableCellList();
    
    prev_row pls_integer;

  begin
    
    loop
      
      -- checking saved (unprocessed) cell
      if info.cellRef is null then
    
        res := dbms_sql.fetch_rows(c);
        if res = 0 then
          ctx_cache(ctx_id).done := true;
          exit;
        end if;
        
        dbms_sql.column_value(c, 1, info.sheetIdx);
        dbms_sql.column_value(c, 2, info.cellRef);
        dbms_sql.column_value(c, 3, info.cellType);
        dbms_sql.column_value(c, 4, info.cellValue);
            
        info.cellRow := ltrim(info.cellRef, LETTERS);
        info.cellCol := rtrim(info.cellRef, DIGITS);
            
        if info.cellRow != prev_row then
          l_nrows := l_nrows - 1;
          if l_nrows = 0 then
            -- save current cell info
            ctx_cache(ctx_id).xdb_reader.cell_info := info;
            exit;
          end if;
        end if;
        
        str := null;
      
      end if;
      
      cell.sheetIdx := info.sheetIdx;
      cell.cellRow := info.cellRow;
      cell.cellCol := info.cellCol;
      
      if info.cellType is null or info.cellType = 'n' then
            
        cell.cellData := anydata.ConvertNumber(to_number(replace(info.cellValue,'.',get_decimal_sep)));
          
      elsif info.cellType = 's' then
          
        str := get_string(ctx_id, info.cellValue);
        if str.strval is not null then
          cell.cellData := anydata.ConvertVarchar2(str.strval);
        elsif str.lobval is not null then
          cell.cellData := anydata.ConvertClob(str.lobval);
        end if;
          
      elsif info.cellType = 'd' then
          
        cell.cellData := anydata.ConvertTimestamp(to_timestamp(info.cellValue,'YYYY-MM-DD"T"HH24:MI:SS.FF3'));
            
      elsif info.cellType = 'b' then
          
        cell.cellData := anydata.ConvertVarchar2(case when info.cellValue = '1' then 'TRUE' else 'FALSE' end);
                
      else
        cell.cellData := anydata.ConvertVarchar2(info.cellValue);
      end if;
      
      --if info.cellType is null and info.cellValue is null
      
      cells.extend;
      cells(cells.last) := cell;
      prev_row := cell.cellRow;  
      info := null;
      --exit when l_nrows = 0;   
        
    end loop;
    
    return cells;
     
  end;


  procedure setFetchSize (p_nrows in number)
  is
  begin
    fetch_size := p_nrows;
  end;
  
  
  procedure useSheetPattern (p_state in boolean)
  is
  begin
    sheet_pattern_enabled := p_state;
  end;


  procedure tableDescribe (
    rtype    out nocopy anytype
  , p_range  in  varchar2
  , p_cols   in  varchar2
  , p_ff     in  boolean default false
  )
  is
    l_type     anytype;
    l_tdef     QI_definition_t;
  begin
    
    if is_compiled_ctx(p_cols) then
      l_tdef := parse_tdef(p_cols);
    else
      if p_ff then
        l_tdef := QI_parseTable(p_range, p_cols, PARSE_COLUMN + PARSE_POSITION);
      else
        l_tdef := QI_parseTable(p_range, p_cols);
      end if;
    end if;
    
    anytype.begincreate(dbms_types.TYPECODE_OBJECT, l_type);

    for i in 1 .. l_tdef.cols.count loop

      l_type.addAttr(
        l_tdef.cols(i).metadata.aname
      , l_tdef.cols(i).metadata.typecode
      , l_tdef.cols(i).metadata.prec
      , l_tdef.cols(i).metadata.scale
      , l_tdef.cols(i).metadata.len
      , l_tdef.cols(i).metadata.csid
      , l_tdef.cols(i).metadata.csfrm
      );

    end loop;

    l_type.endcreate;

    anytype.begincreate(dbms_types.TYPECODE_TABLE, rtype);

    rtype.setInfo(
      null
    , null
    , null
    , null
    , null
    , l_type
    , dbms_types.TYPECODE_OBJECT
    , 0
    );

    rtype.endcreate();
    
  end;


  function tablePrepare(tf_info in sys.ODCITabFuncInfo)
  return anytype
  is
    r  metadata_t;
  begin
    
    r.typecode := tf_info.rettype.getAttrElemInfo(
                    null
                  , r.prec
                  , r.scale
                  , r.len
                  , r.csid
                  , r.csfrm
                  , r.attr_elt_type
                  , r.aname
                  ) ;
                  
     return r.attr_elt_type;
      
  end;


  procedure tableStart (
    p_file         in  blob
  , p_sheetFilter  in  anydata
  , p_range        in  varchar2
  , p_cols         in  varchar2
  , p_method       in  binary_integer
  , p_ctx_id       out binary_integer
  , p_password     in  varchar2
  )  
  is
    ctx_id   binary_integer;
  begin
    
    set_nls_cache;
    
    -- is the table definition in compiled form?
    if is_compiled_ctx(p_cols) then
      ctx_id := get_compiled_ctx(p_cols);
    else
      
      ctx_id := QI_initContext(p_range, p_cols, p_method);
      openSpreadsheet(p_file, p_password, p_sheetFilter, ctx_id);
      
    end if;
    
    p_ctx_id := ctx_id;
        
  end;


  procedure tableStart (
    p_file       in  clob
  , p_cols       in  varchar2
  , p_skip       in  pls_integer
  , p_line_term  in  varchar2
  , p_field_sep  in  varchar2
  , p_text_qual  in  varchar2
  , p_ctx_id     out binary_integer
  )
  is
    ctx_id   binary_integer;
  begin
    
    set_nls_cache;
    
    -- is the table definition in compiled form?
    if is_compiled_ctx(p_cols) then
      ctx_id := get_compiled_ctx(p_cols);
    else
      
      ctx_id := get_free_ctx();
      ctx_cache(ctx_id).r_num := 0;
      ctx_cache(ctx_id).def_cache := QI_parseTable('A'||to_char(p_skip+1), p_cols, PARSE_COLUMN + PARSE_POSITION);
      openFlatFile(p_file, p_field_sep, p_line_term, p_text_qual, ctx_id);
      
    end if;
    
    p_ctx_id := ctx_id;
    
  end;
  
    
  function getRawCells (
    p_file         in blob
  , p_sheetFilter  in anydata
  , p_cols         in varchar2
  , p_range        in varchar2 default null
  , p_method       in binary_integer default DOM_READ
  , p_password     in varchar2 default null
  )
  return ExcelTableCellList pipelined
  is
    ctx_id  binary_integer;
    cells   ExcelTableCellList;
  begin

    set_nls_cache;
    ctx_id := QI_initContext(p_range, p_cols, p_method, PARSE_SIMPLE);
    openSpreadsheet(p_file, p_password, p_sheetFilter, ctx_id);
    
    while not ctx_cache(ctx_id).done loop
      case ctx_cache(ctx_id).file_type
      when FILE_XLSX then
        case ctx_cache(ctx_id).read_method
        when DOM_READ then
          cells := getCells_DOM(ctx_id, fetch_size);
        when STREAM_READ then
          cells := StAX_iterateContext(ctx_cache(ctx_id).extern_key, fetch_size);
          ctx_cache(ctx_id).done := ( cells is empty );
        end case;
      when FILE_XLSB then
        cells := xutl_xlsb.iterate_context(ctx_cache(ctx_id).extern_key, fetch_size);
        ctx_cache(ctx_id).done := ( cells is null );
      when FILE_XLS then
        cells := xutl_xls.iterate_context(ctx_cache(ctx_id).extern_key, fetch_size);
        ctx_cache(ctx_id).done := ( cells is empty );
      when FILE_ODS then
        case ctx_cache(ctx_id).read_method
        when DOM_READ then
          cells := getCells_ODS(ctx_id, fetch_size);
        when STREAM_READ then
          cells := StAX_iterateContext(ctx_cache(ctx_id).extern_key, fetch_size);
          ctx_cache(ctx_id).done := ( cells is empty );
        end case;
        
      when FILE_XSS then
        cells := getCells_XSS(ctx_id, fetch_size);
        
      end case;
      
      if not(cells is null or cells is empty) then
        for i in 1 .. cells.count loop
          
          if cells(i).cellData is not null then
            case cells(i).cellData.getTypeName() 
            when 'SYS.VARCHAR2' then
              if lengthb(cells(i).cellData.accessVarchar2()) > MAX_STRING_SIZE then
                cells(i).cellData := anydata.ConvertClob(cells(i).cellData.accessVarchar2());
              end if;         
            when 'SYS.CHAR' then
              if lengthb(cells(i).cellData.accessChar()) > MAX_STRING_SIZE then
                cells(i).cellData := anydata.ConvertClob(cells(i).cellData.accessChar());
              else
                -- convert to VARCHAR2
                cells(i).cellData := anydata.ConvertVarchar2(cells(i).cellData.accessChar());
              end if;
            when 'SYS.BINARY_DOUBLE' then
              cells(i).cellData := anydata.ConvertNumber(cells(i).cellData.accessBDouble());
            else
              null;
            end case;
          end if;
          
          cells(i).sheetIdx := ctx_cache(ctx_id).sheets(cells(i).sheetIdx).idx;
          
          pipe row (cells(i));
        end loop;
      end if;
      
    end loop;
    
    tableClose(ctx_id);
    
    return;
  
  end;
  

  procedure tableFetch (
    rtype   in out nocopy anytype
  , ctx_id  in binary_integer
  , nrows   in number
  , rws     out nocopy anydataset
  )
  is
    type t_cell_map is table of ExcelTableCell index by varchar2(3);
    
    type t_datum is record (
      val  anydata
    , str  varchar2(32767)
    , num  number
    , dt   date
    , ts   timestamp
    , lob  clob
    );

    l_nrows         number := least(nrows, fetch_size);
    cells           ExcelTableCellList;    
    cell_map        t_cell_map;
    datum           t_datum;
    r_num           pls_integer;
    extern_key      binary_integer;
    cols            QI_column_list_t;
    col_ref         varchar2(10);
    l_prec          pls_integer;
    l_scale         pls_integer;
    previous_row    integer;
    current_row     integer;
    previous_sheet  pls_integer;
    current_sheet   pls_integer;
    hasNonEmptyCell boolean := false;
    
    function datum_is_null return boolean is
    begin
      if datum.val is not null then
        case datum.val.GetTypeName() 
        when 'SYS.VARCHAR2' then
          return ( datum.val.AccessVarchar2() is null );
        when 'SYS.CHAR' then
          return ( datum.val.AccessChar() is null );
        when 'SYS.NUMBER' then
          return ( datum.val.AccessNumber() is null );
        when 'SYS.BINARY_DOUBLE' then
          return ( datum.val.AccessBDouble() is null );
        when 'SYS.DATE' then
          return ( datum.val.AccessDate() is null );
        when 'SYS.CLOB' then
          return ( datum.val.AccessClob() is null );
        when 'SYS.TIMESTAMP' then
          return ( datum.val.AccessTimestamp() is null );
        end case;
      else
        return true;
      end if;      
    end;
    
    procedure set_datum is
    begin
      if datum.val is not null then
        case datum.val.GetTypeName() 
        when 'SYS.VARCHAR2' then
          datum.str := datum.val.AccessVarchar2();
        when 'SYS.CHAR' then
          datum.str := datum.val.AccessChar();
        when 'SYS.NUMBER' then
          datum.num := datum.val.AccessNumber();
        when 'SYS.BINARY_DOUBLE' then
          datum.num := to_number(datum.val.AccessBDouble());
        when 'SYS.DATE' then
          datum.dt := datum.val.AccessDate();
          datum.ts := cast(datum.dt as timestamp);
        when 'SYS.CLOB' then
          datum.lob := datum.val.AccessClob();
        when 'SYS.TIMESTAMP' then
          datum.ts := datum.val.AccessTimestamp();
          --datum.dt := cast(datum.ts as date);
          datum.dt := cast(datum.ts + numtodsinterval(round(extract(second from datum.ts)) - extract(second from datum.ts), 'second') as date);
        end case;
      end if;
    end;

    procedure setRow is
    begin
      
      r_num := r_num + 1;
    
      rws.addInstance;
      rws.piecewise;

      for i in 1 .. cols.count loop

        col_ref := cols(i).metadata.col_ref;
        datum := null;
        
        /*
        if cols(i).cell_meta = META_COMMENT then
      
          datum.str := get_comment(ctx_id, previous_sheet, col_ref || previous_row);
          
        els
        */
        if cols(i).cell_meta = META_SHEET_INDEX then
        
          datum.num := ctx_cache(ctx_id).sheets(previous_sheet).idx;
          
        elsif cols(i).cell_meta = META_SHEET_NAME then
        
          datum.str := ctx_cache(ctx_id).sheets(previous_sheet).name;
        
        elsif cell_map.exists(col_ref) then
        
          if cols(i).cell_meta = META_COMMENT then
            datum.str := cell_map(col_ref).cellNote;
          else
            
            datum.val := cell_map(col_ref).cellData;
            if cols(i).has_default and datum_is_null() then
              datum.val := cols(i).default_value;
            end if;
            set_datum;
          
          end if;
        
        else
        
          datum.val := cols(i).default_value;
          set_datum;

        end if;
        
        case cols(i).metadata.typecode
        when dbms_types.TYPECODE_VARCHAR2 then
          
          if datum.num is not null then
            datum.str := to_char(datum.num);
          elsif datum.lob is not null then
            datum.str := dbms_lob.substr(datum.lob, LOB_CHUNK_SIZE);
          end if;
          datum.str := substrb(datum.str, 1, cols(i).metadata.len);
          rws.setVarchar2(datum.str);
            
        when dbms_types.TYPECODE_NUMBER then
          
          if cols(i).for_ordinality then
            datum.num := r_num;
          elsif datum.str is not null then           
            datum.num := to_number(replace(datum.str,'.',get_decimal_sep));
          end if;
          l_scale := cols(i).metadata.scale;
          if l_scale is not null then 
            datum.num := round(datum.num, l_scale);
          end if;
          l_prec := cols(i).metadata.prec;
          --if l_prec is not null and log(10, datum.num) >= l_prec-l_scale then
          if l_prec is not null and abs(datum.num) >= cols(i).metadata.max_value then
            raise value_out_of_range;
          end if;
          rws.setNumber(datum.num);
          
        when dbms_types.TYPECODE_DATE then
          
          if datum.num is not null then
            datum.dt := get_date_val(datum.num);
          elsif datum.str is not null then
            datum.dt := get_date_val(datum.str, cols(i).format);
          end if;
          rws.SetDate(datum.dt);
          
        when dbms_types.TYPECODE_TIMESTAMP then
          
          if datum.num is not null then
            datum.ts := get_tstamp_val(datum.num);
          elsif datum.str is not null then
            datum.ts := get_tstamp_val(datum.str, cols(i).format);
          end if;
          rws.SetTimestamp(datum.ts);
            
        when dbms_types.TYPECODE_CLOB then
          
          if datum.str is not null then
            datum.lob := to_clob(datum.str);
          end if;
          rws.SetClob(datum.lob);
          
        end case;
          
      end loop;
      
    end;

  begin
    
    debug('Requested rows = '||to_char(nrows));
   
    if not ctx_cache(ctx_id).done then
       
      cols := ctx_cache(ctx_id).def_cache.cols;
      r_num := ctx_cache(ctx_id).r_num;
      extern_key := ctx_cache(ctx_id).extern_key;
      
      case ctx_cache(ctx_id).file_type
      when FILE_XLSX then
        case ctx_cache(ctx_id).read_method
        when DOM_READ then
          cells := getCells_DOM(ctx_id, l_nrows);
        when STREAM_READ then
          cells := StAX_iterateContext(extern_key, l_nrows);
        when STREAM_READ_XDB then
          cells := getCells_XDB(ctx_id, l_nrows);
        end case;
        
      when FILE_XLS then
        cells := xutl_xls.iterate_context(extern_key, l_nrows);
      when FILE_XLSB then
        cells := xutl_xlsb.iterate_context(extern_key, l_nrows);
      when FILE_ODS then
        case ctx_cache(ctx_id).read_method
        when DOM_READ then
          cells := getCells_ODS(ctx_id, l_nrows);
        when STREAM_READ then
          cells := StAX_iterateContext(extern_key, l_nrows);
        end case;
        
      when FILE_XSS then
        cells := getCells_XSS(ctx_id, l_nrows);
      when FILE_FF then
        cells := xutl_flatfile.iterate_context(extern_key, l_nrows);
      end case;
      
      if cells is not null and cells is not empty then
        
        debug('cells.count='||cells.count);
      
        anydataset.beginCreate(dbms_types.TYPECODE_OBJECT, rtype, rws);
        
        for i in 1 .. cells.count loop

          current_row := cells(i).cellRow;
          current_sheet := cells(i).sheetIdx;
          
          if current_row != previous_row or current_sheet != previous_sheet then
            if hasNonEmptyCell then
              setRow;
            end if;
            cell_map.delete;
          end if;
          
          cell_map(cells(i).cellCol) := cells(i);
          hasNonEmptyCell := ( cells(i).cellData is not null );
          
          previous_row := current_row;
          previous_sheet := current_sheet;
                
        end loop;
        
        if hasNonEmptyCell then
          setRow;
        end if;
        
        rws.endCreate;
          
        ctx_cache(ctx_id).r_num := r_num;
        
      end if;
    
    end if;
    
  end;
  

  procedure tableClose (
    p_ctx_id  in binary_integer
  )
  is
  begin
    
    case ctx_cache(p_ctx_id).file_type
    when FILE_XLSX then
  
      case ctx_cache(p_ctx_id).read_method
      when DOM_READ then
        
        ctx_cache(p_ctx_id).string_cache := t_strings();
        --dbms_xmldom.freeNodeList(ctx_cache(p_ctx_id).dom_reader.rlist);
        --dbms_xmldom.freeDocument(ctx_cache(p_ctx_id).dom_reader.doc); 
        
      when STREAM_READ then
        
        StAX_closeContext(ctx_cache(p_ctx_id).extern_key);
        
      when STREAM_READ_XDB then
        
        ctx_cache(p_ctx_id).string_cache := t_strings();
        XDB_closeReader(p_ctx_id, ctx_cache(p_ctx_id).xdb_reader);
        
      end case;
      
    when FILE_XLS then
      
      xutl_xls.free_context(ctx_cache(p_ctx_id).extern_key);
      
    when FILE_XLSB then
      
      xutl_xlsb.free_context(ctx_cache(p_ctx_id).extern_key);
    
    when FILE_ODS then
      
      case ctx_cache(p_ctx_id).read_method
      when DOM_READ then
        dbms_xmldom.freeDocument(ctx_cache(p_ctx_id).dom_reader.doc); 
        
      when STREAM_READ then
        StAX_closeContext(ctx_cache(p_ctx_id).extern_key);
        
      end case;

    when FILE_XSS then
    
      dbms_xmldom.freeDocument(ctx_cache(p_ctx_id).dom_reader.doc);

    when FILE_FF then
      
      xutl_flatfile.free_context(ctx_cache(p_ctx_id).extern_key);
    
    end case;
    
    ctx_cache.delete(p_ctx_id);
    
  end;


  function getFile (
    p_directory in varchar2
  , p_filename  in varchar2
  )
  return blob
  is
    l_dest_offset  integer := 1;
    l_src_offset   integer := 1;
    l_file         bfile := bfilename(p_directory, p_filename);
    l_blob         blob;
  begin
    dbms_lob.createtemporary(l_blob, true);
    dbms_lob.fileopen(l_file, dbms_lob.file_readonly);
    dbms_lob.loadblobfromfile(
      dest_lob    => l_blob
    , src_bfile   => l_file
    , amount      => dbms_lob.getlength(l_file)
    , dest_offset => l_dest_offset
    , src_offset  => l_src_offset
    );
    dbms_lob.fileclose(l_file);
    return l_blob;
  end;


  function getTextFile (
    p_directory in varchar2
  , p_filename  in varchar2
  , p_charset   in varchar2 default 'CHAR_CS'
  ) 
  return clob
  is
    l_file         bfile := bfilename(p_directory, p_filename);
    l_dest_offset  integer := 1;
    l_src_offset   integer := 1;
    l_lang_ctx     integer := dbms_lob.default_lang_ctx;
    l_warning      integer;
    l_clob         clob;
  begin
    dbms_lob.createtemporary(l_clob, true);
    dbms_lob.fileopen(l_file, dbms_lob.lob_readonly);
    dbms_lob.loadclobfromfile(
      dest_lob     => l_clob
    , src_bfile    => l_file
    , amount       => dbms_lob.getlength(l_file)
    , dest_offset  => l_dest_offset
    , src_offset   => l_src_offset
    , bfile_csid   => nls_charset_id(p_charset)
    , lang_context => l_lang_ctx
    , warning      => l_warning
    );
    dbms_lob.fileclose(l_file);
    return l_clob;
  end;


  function getCursorQuery (
    p_sheetFilter in anydata
  , p_cols        in varchar2
  , p_range       in varchar2 
  )
  return varchar2
  is
    query  varchar2(32767) := 
    'SELECT * FROM TABLE(EXCELTABLE.GETROWS(:1,$$SHEETFILTER,$$COLS,''$$RANGE'',:2,:3))';
    
    sheetFilterType  varchar2(257) := p_sheetFilter.GetTypeName();
    sheetFilterText  varchar2(32767);
    dummy            pls_integer;
    inputSheetList   ExcelTableSheetList;
    
    function to_string_literal (str in varchar2) return varchar2 is
    begin
      return '''' || replace(str, '''', '''''') || '''';
    end;
  
  begin
    
    -- Multi-sheet support : sheetFilter parameter now needs to be hardcoded in order to
    -- avoid error "PLS-00307: too many declarations of 'ODCITABLEDESCRIBE' match this call" 
    case
    when sheetFilterType = 'SYS.VARCHAR2' then
      
      sheetFilterText := to_string_literal(p_sheetFilter.AccessVarchar2());
      
    when sheetFilterType like '%.EXCELTABLESHEETLIST' then
      
      dummy := p_sheetFilter.GetCollection(inputSheetList);
      
      sheetFilterText := 'EXCELTABLESHEETLIST(';
      if inputSheetList is not null then
        for i in 1 .. inputSheetList.count loop
          if i > 1 then
            sheetFilterText := sheetFilterText || ',';
          end if;
          sheetFilterText := sheetFilterText || to_string_literal(inputSheetList(i));
        end loop;
      end if;
      sheetFilterText := sheetFilterText || ')';
      
    end case;
  
    query := replace(query, '$$SHEETFILTER', sheetFilterText);
    query := replace(query, '$$COLS', to_string_literal(p_cols));
    query := replace(query, '$$RANGE', p_range);
    
    debug(query);
    
    return query;
    
  end;
  
  
  function getCursor (
    p_file     in blob
  , p_sheet    in varchar2
  , p_cols     in varchar2
  , p_range    in varchar2 default null
  , p_method   in binary_integer default DOM_READ
  , p_password in varchar2 default null    
  )
  return sys_refcursor
  is
    l_rc     sys_refcursor;
    l_query  varchar2(32767) := getCursorQuery(anydata.ConvertVarchar2(p_sheet), p_cols, p_range);
  begin
    open l_rc for l_query using p_file, p_method, p_password;
    return l_rc;
  end;


  function getCursor (
    p_file     in blob
  , p_sheets   in ExcelTableSheetList
  , p_cols     in varchar2
  , p_range    in varchar2 default null
  , p_method   in binary_integer default DOM_READ
  , p_password in varchar2 default null    
  )
  return sys_refcursor
  is
    l_rc     sys_refcursor;
    l_query  varchar2(32767) := getCursorQuery(anydata.ConvertCollection(p_sheets), p_cols, p_range);
  begin
    open l_rc for l_query using p_file, p_method, p_password;
    return l_rc;
  end;


  function getCursor (
    p_file      in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer
  , p_line_term in varchar2
  , p_field_sep in varchar2 default null
  , p_text_qual in varchar2 default null    
  )
  return sys_refcursor
  is
    l_rc     sys_refcursor;
    l_query  varchar2(32767) := 'SELECT * FROM TABLE(EXCELTABLE.GETROWS(:1,$$COLS,$$SKIP,:2,:3,:4))';

    function to_string_literal (str in varchar2) return varchar2 is
    begin
      return '''' || replace(str, '''', '''''') || '''';
    end;

  begin
    l_query := replace(l_query, '$$COLS', to_string_literal(p_cols));
    l_query := replace(l_query, '$$SKIP', to_char(nvl(p_skip,0)));
    debug(l_query);
    open l_rc for l_query using p_file, p_line_term, p_field_sep, p_text_qual;
    return l_rc;
  end;


  function createDMLStatement (
    ctx_id    in pls_integer
  , dml_type  in pls_integer
  , err_log   in varchar2
  )
  return varchar2
  is
    
    INSERT_STMT  constant varchar2(32767) := 'INSERT INTO $$TABLE ($$LIST) $$QUERY';
    UPDATE_STMT  constant varchar2(32767) := 
    'MERGE INTO $$TABLE t USING ($$QUERY) v ON ($$QPREDLIST) WHEN MATCHED THEN UPDATE SET $$QSETLIST';
    MERGE_STMT   constant varchar2(32767) := UPDATE_STMT ||
    ' WHEN NOT MATCHED THEN INSERT ($$LIST) VALUES ($$QLIST)';
    DELETE_STMT  constant varchar2(32767) := 'DELETE $$TABLE WHERE ($$LIST) IN ($$QUERY)';
       
    --query                varchar2(32767) := 'SELECT $$HINT$$LIST FROM TABLE(EXCELTABLE.GETROWS(NULL,'''',''$$COLS'',''$$RANGE'',NULL,''''))';
    query                varchar2(32767) := 'SELECT $$HINT$$LIST FROM TABLE(EXCELTABLE.GETROWS(NULL,'''',''$$COLS''))';
    
    tdef                 QI_definition_t := ctx_cache(ctx_id).def_cache;
    tab_info             t_table_info := ctx_cache(ctx_id).table_info;
    has_key              boolean := false;

    col_name             varchar2(128);
    quoted_id            varchar2(130);  
    simple_list          varchar2(32767);
    simple_key_list      varchar2(32767);
    qualified_list       varchar2(32767);
    qualified_pred_list  varchar2(32767);
    qualified_set_list   varchar2(32767);
    
    cmp_ctx              varchar2(32767);  
    stmt                 varchar2(32767);
    hint                 varchar2(128);
    
    procedure put (list in out nocopy varchar2, val in varchar2, sep in varchar2 default ',')
    is
    begin
      if list is not null then
        list := list || sep;
      end if;
      list := list || val;
    end;
    
  begin
    
    cmp_ctx := compile_tdef(ctx_id);
    query := replace(query, '$$COLS', cmp_ctx);
    --query := replace(query, '$$RANGE', null);

    case dml_type
    when DML_INSERT then
      stmt := INSERT_STMT;
    when DML_UPDATE then
      stmt := UPDATE_STMT;
    when DML_MERGE then
      stmt := MERGE_STMT;
    when DML_DELETE then
      stmt := DELETE_STMT;
    else
      error(DML_UNKNOWN_TYPE);
    end case;

    col_name := tdef.colSet.first;

    for i in 1 .. tdef.cols.count loop
        
      quoted_id := dbms_assert.enquote_name(tdef.cols(i).metadata.aname, false);
      
      put(simple_list, quoted_id);
       
      if dml_type in (DML_UPDATE, DML_MERGE) then
        if tdef.cols(i).is_key then
          has_key := true;
          put(qualified_pred_list, 't.'||quoted_id||' = v.'||quoted_id, ' AND ');
        else
          put(qualified_set_list, 't.'||quoted_id||' = v.'||quoted_id);
        end if;
        put(qualified_list, 'v.'||quoted_id);
      elsif dml_type = DML_DELETE then
        if tdef.cols(i).is_key then
          has_key := true;
          put(simple_key_list, quoted_id);
        end if;
      end if;
        
      col_name := tdef.colSet.next(col_name);
      
    end loop;
       
    stmt := replace( stmt
                   , '$$TABLE'
                   , dbms_assert.enquote_name(tab_info.schema_name, false) || 
                     '.' || 
                     dbms_assert.enquote_name(tab_info.table_name, false)
                   );
    
    stmt := replace(stmt, '$$QUERY', query);
    
    if dml_type = DML_INSERT then
      
      stmt := replace(stmt, '$$LIST', simple_list);
      
    elsif dml_type in (DML_UPDATE, DML_MERGE, DML_DELETE) then
      
      if not has_key then
        error(DML_NO_KEY);
      end if;
      
      if dml_type in (DML_UPDATE, DML_MERGE) then
        stmt := replace(stmt, '$$LIST', simple_list);
        hint := '/*+ no_merge*/ ';
        stmt := replace(stmt, '$$QPREDLIST', qualified_pred_list);
        stmt := replace(stmt, '$$QSETLIST', qualified_set_list);
        if dml_type = DML_MERGE then
          stmt := replace(stmt, '$$QLIST', qualified_list);
        end if;
      else 
        stmt := replace(stmt, '$$LIST', simple_key_list);
      end if;
      
    end if;
    
    stmt := replace(stmt, '$$HINT', hint);
    
    if err_log is not null then
      stmt := stmt || ' ' || trim(err_log);
    end if;
    
    debug(stmt);
    
    return stmt;
    
  end;


  function createDMLContext (
    p_table_name in varchar2
  )
  return DMLContext
  is
    ctx_id    binary_integer := get_free_ctx();
    tab_info  t_table_info;
  begin
    
    ctx_cache(ctx_id).r_num := 0;
    tab_info := resolve_table(p_table_name);
    
    if tab_info.dblink is not null then
      error('Database link not supported for target table definition');
    end if;
    
    ctx_cache(ctx_id).table_info := tab_info;
    ctx_cache(ctx_id).def_cache.cols := QI_column_list_t();
    
    return ctx_id;
  
  end;
  
  /*
  procedure closeDMLContext (
    p_ctx_id in DMLContext
  )
  is
  begin
    tableClose(p_ctx_id);
  end;
  */

  procedure mapColumn (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in anydata default null
  )
  is
  
    tab_info  t_table_info := ctx_cache(p_ctx).table_info;
  
    cursor c_column_info (
      p_owner       in varchar2
    , p_table_name  in varchar2
    , p_column_name in varchar2
    ) 
    is
    select column_name, data_type, data_length, data_precision, data_scale
    from all_tab_columns 
    where owner = p_owner
    and table_name = p_table_name
    and column_name = p_column_name;
    
    col_info  c_column_info%rowtype;
    col       QI_column_t;
    
    pos       pls_integer;
    
    procedure add_column (cols in out nocopy QI_column_list_t)
    is
    begin
      cols.extend;
      cols(cols.last) := col;
    end;

    function validate_position (item in varchar2) return varchar2 is
    begin
      if item is null then
        error('Missing position reference in item ''%s''', item);
      elsif not regexp_like(item, '^\d+$') then
        error('Invalid position reference ''%s''', item);
      end if;
      return item;
    end;
  
  begin
    
    open c_column_info (tab_info.schema_name, tab_info.table_name, p_col_name);
    fetch c_column_info into col_info;
    close c_column_info;
    
    if col_info.column_name is null then
      error('"%s": invalid identifier', p_col_name);
    end if;
    
    col.metadata.aname := col_info.column_name;
    col.default_value := p_default;
    col.has_default := ( col.default_value is not null );
    
    if p_meta = META_ORDINALITY then
      
      col.metadata.typecode := dbms_types.TYPECODE_NUMBER;
      col.cell_meta := META_VALUE;
      col.for_ordinality := true;
      ctx_cache(p_ctx).def_cache.hasOrdinal := true;
      
    else
    
      case col_info.data_type
      when 'NUMBER' then
        col.metadata.typecode := dbms_types.TYPECODE_NUMBER;
        col.metadata.prec := col_info.data_precision;
        col.metadata.scale := col_info.data_scale;
        col.metadata.max_value := 10**(col.metadata.prec - col.metadata.scale);
      when 'VARCHAR2' then
        col.metadata.typecode := dbms_types.TYPECODE_VARCHAR2;
        col.metadata.len := col_info.data_length;
        col.metadata.csid := DB_CSID;
        col.metadata.csfrm := 1;
      when 'DATE' then
        col.metadata.typecode := dbms_types.TYPECODE_DATE;
        col.format := p_format; 
      when 'CLOB' then
        col.metadata.typecode := dbms_types.TYPECODE_CLOB;
        col.metadata.csid := DB_CSID;
        col.metadata.csfrm := 1;
      else 
        if col_info.data_type like 'TIMESTAMP(_)' then
          col.metadata.typecode := dbms_types.TYPECODE_TIMESTAMP;
          col.format := p_format;
          col.metadata.scale := col_info.data_scale;
        else
          error(UNSUPPORTED_DATATYPE, col_info.data_type);
        end if;
      end case;
      
      if p_col_ref is null then
        col.cell_meta := nvl(p_meta, META_CONSTANT);
      else

        -- column reference (A, B, C, ...) or position reference (start:end)
        pos := instr(p_col_ref, token_map(T_COLON));
        if pos != 0 then
          col.position.start_offset := validate_position(substr(p_col_ref, 1, pos-1));
          col.position.end_offset := validate_position(substr(p_col_ref, pos+1));
          col.is_positional := true;
        else
          col.metadata.col_ref := p_col_ref;
        end if;
        
        col.cell_meta := nvl(p_meta, META_VALUE);
        if col.cell_meta = META_COMMENT then
          ctx_cache(p_ctx).def_cache.hasComment := true;
        end if;
      
      end if;
      
      col.is_key := p_key;
      
    end if;
    
    add_column(ctx_cache(p_ctx).def_cache.cols);
    
  end;


  procedure mapColumnWithDefault (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in varchar2
  )
  is
  begin
    mapColumn(p_ctx, p_col_name, p_col_ref, p_format, p_meta, p_key, anydata.ConvertVarchar2(p_default));
  end;


  procedure mapColumnWithDefault (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in number
  )
  is
  begin
    mapColumn(p_ctx, p_col_name, p_col_ref, p_format, p_meta, p_key, anydata.ConvertNumber(p_default));
  end;
  
  
  procedure mapColumnWithDefault (
    p_ctx      in DMLContext
  , p_col_name in varchar2
  , p_col_ref  in varchar2 default null
  , p_format   in varchar2 default null
  , p_meta     in pls_integer default null
  , p_key      in boolean default false
  , p_default  in date
  )
  is
  begin
    mapColumn(p_ctx, p_col_name, p_col_ref, p_format, p_meta, p_key, anydata.ConvertDate(p_default));
  end;
  

  function loadDataImpl (
    p_ctx        in DMLContext 
  , p_file       in blob
  , sheetFilter  in anydata 
  , p_range      in varchar2
  , p_method     in binary_integer
  , p_password   in varchar2
  , p_dml_type   in pls_integer
  , p_err_log    in varchar2
  )
  return integer
  is
    stmt   varchar2(32767);
    nrows  integer;
  begin

    set_nls_cache;
    
    ctx_cache(p_ctx).read_method := p_method;
    ctx_cache(p_ctx).r_num := 0;
    ctx_cache(p_ctx).curr_sheet := 0;
    ctx_cache(p_ctx).def_cache.range := QI_parseRange(p_range);
    validate_columns(ctx_cache(p_ctx).def_cache);
    openSpreadsheet(p_file, p_password, sheetFilter, p_ctx);
    
    stmt := createDMLStatement(p_ctx, p_dml_type, p_err_log);
    execute immediate stmt;
    nrows := sql%rowcount;
    
    return nrows;
    
  end;


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
  return integer
  is
  begin
    return loadDataImpl(
             p_ctx
           , p_file
           , anydata.ConvertVarchar2(p_sheet) 
           , p_range
           , p_method
           , p_password
           , p_dml_type
           , p_err_log
           );
  end;


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
  return integer
  is
  begin
    return loadDataImpl(
             p_ctx
           , p_file
           , anydata.ConvertCollection(p_sheets)
           , p_range
           , p_method
           , p_password
           , p_dml_type
           , p_err_log
           );
  end;
  
  
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
  return integer
  is
    stmt   varchar2(32767);
    nrows  integer;
  begin

    set_nls_cache;
    
    ctx_cache(p_ctx).r_num := 0;
    ctx_cache(p_ctx).def_cache.range := QI_parseRange('A'||to_char(p_skip+1));
    validate_columns(ctx_cache(p_ctx).def_cache);
    openFlatFile(p_file, p_field_sep, p_line_term, p_text_qual, p_ctx);
    
    stmt := createDMLStatement(p_ctx, p_dml_type, p_err_log);
    execute immediate stmt;
    nrows := sql%rowcount;
    
    return nrows;
    
  end;
  
  function getSheets (
    p_file         in blob
  , p_password     in varchar2 default null
  , p_method       in binary_integer default DOM_READ
  )
  return ExcelTableSheetList pipelined
  is
    ctx_id  binary_integer;
    old_sheet_pattern_enabled constant boolean := sheet_pattern_enabled;

    procedure cleanup
    is
    begin
      sheet_pattern_enabled := old_sheet_pattern_enabled;
      tableClose(ctx_id);      
    end cleanup;
  begin
    -- copied from getRawCells
  
    set_nls_cache;
    -- get just the first cell since that is enough to get all sheets
    ctx_id := QI_initContext(p_range => 'A1:A1', p_cols => 'A', p_method => p_method, p_parse_options => PARSE_SIMPLE);
    sheet_pattern_enabled := true;
    -- get all sheets using the regular expression .*
    openSpreadsheet(p_file, p_password, anydata.ConvertVarchar2('.*'), ctx_id);

    for i in 1 .. ctx_cache(ctx_id).sheets.count loop
      pipe row (ctx_cache(ctx_id).sheets(i).name);
    end loop;
    
    cleanup;
    
    return;
  exception
    when others
    then
      cleanup;
      raise;
  end getSheets;

  function isReadMethodAvailable (
    p_method in binary_integer
  )
  return boolean
  is
    l_result boolean := null;
  begin
    case
      -- STREAM_READ_XDB does not require Java, so the the function must always return true in this case.
      when p_method in (DOM_READ, STREAM_READ_XDB)
      then l_result := true;
      when p_method = STREAM_READ
      then
        -- try to close a null context: should fail due to a null pointer exception
        begin
          StAX_closeContext(null);
          raise program_error; -- should not come here
        exception
          when others
          then
            -- Possible exceptions:
            --
            -- ORA-29532: Java call terminated by uncaught Java exception: java.lang.NullPointerException
            -- ORA-29531: no method terminate in class db/office/spreadsheet/ReadContext
            -- ORA-29540: class db/office/spreadsheet/ReadContext does not exist
            if sqlerrm like '%java.lang.NullPointerException%'
            then
              l_result := true;
            else
              l_result := false;
            end if;
        end;
    end case;

    return l_result;
  end  isReadMethodAvailable;

  begin
    
    init_state();

end ExcelTable;
/
