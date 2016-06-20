create or replace package ExcelTable is
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
    p_file   in  blob
  , p_sheet  in  varchar2
  , p_cols   in  varchar2
  , p_range  in  varchar2 default null
  ) 
  return anydataset pipelined
  using ExcelTableImpl;
    
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
    p_file   in  blob
  , p_sheet  in  varchar2
  , p_range  in  varchar2
  , p_cols   in  varchar2
  , p_doc_id out raw
  , p_ctx_id out raw
  );

  procedure tableFetch(
    p_type   in out nocopy anytype
  , p_ctx_id in out nocopy raw
  , p_rnum   in out nocopy integer
  , p_done   in out nocopy integer
  , nrows    in number
  , rws      out nocopy anydataset
  );
  
  procedure tableClose(
    p_doc_id  in raw
  , p_ctx_id  in raw
  );
  
  function getFile (
    p_directory in varchar2
  , p_filename  in varchar2
  ) 
  return blob;
  
end ExcelTable;
/
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
  
  QUOTED_IDENTIFIER      constant binary_integer := 1;
  DIGITS                 constant varchar2(10) := '0123456789';
  LETTERS                constant varchar2(26) := 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  
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
  FOR_ORDINALITY_CLAUSE  constant varchar2(100) := 'At most one "for ordinality" clause is allowed';
  MIXED_COLUMN_DEF       constant varchar2(100) := 'Cannot mix positional and named column definitions';
  EMPTY_COL_REF          constant varchar2(100) := 'Missing column reference for "%s"';
  INVALID_COL_REF        constant varchar2(100) := 'Invalid column reference ''%s''';
  INVALID_COL            constant varchar2(100) := 'Column out of range ''%s''';

  -- OOX Constants
  RS_OFFICEDOC           constant varchar2(100) := 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
  --CT_STYLES              constant varchar2(100) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
  --CT_WORKBOOK            constant varchar2(100) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml';
  --CT_WORKBOOK_ME         constant varchar2(100) := 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';
  CT_SHAREDSTRINGS       constant varchar2(100) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml';
  --CT_WORKSHEET           constant varchar2(100) := 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';
  SML_NSMAP              constant varchar2(100) := 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
  	
  -- DB Constants
  DB_CSID                constant pls_integer := nls_charset_id('CHAR_CS');
  DB_CHARSET             constant varchar2(30) := nls_charset_name(DB_CSID);
  MAX_CHAR_SIZE          constant pls_integer := utl_i18n.get_max_character_size(DB_CHARSET);
  LOB_CHUNK_SIZE         constant pls_integer := trunc(32767 / MAX_CHAR_SIZE);
  MAX_STRING_SIZE        pls_integer;

  -- Internal structure definitions
  type metadata_t is record (
    typecode       pls_integer
  , prec           pls_integer
  , scale          pls_integer
  , len            pls_integer
  , csid           pls_integer
  , csfrm          pls_integer
  , attr_elt_type  anytype
  , aname          varchar2(30)
  , schema_name    varchar2(30)
  , type_name      varchar2(30)
  , version        varchar2(30)
  , numelems       pls_integer
  -- extra fields
  , len_in_char    pls_integer
  , col_ref        varchar2(3)
  );
  
  type QI_cell_ref_t is record (c varchar2(3), cn pls_integer, r pls_integer); 
  type QI_range_t is record (start_ref QI_cell_ref_t, end_ref QI_cell_ref_t);
  type QI_column_t is record (metadata metadata_t, format varchar2(30), for_ordinality boolean default false);
  type QI_column_list_t is table of QI_column_t;
  type QI_definition_t is record (range QI_range_t, cols QI_column_list_t);

  type token_map_t is table of varchar2(30) index by binary_integer;
  type token_t is record (type binary_integer, strval varchar2(4000), intval binary_integer, pos binary_integer);
  type tokenizer_t is record (expr varchar2(4000), pos binary_integer, options binary_integer);

  -- OOX
  type t_offsets is table of integer index by varchar2(260);
  type t_archive is record (offsets t_offsets, content blob);
  type t_workbook is record (path varchar2(260), content xmltype, rels xmltype);
  type t_exceldoc is record (file t_archive, content_map xmltype, workbook t_workbook);
  -- string cache
  type t_string_rec is record (strval varchar2(32767), lobval clob /*, len integer*/);
  type t_strings is table of t_string_rec;

  token_map         token_map_t;
  tokenizer         tokenizer_t;
  string_cache      t_strings;
  def_cache         QI_definition_t;
  nls_date_format   varchar2(64);
  nls_numeric_char  varchar2(2);
  
   
  -- ----------------------------------------------------------------------------------------------
  -- Open a zip archive and read entries from central directory segment
  -- References :
  -- https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT
  -- https://en.wikipedia.org/wiki/Zip_%28file_format%29
  function Zip_openArchive (p_zip in blob)
  return t_archive
  is

    ecds       binary_integer; -- End of central directory signature
    oscd       binary_integer; -- Offset of start of central directory, relative to start of archive
    tncdr      binary_integer; -- Total number of central directory records
    fnl        binary_integer; -- File name length
    efl        binary_integer; -- Extra field length
    fcl        binary_integer; -- File comment length
    fn         varchar2(260);  -- File name
    lfh        binary_integer; -- Local file header
    gpb        raw(2);         -- General Purpose Bits
    enc        varchar2(30);
    cdrPtr     binary_integer := 0;
    my_archive t_archive;

  begin

    ecds := dbms_lob.instr(p_zip, hextoraw('504B0506'));
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
        lfh := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_zip, 4, cdrPtr+42), utl_raw.little_endian) + 1;
        my_archive.offsets(fn) := lfh;
      end if;
      -- next entry in central directory
      cdrPtr := cdrPtr + 46 + fnl + efl + fcl;

    end loop;

    my_archive.content := p_zip;

    return my_archive;

  end;

  -- ----------------------------------------------------------------------------------------------
  -- Get a zip entry by its name
  -- 
  function Zip_getEntry (
    p_archive   in out nocopy t_archive
  , p_entryname in varchar2
  )
  return blob
  is
    tmp        blob := hextoraw('1F8B08000000000000FF'); -- gzip magic header + flags
    entry      blob;
    fnl        binary_integer; -- File name length
    efl        binary_integer; -- Extra field length
    csz        binary_integer; -- Compressed size
    lfh        binary_integer; -- Local file header
  begin
    lfh := p_archive.offsets(p_entryname); -- local file header
    fnl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_archive.content, 2, lfh+26), utl_raw.little_endian);
    efl := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_archive.content, 2, lfh+28), utl_raw.little_endian);
    csz := utl_raw.cast_to_binary_integer(dbms_lob.substr(p_archive.content, 4, lfh+18), utl_raw.little_endian);
    
    dbms_lob.copy(tmp, p_archive.content, csz, 11, lfh + 30 + fnl + efl);
    dbms_lob.append(tmp, dbms_lob.substr(p_archive.content, 4, lfh + 14)); -- CRC32
    dbms_lob.append(tmp, dbms_lob.substr(p_archive.content, 4, lfh + 22)); -- uncompressed size
    
    entry := utl_compress.lz_uncompress(tmp);
    return entry;
  end;
  
  -- ----------------------------------------------------------------------------------------------
  -- Get a zip entry as XMLType
  --  assuming the part has been encoded in UTF-8, as Excel does natively
  function Zip_getXML (
    p_archive   in out nocopy t_archive
  , p_partname  in varchar2
  )
  return xmltype
  is
  begin
    return xmltype(Zip_getEntry(p_archive, p_partname), nls_charset_id('AL32UTF8'));
  end;
  
  -- convert a base26-encoded number to decimal
  function base26decode (p_str in varchar2) 
  return pls_integer 
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
  
  
  -- SELECT_CATALOG_ROLE required
  -- grant select on sys.v_$parameter to <user>;
  function get_max_string_size 
  return pls_integer 
  is
    l_result  pls_integer;
  begin
    select case when value = 'EXTENDED' then 32767 else 4000 end
    into l_result
    from v$parameter
    where name = 'max_string_size' ;
    return l_result;
  exception
    when no_data_found then
      return 4000 ;
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
  end;


  function get_date_format 
  return varchar2 
  is
  begin
    return nls_date_format;  
  end;

 
  function get_decimal_sep 
  return varchar2 
  is
  begin
    return substr(nls_numeric_char, 1, 1);  
  end;

  
  function get_string_val (p_idx in binary_integer) 
  return varchar2 
  is
    rec  t_string_rec := string_cache(p_idx+1);
  begin
    if rec.strval is not null then
      return rec.strval;
    else
      return dbms_lob.substr(rec.lobval, LOB_CHUNK_SIZE);
    end if;
  end;


  function get_clob_val (p_idx in binary_integer) 
  return clob is
    rec  t_string_rec := string_cache(p_idx+1);
  begin
    if rec.strval is not null then
      return to_clob(rec.strval);
    else
      return rec.lobval;
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
      dbms_lob.createtemporary(p_content, false);
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
      buf := substrc(tmp,1);
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
  
    
  procedure init is
  begin
    MAX_STRING_SIZE := get_max_string_size();  
  
	  token_map(T_NAME)   := '<name>';
	  token_map(T_INT)    := '<integer>';
	  token_map(T_IDENT)  := '<identifier>';
    token_map(T_STRING) := '<string literal>';
    token_map(T_EOF)    := '<eof>';
    token_map(T_COMMA)  := ',';
    token_map(T_LEFT)   := '(';
    token_map(T_RIGHT)  := ')';
	end;

  
  procedure error (
    p_message in varchar2
  , p_arg1 in varchar2 default null
  , p_arg2 in varchar2 default null
  , p_arg3 in varchar2 default null
  ) 
  is
  begin
    raise_application_error(-20000, utl_lms.format_message(p_message, p_arg1, p_arg2, p_arg3));
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
          --or c = '_'
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
    
    --dbms_output.put_line('['||to_char(token.pos,'fm099')||'] '||to_char(token.type,'99')||' '||token.strval);
    
    return token;
  
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
    if l_coln not between nvl(p_range.start_ref.cn, 1) and nvl(p_range.end_ref.cn, 16384) then
      error(INVALID_COL, p_col_ref);
    end if;   
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
      if l_coln > 16384 then
        error(RANGE_INVALID_COL, l_col);
      end if;
      l_rnum := to_number(l_row);
      if l_rnum not between 1 and 1048576 then
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
    p_range in varchar2
  , p_cols in varchar2
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
    start_col            pls_integer;
    end_col              pls_integer;
    pos                  pls_integer;
    for_ordinality_check boolean := false;
    col_cnt              pls_integer;
    col_ref_cnt          pls_integer := 0;

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
      elsif lengthb(strval) > 30 then
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
        
      elsif accept(T_NAME, 'CLOB') then
        col.metadata.typecode := dbms_types.TYPECODE_CLOB;
        col.metadata.csid := DB_CSID;
        col.metadata.csfrm := 1;
                
      elsif accept(T_NAME, 'FOR') then
        expect(T_NAME, 'ORDINALITY');
        col.metadata.typecode := dbms_types.TYPECODE_NUMBER;
        col.for_ordinality := true;
        
      else
        error(UNSUPPORTED_DATATYPE, token.strval);
      end if;
      
      pos := token.pos;
      strval := token.strval;
      if accept(T_NAME, 'COLUMN') then
        if col.for_ordinality then
          error(UNEXPECTED_SYMBOL, pos, strval);
        end if;
        strval := token.strval;
        expect(T_STRING);
        col.metadata.col_ref := strval;
        col_ref_cnt := col_ref_cnt + 1;
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

    tokenizer.expr := p_cols;
    tokenizer.pos := 0;
    tokenizer.options := QUOTED_IDENTIFIER;

    token := next_token();
    table_expr;
    expect(T_EOF);
    
    col_cnt := tdef.cols.count;
    start_col := nvl(tdef.range.start_ref.cn, 1);
    end_col := nvl(tdef.range.end_ref.cn, col_cnt);
    pos := 0;
    for i in 1 .. col_cnt loop
      if tdef.cols(i).for_ordinality then
        if for_ordinality_check then
          error(FOR_ORDINALITY_CLAUSE);
        else
          for_ordinality_check := true;
        end if;
      elsif tdef.cols(i).metadata.col_ref is not null then
        validate_column(tdef.cols(i).metadata.col_ref, tdef.cols(i).metadata.aname, tdef.range);
      elsif pos < end_col and tdef.cols(i).metadata.col_ref is null then 
        tdef.cols(i).metadata.col_ref := base26encode(start_col + pos);
        pos := pos + 1;
      end if;
    end loop;
    
    -- check for mixed column definitions
    if col_ref_cnt != 0 then
      -- skip for_ordinality clause (if exists)
      if for_ordinality_check then
        col_cnt := col_cnt - 1;
      end if;
      if col_ref_cnt != col_cnt then
        error(MIXED_COLUMN_DEF);
      end if;
    end if;
    
    return tdef;
   
  end;
  

  procedure QI_parseTable (
    p_range in varchar2
  , p_cols  in varchar2
  )
  is
  begin
    def_cache := QI_parseTable(p_range, p_cols);
  end;


  function OX_getPathByType (
    p_doc          in out nocopy t_exceldoc
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


  function OX_getWorkbookPath (p_doc in out nocopy t_exceldoc)
  return varchar2
  is  
    l_path   varchar2(260);
    l_rels   xmltype := Zip_getXML(p_doc.file, '_rels/.rels');
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


  function OX_getPathBySheetName (
    p_doc       in out nocopy t_exceldoc
  , p_sheetname in varchar2
  )
  return varchar2
  is
    l_path   varchar2(260);
    l_rid    varchar2(30);
  begin

    select x.rid
    into l_rid
    from xmltable(
           xmlnamespaces(
             default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
           , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships' as "r"
           )
         , '/workbook/sheets/sheet[@name=$wsName]'
           passing p_doc.workbook.content
                 , p_sheetname as "wsName"
           columns rid varchar2(30) path '@r:id'
         ) x ;
   
    select x.partname
    into l_path
    from xmltable(
           xmlnamespaces(default 'http://schemas.openxmlformats.org/package/2006/relationships')
         , 'for $r in /Relationships/Relationship
            where $r/@Id = $rid
            return resolve-uri($r/@Target, $path)'
           passing p_doc.workbook.rels
                 , l_rid as "rid"
                 , p_doc.workbook.path as "path"
           columns partname varchar2(256) path '.'
         ) x ;
         
    return l_path;

  end;


  procedure OX_loadStringCache (p_doc in out nocopy t_exceldoc) 
  is
    l_path     varchar2(260) := OX_getPathByType(p_doc, CT_SHAREDSTRINGS);
    --l_strings  xmltype;
    l_query    varchar2(2000) := q'~
select /*+ no_xml_query_rewrite */ 
       x.strval
     , x.lobval
from xmltable(
       xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
     , '/sst/si'
       passing :1
       columns strval  varchar2(4000) path '.[string-length() le $$1]'
             , lobval  clob           path '.[string-length() gt $$1]'
     ) x
~';

  begin
    
    if l_path is not null then
      
      execute immediate replace(l_query, '$$1', trunc(MAX_STRING_SIZE / MAX_CHAR_SIZE))
      bulk collect into string_cache
      using Zip_getXML(p_doc.file, l_path) ;
      
      --l_strings := Zip_getXML(p_doc.file, l_path);
    
/*      select \*+ no_xml_query_rewrite *\ 
             x.strval
           , x.lobval
           , x.len
      bulk collect into string_cache
      from xmltable(
             xmlnamespaces(default 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
           , '/sst/si'
             passing l_strings
             columns strval  varchar2(4000) path '.[string-length() le 1000]'
                   , lobval  clob           path '.[string-length() gt 1000]'
                   , len     number         path 'string-length(.)'
           ) x ;*/
    
    end if;
    
  end;


  procedure OX_openWorkbook (
    p_doc   in out nocopy t_exceldoc
  , p_file  in  blob
  ) 
  is
  begin
    p_doc.file := Zip_openArchive(p_file);
    p_doc.content_map := Zip_getXML(p_doc.file, '[Content_Types].xml');
    p_doc.workbook.path := OX_getWorkbookPath(p_doc); --OX_getPathByType(p_doc, CT_WORKBOOK);
    p_doc.workbook.content := Zip_getXML(p_doc.file, p_doc.workbook.path);
    p_doc.workbook.rels := Zip_getXML(p_doc.file, regexp_replace(p_doc.workbook.path, '(.*)/(.*)$', '\1/_rels/\2.rels'));
  end;


  function OX_openWorksheet (
    p_file   in  blob
  , p_sheet  in  varchar2  
  ) 
  return raw 
  is
    l_xldoc  t_exceldoc;
    l_doc    dbms_xmldom.DOMDocument;
    l_sheet  xmltype;
  begin
    OX_openWorkbook(l_xldoc, p_file);
    OX_loadStringCache (l_xldoc);
    l_sheet := Zip_getXML(l_xldoc.file, OX_getPathBySheetName(l_xldoc, p_sheet));
    l_doc := dbms_xmldom.newDOMDocument(l_sheet);
    return l_doc.id;
  end;


  procedure tableDescribe (
    rtype    out nocopy anytype
  , p_range  in  varchar2
  , p_cols   in  varchar2
  )
  is
    l_type  anytype;
    l_tdef  QI_definition_t;
  begin
    
    l_tdef := QI_parseTable(p_range, p_cols);
    
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
    p_file   in  blob
  , p_sheet  in  varchar2
  , p_range  in  varchar2
  , p_cols   in  varchar2
  , p_doc_id out raw
  , p_ctx_id out raw
  )  
  is
    l_doc    dbms_xmldom.DOMDocument;
    l_nlist  dbms_xmldom.DOMNodeList;
    l_xpath  varchar2(2000) := '/worksheet/sheetData/row';
    l_range  QI_range_t;
  begin
    
    QI_parseTable(p_range, p_cols);
    set_nls_cache;
    l_range := def_cache.range;

    l_doc.id := OX_openWorksheet(p_file, p_sheet);
  
    if l_range.start_ref.r is not null then
      l_xpath := l_xpath || '[@r>=' || l_range.start_ref.r || ']';
    end if;
    if l_range.end_ref.r is not null then
      l_xpath := l_xpath || '[@r<=' || l_range.end_ref.r || ']';
    end if;
    
    l_nlist := dbms_xslprocessor.selectNodes(dbms_xmldom.makeNode(l_doc), l_xpath, SML_NSMAP);
    
    p_doc_id := l_doc.id;
    p_ctx_id := l_nlist.id;
    
  end;


  procedure tableFetch (
    p_type   in out nocopy anytype
  , p_ctx_id in out nocopy raw
  , p_rnum   in out nocopy integer
  , p_done   in out nocopy integer
  , nrows    in number
  , rws      out nocopy anydataset
  )
  is
  
    value_out_of_range  exception;
    pragma exception_init (value_out_of_range, -1438);
  
    type node_map_t is table of dbms_xmldom.DOMNode index by varchar2(3);
    cells       node_map_t;
  
    l_nrows     integer := 0;
    l_cols      QI_column_list_t;
    l_col       varchar2(10);
    
    l_varchar2  varchar2(32767);
    l_number    number;
    l_date      date;
    l_clob      clob;
    l_val       varchar2(4000);
    l_type      varchar2(10);
    l_format    varchar2(30);
    
    l_prec      pls_integer;
    l_scale     pls_integer;
    
    ds_open     boolean := false;
    
    l_rnode     dbms_xmldom.DOMNode;      
    l_n         dbms_xmldom.DOMNode;
    l_rlist     dbms_xmldom.DOMNodeList;
    l_nlist     dbms_xmldom.DOMNodeList;

  begin
    
    l_rlist.id := p_ctx_id;
    
    if dbms_xmldom.isNull(l_rlist) then
      p_done := 1;
    end if;

    if p_done = 0 then

      l_cols := def_cache.cols;
      
      loop
          
        if not ds_open then
          anydataset.beginCreate(dbms_types.TYPECODE_OBJECT, p_type, rws);
          ds_open := true;
        end if;
        
        rws.addInstance;
        rws.piecewise;

        l_rnode := dbms_xmldom.item(l_rlist, p_rnum);
        l_nlist := dbms_xslprocessor.selectNodes(l_rnode, 'c', SML_NSMAP);
        
        cells.delete;
        for i in 0 .. dbms_xmldom.getLength(l_nlist) - 1 loop
          l_n := dbms_xmldom.item(l_nlist, i);
          l_col := rtrim(dbms_xslprocessor.valueOf(l_n, '@r'), '0123456789');
          cells(l_col) := l_n;
        end loop;
        l_n := null;
        
        dbms_xmldom.freeNodeList(l_nlist);
        dbms_xmldom.freeNode(l_rnode);
        
        for i in 1 .. l_cols.count loop

          l_val := null;
          l_type := null;
          l_clob := null;

          l_col := l_cols(i).metadata.col_ref;
          if cells.exists(l_col) then
          
            l_n := cells(l_col);  
          
            begin
              l_val := dbms_xslprocessor.valueOf(l_n, 'v', SML_NSMAP);
            exception
              when value_error then
                readclob(dbms_xslprocessor.selectSingleNode(l_n, 'v/text()', SML_NSMAP), l_clob);
            end;
            
            l_type := dbms_xslprocessor.valueOf(l_n, '@t');
              
          end if;
          
          case l_cols(i).metadata.typecode
          when dbms_types.TYPECODE_VARCHAR2 then
            if l_type = 's' then
              l_varchar2 := get_string_val(l_val);
            elsif l_type = 'inlineStr' then
              
              begin
                l_varchar2 := dbms_xslprocessor.valueOf(l_n, 'is', SML_NSMAP);
              exception
                when value_error then
                  readclob(dbms_xslprocessor.selectNodes(l_n, 'is//t/text()', SML_NSMAP), l_clob);
                  l_varchar2 := dbms_lob.substr(l_clob, LOB_CHUNK_SIZE);
              end;
              
            elsif l_clob is not null then 
              l_varchar2 := dbms_lob.substr(l_clob, LOB_CHUNK_SIZE);
            else
              l_varchar2 := l_val;
            end if;
            l_varchar2 := substrb(l_varchar2, 1, l_cols(i).metadata.len);
            rws.setVarchar2(l_varchar2);
            
          when dbms_types.TYPECODE_NUMBER then
            if l_cols(i).for_ordinality then
              l_number := p_rnum + 1;
            elsif l_type = 's' then
              l_number := to_number(replace(get_string_val(l_val),'.',get_decimal_sep));
            else
              l_number := to_number(replace(l_val,'.',get_decimal_sep));
            end if;
            l_scale := l_cols(i).metadata.scale;
            if l_scale is not null then 
              l_number := round(l_number, l_scale);
            end if;
            l_prec := l_cols(i).metadata.prec;
            if l_prec is not null and log(10, l_number) >= l_prec-l_scale then
              raise value_out_of_range;
            end if;
            rws.setNumber(l_number);
            
          when dbms_types.TYPECODE_DATE then
            if l_type = 's' then
              l_format := nvl(l_cols(i).format, get_date_format);
              l_date := to_date(get_string_val(l_val), l_format);               
            else
              l_number := to_number(l_val);
              -- Excel bug workaround : date 1900-02-29 doesn't exist yet Excel stores it at serial #60
              -- The following skips it and converts to Oracle date correctly
              if l_number > 60 then
                l_date := date '1899-12-30' + l_number;
              elsif l_number < 60 then
                l_date := date '1899-12-31' + l_number;
              else
                l_date := null;
              end if;
            end if;
            rws.SetDate(l_date);
            
          when dbms_types.TYPECODE_CLOB then
            if l_type = 's' then
              l_clob := get_clob_val(l_val);
            elsif l_type = 'inlineStr' then
              readclob(dbms_xslprocessor.selectNodes(l_n, 'is//t/text()', SML_NSMAP), l_clob);
            elsif l_clob is null then
              l_clob := to_clob(l_val);
            end if;
            rws.SetClob(l_clob);
          
          end case;          
          
          dbms_xmldom.freeNode(l_n);
          
        end loop;
                
        l_nrows := l_nrows + 1;
        p_rnum := p_rnum + 1;

        if p_rnum = dbms_xmldom.getLength(l_rlist) then 
          p_done := 1;
          exit;
        elsif l_nrows = nrows then
          exit;
        end if;

      end loop;

      if ds_open then
        rws.endCreate;
      end if;
       
    end if;
     
  end;
  
  
  procedure tableClose(
    p_doc_id  in raw
  , p_ctx_id  in raw
  )
  is
    l_doc   dbms_xmldom.DOMDocument;
    l_nlist dbms_xmldom.DOMNodeList;
  begin
    l_doc.id := p_doc_id;
    l_nlist.id := p_ctx_id;
    dbms_xmldom.freeNodeList(l_nlist);
    dbms_xmldom.freeDocument(l_doc);
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
    dbms_lob.createtemporary(l_blob, false);
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


  /*
  procedure parse_test (p_expr in varchar2) 
  is
    token token_t := null;
  begin
    tokenizer.expr := p_expr;
    tokenizer.pos := 0;
    tokenizer.options := QUOTED_IDENTIFIER;
    loop
    token := next_token();
    if token.type = -1 then
      exit;
    end if;
    dbms_output.put_line(token.type || ' ' || token.strval);
    end loop;
  end;
  */


  begin
   
    init();

end ExcelTable;
/
