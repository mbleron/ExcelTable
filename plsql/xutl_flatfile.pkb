create or replace package body xutl_flatfile is

  BLOCK_SIZE constant pls_integer := 32767;

  type column_map_t is table of varchar2(3) index by pls_integer;

  type posn_field_t is record (
    id                  pls_integer
  , start_pos           pls_integer
  , end_pos             pls_integer
  , sz                  pls_integer
  , block_start         pls_integer
  , block_end           pls_integer
  , block_start_offset  pls_integer
  , block_end_offset    pls_integer
  , is_single_block     boolean
  );
  
  type posn_field_list_t is table of posn_field_t;

  type file_descriptor_t is record (
    field_separator  varchar2(1)
  , line_terminator  varchar2(2)
  , text_qualifier   varchar2(1)
  , fields           posn_field_list_t
  );

  type field_t is record (
    id            pls_integer
  , str_value     varchar2(32767)
  , lob_value     clob
  , is_lob        boolean := false
  , sz            pls_integer := 0
  , start_offset  pls_integer
  , end_offset    pls_integer
  );

  type block_t is record (
    content  varchar2(32767)
  , sz       pls_integer := 0
  , free     pls_integer
  );
    
  type block_list_t is table of block_t;

  type buffer_t is record (
    content   varchar2(32767)
  , sz        pls_integer
  , offset    pls_integer
  , available pls_integer := 0
  );

  type stream_t is record (
    content   clob
  , sz        integer
  , offset    integer
  , available integer
  );

  type context_t is record (
    stream   stream_t
  , buf      buffer_t
  , fd       file_descriptor_t
  , colmap   column_map_t
  , r_num    integer := 0
  , done     boolean := false
  , ctype    pls_integer
  , skip     pls_integer
  );
  
  type context_cache_t is table of context_t index by pls_integer;
  ctx_cache  context_cache_t;

  function base26encode (colNum in pls_integer) 
  return varchar2
  is
    output  varchar2(3);
    num     pls_integer := colNum;
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
  is
  begin
    return ascii(substr(colRef,-1,1))-64 
         + nvl((ascii(substr(colRef,-2,1))-64)*26, 0)
         + nvl((ascii(substr(colRef,-3,1))-64)*676, 0);
  end;

  procedure error (
    errcode in pls_integer
  , message in varchar2
  , arg1    in varchar2 default null
  , arg2    in varchar2 default null
  )
  is
  begin
    raise_application_error(errcode, utl_lms.format_message(message, arg1, arg2));
  end;

  function parseColumnList (
    cols  in varchar2
  , sep   in varchar2 default ','
  )
  return column_map_t
  is
    colMap  column_map_t;
    token   varchar2(30);
    p1      pls_integer := 1;
    p2      pls_integer;
    p3      pls_integer;
    
    c1      pls_integer;
    c2      pls_integer;
    
    function validate (item in varchar2, pos in pls_integer) return varchar2 is
    begin
      if item is null then
        error(-20741, 'Missing column reference at position %d', pos);
      elsif not regexp_like(item, '^[A-Z]+$') then
        error(-20741, 'Invalid column reference ''%s'' at position %d', item, pos);
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
          error(-20741, 'Invalid range expression at position %d', p1);
        end if;
        for i in c1 .. c2 loop
          colMap(i) := base26encode(i);
        end loop;
      else
        token := validate(token, p1);
        colMap(base26decode(token)) := token;
      end if;
      
      exit when p2 = 0;
      p1 := p2 + 1;
      
    end loop;
    
    return colMap;
  
  end;

  function parseFieldList (
    cols  in varchar2
  , sep   in varchar2 default ','
  )
  return posn_field_list_t
  is
    fields  posn_field_list_t := posn_field_list_t();
    token   varchar2(30);
    p1      pls_integer := 1;
    p2      pls_integer;
    p3      pls_integer;
    
    c1      pls_integer;
    c2      pls_integer;
    
    i       pls_integer := 0;
    
    function validate (item in varchar2, pos in pls_integer) return varchar2 is
    begin
      if item is null then
        error(-20741, 'Missing position reference at position %d', pos);
      elsif not regexp_like(item, '^\d+$') then
        error(-20741, 'Invalid position reference ''%s'' at position %d', item, pos);
      end if;
      return item;
    end;

    procedure add_field(p_start in pls_integer, p_end in pls_integer) is
      f  posn_field_t;
    begin
      i := i + 1;
      
      f.id := i;
      f.start_pos := p_start;
      f.end_pos := p_end;
      f.sz := p_end - p_start + 1;
      f.block_start := trunc((f.start_pos - 1)/BLOCK_SIZE) + 1;
      f.block_end := trunc((f.end_pos - 1)/BLOCK_SIZE) + 1;
      f.block_start_offset := mod(f.start_pos-1, BLOCK_SIZE)+1;
      f.block_end_offset := f.sz - (f.block_end - f.block_start)*BLOCK_SIZE + f.block_start_offset - 1;    
      f.is_single_block := (f.block_start = f.block_end);
      
      fields.extend;
      fields(i) := f;    
    end;
    
  begin

    loop
      
      p2 := instr(cols, sep, p1);
      if p2 = 0 then
        token := substr(cols, p1);
      else
        token := substr(cols, p1, p2-p1);
      end if;
      
      p3 := instr(token, ':');
      if p3 != 0 then
        c1 := validate(substr(token, 1, p3-1), p1);
        c2 := validate(substr(token, p3+1), p1+p3);
        if c2 < c1 then
          error(-20741, 'Invalid field definition at position %d', p1);
        end if;
        add_field(c1, c2);
      else
        error(-20741, 'Invalid field definition at position %d', p1);
      end if;
      
      exit when p2 = 0;
      p1 := p2 + 1;
      
    end loop;
    
    return fields;
  
  end;

  procedure bufferize (
    stream  in out nocopy stream_t
  , buf     in out nocopy buffer_t
  )
  is
    amount  pls_integer := 384;
  begin
    dbms_lob.read(stream.content, amount, stream.offset, buf.content);
    stream.offset := stream.offset + amount;
    stream.available := stream.available - amount;
    buf.sz := amount;
    buf.offset := 1;
  end;
  
  function read_fields (
    stream in out nocopy stream_t
  , buf    in out nocopy buffer_t
  , ctx_id in pls_integer
  , nrows  in pls_integer
  )
  return ExcelTableCellList
  is
  
    fd      file_descriptor_t := ctx_cache(ctx_id).fd;
    skip    pls_integer := ctx_cache(ctx_id).skip;
    colmap  column_map_t := ctx_cache(ctx_id).colmap;
  
    LT_SZ   constant pls_integer := length(fd.line_terminator);
    FS_SZ   constant pls_integer := length(fd.field_separator);
  
    fields            ExcelTableCellList := ExcelTableCellList();
    field             field_t;
    
    r_num             pls_integer := ctx_cache(ctx_id).r_num;
    r_cnt             pls_integer := 0;
    fs_offset         pls_integer; -- field separator offset
    lt_offset         pls_integer; -- line terminator offset
    tq_offset         pls_integer; -- text qualifier offset
    field_chunk_sz    pls_integer;
    lt_chunk          varchar2(2);
    lt_chunk_sz       pls_integer;
    lt_chunk_tail_sz  pls_integer;
    
    USE_TEXT_QUALIFIER  constant boolean := (fd.text_qualifier is not null);

    procedure read_field is
    begin
      field_chunk_sz := field.end_offset - field.start_offset + 1;
      if field.sz != 0 then
        field.str_value := field.str_value || substr(buf.content, field.start_offset, field_chunk_sz);
        field.sz := field.sz + field_chunk_sz;
      else
        field.str_value := substr(buf.content, field.start_offset, field_chunk_sz);
        field.sz := field_chunk_sz;
      end if;    
    end;

    procedure add_field is
    begin
      field.id := field.id + 1;
      if r_num > skip and colmap.exists(field.id) then
        fields.extend;
        fields(fields.count) := ExcelTableCell(r_num, colmap(field.id), null, anydata.ConvertVarchar2(field.str_value), null, null);
      end if;
      -- reset field size
      field.sz := 0;    
    end;

    procedure next_line is
    begin
      r_num := r_num + 1;
      -- reset field id
      field.id := 0;
      if r_num > skip then
        r_cnt := r_cnt + 1;
      end if;
    end;
  
  begin
    
    field.id := 0;

    loop 
      
      field.start_offset := buf.offset;
      
      -- new field starting with text qualifier?
      if USE_TEXT_QUALIFIER and field.sz = 0 and substr(buf.content, buf.offset, 1) = fd.text_qualifier then
      
        buf.offset := buf.offset + 1;
          
        loop
            
          field.start_offset := buf.offset;
          tq_offset := instr(buf.content, fd.text_qualifier, buf.offset);
            
          if tq_offset != 0 then
              
            buf.offset := tq_offset + 1;
            field.end_offset := tq_offset - 1;
            read_field;
              
            if buf.offset > buf.sz then
              bufferize(stream, buf);
            end if;
            -- check whether tq found is an escape character
            if substr(buf.content, buf.offset, 1) = fd.text_qualifier then
              --tq found was an escape character
              field.str_value := field.str_value || fd.text_qualifier;
              buf.offset := buf.offset + 1;
              if buf.offset > buf.sz then
                bufferize(stream, buf);
              end if;
            elsif substr(buf.content, buf.offset, 1) = fd.field_separator then
              --tq found was this field's end delimiter
              buf.offset := buf.offset + 1;
              if buf.offset > buf.sz then
                bufferize(stream, buf);
              end if;
              exit;
            else
              error(-20742, 'Bad escape sequence');
            end if;
              
          else
              
            field.str_value := field.str_value || substr(buf.content, field.start_offset);
            bufferize(stream, buf);
            
          end if;
          
        end loop;
        
        add_field;
      
      else
      
        fs_offset := instr(buf.content, fd.field_separator, buf.offset);
        lt_offset := instr(buf.content, fd.line_terminator, buf.offset);
        
        if fs_offset != 0 and lt_offset != 0 and fs_offset < lt_offset
           or fs_offset != 0 and lt_offset = 0  
        then
          
          field.end_offset := fs_offset - 1;
          read_field;
          add_field;          
          buf.offset := fs_offset + FS_SZ;
          
        elsif fs_offset != 0 and lt_offset != 0 and lt_offset < fs_offset
           or lt_offset != 0 and fs_offset = 0
        then

          field.end_offset := lt_offset - 1;
          read_field; 
          add_field;
          buf.offset := lt_offset + LT_SZ;          
          next_line;
          
        else
          -- read until end of current block and bufferize
          lt_chunk := null;
          lt_chunk_sz := 0;
          for i in reverse 1 .. LT_SZ - 1 loop
            if buf.content like '%'||substr(fd.line_terminator,1,i) then
              lt_chunk := substr(fd.line_terminator,1,i);
              lt_chunk_sz := i;
              exit;
            end if;
          end loop;
          
          field.end_offset := buf.sz - lt_chunk_sz;
          read_field;
          
          if stream.available != 0 then
            bufferize(stream, buf);
            
            lt_chunk_tail_sz := LT_SZ - lt_chunk_sz;
            if lt_chunk_sz != 0 and buf.content like substr(fd.line_terminator,-lt_chunk_tail_sz)||'%' then
              
              add_field;
              buf.offset := buf.offset + lt_chunk_tail_sz;
              next_line;
              
            else
              -- append chunk to current field
              if field.sz != 0 then
                field.str_value := field.str_value || lt_chunk;
                field.sz := field.sz + lt_chunk_sz;
              else
                field.str_value := lt_chunk;
                field.sz := lt_chunk_sz;
              end if;
            end if;
            
          else
            if not (field.id = 0 and field.sz = 0) then
              add_field;
            end if;
            ctx_cache(ctx_id).done := true;
            exit;
          end if;
          
        end if;
        
      end if;
      
      exit when r_cnt >= nrows;
    
    end loop;
    
    ctx_cache(ctx_id).r_num := r_num;
    
    return fields;

  end;

  function read_fields_posn (
    stream in out nocopy stream_t
  , buf    in out nocopy buffer_t
  , ctx_id in pls_integer
  , nrows  in pls_integer
  )
  return ExcelTableCellList
  is
    --line_terminator_start_pattern varchar2(2) := '%'||substr(fd.line_terminator,1,1);
    --line_terminator_end_pattern varchar2(2) := substr(fd.line_terminator,-1)||'%';
    --line_terminator_start varchar2(2) := substr(fd.line_terminator,1,1);
    --line_terminator_end varchar2(2) := substr(fd.line_terminator,-1);
    
    fd                file_descriptor_t := ctx_cache(ctx_id).fd;
    skip              pls_integer := ctx_cache(ctx_id).skip;
    colmap            column_map_t := ctx_cache(ctx_id).colmap;
    
    LT_SZ             constant pls_integer := length(fd.line_terminator);
    BUFFER_SIZE       constant pls_integer := 1024;
    
    fields            ExcelTableCellList := ExcelTableCellList();
    line              block_list_t := block_list_t();
    lt_offset         pls_integer;
    lt_chunk_sz       pls_integer := 0;
    lt_chunk_tail_sz  pls_integer;
    --lt_partial_match  boolean;
    field_value       varchar2(32767);
    block_cnt         pls_integer;
    r_num             pls_integer := ctx_cache(ctx_id).r_num;
    r_cnt             pls_integer := 0;

    procedure init_lt_chunk
    is
    begin
      lt_chunk_sz := 0;
      for i in reverse 1 .. LT_SZ - 1 loop
        if buf.content like '%'||substr(fd.line_terminator,1,i) then
          lt_chunk_sz := i;
          exit;
        end if;
      end loop;
      lt_chunk_tail_sz := LT_SZ - lt_chunk_sz;
    end;
    
    procedure push(chunk in out nocopy buffer_t)
    is
      i pls_integer := nvl(line.last,0);
      chunk_amount pls_integer;
    begin
      
      if chunk.sz != 0 then
          
        if chunk.sz <= line(i).free then
          line(i).content := line(i).content || chunk.content;
          line(i).sz := line(i).sz + chunk.sz;
          line(i).free := line(i).free - chunk.sz;
        else
          chunk.available := chunk.sz;
          chunk.offset := 1;
          while chunk.available != 0 loop
            
            chunk_amount := least(line(i).free, chunk.available);
            line(i).content := line(i).content || substr(chunk.content, chunk.offset, chunk_amount);
            chunk.offset := chunk.offset + chunk_amount;
            chunk.available := chunk.available - chunk_amount;
            line(i).sz := line(i).sz + chunk_amount;
            line(i).free := line(i).free - chunk_amount;
            
            if line(i).free = 0 then
              i := i + 1;
              line.extend;
              line(i).sz := 0;
              line(i).free := BLOCK_SIZE;
            end if;
            
          end loop;
          
        end if;
      
      end if;
      
    end;
    
    procedure bufferize
    is
      amount  pls_integer := BUFFER_SIZE;
    begin
      dbms_lob.read(stream.content, amount, stream.offset, buf.content);
      stream.offset := stream.offset + amount;
      stream.available := stream.available - amount;
      buf.sz := amount;
      buf.offset := 1;
      buf.available := buf.sz;
    exception
      when no_data_found then
        buf.content := null;
        buf.available := 0;
    end;

    procedure read_line
    is
      chunk  buffer_t;
    begin
    
      --line.delete;
      if line.count > 1 then
        line.trim(line.count-1);
      end if;
      
      -- init
      line(1).content := null;
      line(1).sz := 0;
      line(1).free := BLOCK_SIZE;   
    
      loop
        
        if buf.available = 0 then
          bufferize;
        end if;
        
        lt_offset := instr(buf.content, fd.line_terminator, buf.offset);
        
        if lt_offset != 0 then
          chunk.sz := lt_offset - buf.offset;
          chunk.content := substr(buf.content, buf.offset, chunk.sz);
          push(chunk);
          buf.available := buf.available - chunk.sz - LT_SZ;
          buf.offset := lt_offset + LT_SZ;
          exit;
        else
          
          --check if buffer end matches start of line terminator sequence
          --lt_partial_match := ( LT_SZ > 1 and buf.content like line_terminator_start_pattern );
          init_lt_chunk;
          
          chunk.sz := buf.sz - buf.offset + 1;
          chunk.content := substr(buf.content, buf.offset);
                
          bufferize;
          
          --if start of new buffer matches end of line terminator sequence
          --if lt_partial_match and buf.content like line_terminator_end_pattern then
          if lt_chunk_sz != 0 and buf.content like substr(fd.line_terminator,-lt_chunk_tail_sz)||'%'
          then
            chunk.sz := chunk.sz - lt_chunk_sz;
            chunk.content := substr(chunk.content, 1, chunk.sz);
            buf.offset := buf.offset + lt_chunk_tail_sz;
            buf.available := buf.available - lt_chunk_tail_sz;
            push(chunk);
            exit;
          else
            push(chunk);
            exit when buf.available = 0;
          end if;  

        end if;
      
      end loop;
      
      r_num := r_num + 1;
      
    end;
    
  begin
       
    line.extend;
    
    while stream.available != 0 or buf.available != 0 loop
      
      read_line;
      
      if r_num > skip then
      
        block_cnt := line.count;
        
        for i in 1 .. fd.fields.count loop
          
          field_value := null;
          
          if fd.fields(i).block_start <= block_cnt then
        
            if fd.fields(i).is_single_block then     
              field_value := substr(line(fd.fields(i).block_start).content, fd.fields(i).block_start_offset, fd.fields(i).sz);
            else
              -- first block
              field_value := substr(line(fd.fields(i).block_start).content, fd.fields(i).block_start_offset);
              -- full blocks
              for j in fd.fields(i).block_start + 1 .. least(fd.fields(i).block_end, block_cnt) - 1 loop
                field_value := field_value || line(j).content;
              end loop;
              -- last block
              if fd.fields(i).block_end <= block_cnt then
                field_value := field_value || substr(line(fd.fields(i).block_end).content, 1, fd.fields(i).block_end_offset);
              end if;
              
            end if;
          
          end if;
          
          fields.extend;
          fields(fields.last) := ExcelTableCell(r_num, colmap(i), null, anydata.ConvertVarchar2(trim(field_value)), null, null);
        
        end loop;
        
        r_cnt := r_cnt + 1;
      
      end if;
      
      exit when r_cnt >= nrows;
      
    end loop;
    
    if stream.available = 0 then
      ctx_cache(ctx_id).done := true;
    end if;
    
    ctx_cache(ctx_id).r_num := r_num;
    
    return fields;

  end;
  
  procedure set_file_descriptor (
    p_ctx_id    in pls_integer
  , p_field_sep in varchar2 default DEFAULT_FIELD_SEP
  , p_line_term in varchar2 default DEFAULT_LINE_TERM
  , p_text_qual in varchar2 default DEFAULT_TEXT_QUAL
  )
  is
  begin
    ctx_cache(p_ctx_id).fd.field_separator := p_field_sep;
    ctx_cache(p_ctx_id).fd.line_terminator := p_line_term;
    ctx_cache(p_ctx_id).fd.text_qualifier := p_text_qual;
  end;
  
  function new_context (
    p_content  in clob
  , p_cols     in varchar2
  , p_skip     in pls_integer
  , p_type     in pls_integer default TYPE_DELIMITED
  )
  return pls_integer
  is
    ctx     context_t;
    ctx_id  pls_integer; 
  begin
    
    ctx.stream.content := p_content;
    ctx.stream.sz := dbms_lob.getlength(ctx.stream.content);
    ctx.stream.offset := 1;
    ctx.stream.available := ctx.stream.sz;
    
    ctx.ctype := p_type;
    ctx.skip := p_skip;
    
    case ctx.ctype
    when TYPE_DELIMITED then
      ctx.colmap := parseColumnList(p_cols);
      ctx.r_num := 1;
    when TYPE_POSITIONAL then
      ctx.fd.fields := parseFieldList(p_cols);
      for i in 1 .. ctx.fd.fields.count loop
        ctx.colmap(i) := base26encode(i);
      end loop;
    else
      error(-20742, 'Invalid context type');
    end case;
    
    ctx_id := nvl(ctx_cache.last, 0) + 1;
    ctx_cache(ctx_id) := ctx;
    
    return ctx_id;
    
  end;

  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer
  )
  return ExcelTableCellList
  is
    fields  ExcelTableCellList;
  begin
    if not ctx_cache(p_ctx_id).done then
      case ctx_cache(p_ctx_id).ctype
      when TYPE_DELIMITED then
        fields := read_fields(ctx_cache(p_ctx_id).stream, ctx_cache(p_ctx_id).buf, p_ctx_id, p_nrows);
      when TYPE_POSITIONAL then
        fields := read_fields_posn(ctx_cache(p_ctx_id).stream, ctx_cache(p_ctx_id).buf, p_ctx_id, p_nrows);
      end case;
    end if;
    return fields;
  end;
  
  procedure free_context (
    p_ctx_id  in pls_integer 
  )
  is
  begin
    dbms_lob.freetemporary(ctx_cache(p_ctx_id).stream.content);
    ctx_cache.delete(p_ctx_id);
  end;
  
  function get_fields_delimited (
    p_content   in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer default 0
  , p_line_term in varchar2 default DEFAULT_LINE_TERM
  , p_field_sep in varchar2 default DEFAULT_FIELD_SEP
  , p_text_qual in varchar2 default DEFAULT_TEXT_QUAL
  )
  return ExcelTableCellList 
  pipelined
  is
    ctx_id  pls_integer;
    fields  ExcelTableCellList;
  begin
    
    ctx_id := new_context(p_content, p_cols, p_skip, TYPE_DELIMITED);
    set_file_descriptor(ctx_id, p_field_sep, p_line_term, p_text_qual);
    
    while not ctx_cache(ctx_id).done loop
      fields := iterate_context(ctx_id, 100);
      for i in 1 .. fields.count loop
        pipe row (fields(i));
      end loop;
    end loop;
    
    free_context(ctx_id);
    
  end;

  function get_fields_positional (
    p_content   in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer default 0
  , p_line_term in varchar2 default DEFAULT_LINE_TERM
  )
  return ExcelTableCellList 
  pipelined
  is
    ctx_id  pls_integer;
    fields  ExcelTableCellList;
  begin
    
    ctx_id := new_context(p_content, p_cols, p_skip, TYPE_POSITIONAL);
    set_file_descriptor(ctx_id, p_line_term => p_line_term);
    
    while not ctx_cache(ctx_id).done loop
      fields := iterate_context(ctx_id, 100);
      for i in 1 .. fields.count loop
        pipe row (fields(i));
      end loop;
    end loop;
    
    free_context(ctx_id);
    
  end;  

end xutl_flatfile;
/
