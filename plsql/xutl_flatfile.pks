create or replace package xutl_flatfile is
/* ======================================================================================

  MIT License

  Copyright (c) 2019 Marc Bleron

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
    Marc Bleron       2019-05-21     Creation

====================================================================================== */

  -- Flat File constants
  FF_COMMA               constant varchar2(1) := chr(44); -- ,
  FF_SEMICOLON           constant varchar2(1) := chr(59); -- ;
  FF_CRLF                constant varchar2(2) := chr(13)||chr(10);
  FF_LF                  constant varchar2(2) := chr(10);
  FF_QUOTATION_MARK      constant varchar2(1) := chr(34); -- "
  FF_APOSTROPHE          constant varchar2(1) := chr(39); -- '
  
  TYPE_DELIMITED         constant pls_integer := 0;
  TYPE_POSITIONAL        constant pls_integer := 1;
  
  DEFAULT_FIELD_SEP      constant varchar2(1) := FF_COMMA;
  DEFAULT_LINE_TERM      constant varchar2(2) := FF_CRLF;
  DEFAULT_TEXT_QUAL      constant varchar2(1) := FF_QUOTATION_MARK;
  
  function new_context (
    p_content  in clob
  , p_cols     in varchar2
  , p_skip     in pls_integer
  , p_type     in pls_integer default TYPE_DELIMITED
  )
  return pls_integer;
  
  procedure free_context (
    p_ctx_id  in pls_integer 
  );
  
  function iterate_context (
    p_ctx_id  in pls_integer
  , p_nrows   in pls_integer
  )
  return ExcelTableCellList;

  procedure set_file_descriptor (
    p_ctx_id    in pls_integer
  , p_field_sep in varchar2 default DEFAULT_FIELD_SEP
  , p_line_term in varchar2 default DEFAULT_LINE_TERM
  , p_text_qual in varchar2 default DEFAULT_TEXT_QUAL
  );
  
  function get_fields_delimited (
    p_content   in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer default 0
  , p_line_term in varchar2 default DEFAULT_LINE_TERM
  , p_field_sep in varchar2 default DEFAULT_FIELD_SEP
  , p_text_qual in varchar2 default DEFAULT_TEXT_QUAL
  )
  return ExcelTableCellList 
  pipelined;

  function get_fields_positional (
    p_content   in clob
  , p_cols      in varchar2
  , p_skip      in pls_integer default 0
  , p_line_term in varchar2 default DEFAULT_LINE_TERM
  )
  return ExcelTableCellList 
  pipelined;

end xutl_flatfile;
/
