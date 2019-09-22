create or replace type ExcelTableCell as object (
  cellRow   integer
, cellCol   varchar2(3)
, cellType  varchar2(10)
, cellData  anydata
, sheetIdx  integer
)
/
