alter session set plsql_optimize_level=3;

prompt Creating package XUTL_CDF ...
@@MSUtilities/CDFReader/xutl_cdf.pks
@@MSUtilities/CDFReader/xutl_cdf.pkb

prompt Creating package XUTL_OFFCRYPTO ...
@@MSUtilities/OfficeCrypto/xutl_offcrypto.pks
@@MSUtilities/OfficeCrypto/xutl_offcrypto.pkb

prompt Creating type ExcelTableSheetList ...
@@plsql/ExcelTableSheetList.tps

prompt Creating type ExcelTableCell ...
@@plsql/ExcelTableCell.tps

prompt Creating type ExcelTableCellList ...
@@plsql/ExcelTableCellList.tps

prompt Creating package XUTL_XLS ...
@@plsql/xutl_xls.pks
@@plsql/xutl_xls.pkb

prompt Creating package XUTL_XLSB ...
@@plsql/xutl_xlsb.pks
@@plsql/xutl_xlsb.pkb

prompt Creating package XUTL_FLATFILE ...
@@plsql/xutl_flatfile.pks
@@plsql/xutl_flatfile.pkb

prompt Creating type ExcelTableImpl ...
@@plsql/ExcelTableImpl.tps

prompt Creating package ExcelTable ...
@@plsql/ExcelTable.pks
@@plsql/ExcelTable.pkb

prompt Creating type body ExcelTableImpl ...
@@plsql/ExcelTableImpl.tpb

