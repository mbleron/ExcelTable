alter session set plsql_optimize_level=3;

prompt Creating package XUTL_CDF ...
@@MSUtilities/CDFManager/xutl_cdf.pks
@@MSUtilities/CDFManager/xutl_cdf.pkb

prompt Creating package XUTL_OFFCRYPTO ...
@@MSUtilities/OfficeCrypto/xutl_offcrypto.pks
@@MSUtilities/OfficeCrypto/xutl_offcrypto.pkb

prompt Creating type ExcelTableSheetList ...
@@ExcelCommons/plsql/ExcelTableSheetList.tps

prompt Creating type ExcelTableCell ...
@@ExcelCommons/plsql/ExcelTableCell.tps

prompt Creating type ExcelTableCellList ...
@@ExcelCommons/plsql/ExcelTableCellList.tps

prompt Creating package ExcelTypes ...
@@ExcelCommons/plsql/ExcelTypes.pks
@@ExcelCommons/plsql/ExcelTypes.pkb

prompt Creating package XUTL_XLS ...
@@ExcelCommons/plsql/xutl_xls.pks
@@ExcelCommons/plsql/xutl_xls.pkb

prompt Creating package XUTL_XLSB ...
@@ExcelCommons/plsql/xutl_xlsb.pks
@@ExcelCommons/plsql/xutl_xlsb.pkb

prompt Creating package XUTL_FLATFILE ...
@@ExcelCommons/plsql/xutl_flatfile.pks
@@ExcelCommons/plsql/xutl_flatfile.pkb

prompt Creating type ExcelTableImpl ...
@@plsql/ExcelTableImpl.tps

prompt Creating package ExcelTable ...
@@plsql/ExcelTable.pks
@@plsql/ExcelTable.pkb

prompt Creating type body ExcelTableImpl ...
@@plsql/ExcelTableImpl.tpb

