
# Change Log

## [5.4.1] - 2024-01-10

### Fixed
* Missing VARIANT case while parsing a compiled table definition

## [5.4] - 2023-09-14

### Added
* Documentation: moved change log to its own file
* Documentation: new SVG syntax diagrams
* Error handling for heterogeneous data (issue #45)
* New "VARIANT" object type in SQL projection

## [5.3] - 2023-05-23

* Enhancement : Date type detection

## [5.2.2] - 2021-08-20
* Enhancement : issue #34

## [5.2.1] - 2021-08-09
* Fix : issue #33

## [5.2.0] - 2021-07-13
* Enhancements :
  * added isReadMethodAvailable function
  * added default parameter p_method to getSheets function (backwards compatible)

## [5.1.1] - 2021-02-12
* Fix : issue #26

## [5.1] - 2020-04-17

* Enhancements :
	* added getSheets function
	* documentation links updated for procedure loadData
	* documentation for mapColumn mentions now that the column name is case sensitive

## [5.0] - 2020-03-25
* Fix : issue #18 
* Enhancements : 
  * issue #19
  * Support for strict OOXML documents  
  * Streaming read method for ODF files  
  * Raw cells listing

## [4.0.1] - 2019-09-29
* Fix : issue #14
* Enhancement : issue #15

## [4.0] - 2019-09-22
* Added support for delimited and positional flat files
* Fix : issue #12
* Fix : issue #13

## [3.2.1] - 2019-05-14
* Fix : requested rows count wrongly decremented for empty row
* Fix : getCursor() failure with multi-sheet support

## [3.2] - 2019-05-01
* Added support for XML spreadsheetML files (.xml)

## [3.1] - 2019-04-20
* New default value feature in DML API

## [3.0] - 2019-03-30
* Multi-sheet support

## [2.3.2] - 2018-10-22
* XUTL_XLS enhancement (new buffered lob reader)

## [2.3.1] - 2018-09-15
* XUTL_XLS enhancement

## [2.3] - 2018-08-23
* New API for DML operations
* Internal modularization, unified interface for cell sources

## [2.2] - 2018-07-07
* Added support for OpenDocument (ODF) spreadsheets (.ods), including encrypted files
* Added support for TIMESTAMP data type

## [2.1] - 2018-04-22
* Added support for Excel Binary File Format (.xlsb)

## [2.0] - 2018-04-01
* Added support for Excel 97-2003 files (.xls)

## [1.6.1] - 2018-03-17
* Added large strings support for versions prior 11.2.0.2 

## [1.6] - 2017-12-31

* Added cell comments extraction
* Internal modularization

## [1.5] - 2017-07-10

* Fixed bug related to zip archives created with data descriptors. Now reading CRC-32, compressed and uncompressed sizes directly from Central Directory entries.
* Removed dependency to V$PARAMETER view (thanks [Paul](https://paulzipblog.wordpress.com/) for the suggestion)

## [1.4] - 2017-06-11

* Added getCursor() function
* Fixed NullPointerException when using streaming method and file has no sharedStrings

## [1.3] - 2017-05-30

* Added support for password-encrypted files
* Fixed minor bugs

## [1.2] - 2016-10-30

* Added new streaming read method
* Added setFetchSize() procedure

## [1.1] - 2016-06-25

* Added internal collection and LOB freeing

## [1.0] - 2016-05-01

* Creation
