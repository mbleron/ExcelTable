package db.office.spreadsheet.odf;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import db.office.spreadsheet.Cell;
import db.office.spreadsheet.CellReader;
import db.office.spreadsheet.CellReaderException;
import db.office.spreadsheet.CellRef;
import db.office.spreadsheet.Row;
import db.office.spreadsheet.Sheet;
import oracle.xml.parser.v2.SAXParser;

public class ODFCellReaderImpl implements CellReader {

	static final String ODF_TABLE_NS = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
	static final String ODF_OFFICE_NS = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
	static final String ODF_TEXT_NS = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
	
	static final QName QNAME_TABLE = new QName(ODF_TABLE_NS, "table");
	static final QName QNAME_TNAME = new QName(ODF_TABLE_NS, "name");
	static final QName QNAME_TROW = new QName(ODF_TABLE_NS, "table-row");
	static final QName QNAME_TCELL = new QName(ODF_TABLE_NS, "table-cell");
	static final QName QNAME_VTYPE = new QName(ODF_OFFICE_NS, "value-type");
	static final QName QNAME_VALUE = new QName(ODF_OFFICE_NS, "value");
	static final QName QNAME_DATE_VALUE = new QName(ODF_OFFICE_NS, "date-value");
	static final QName QNAME_STRING_VALUE = new QName(ODF_OFFICE_NS, "string-value");
	static final QName QNAME_BOOLEAN_VALUE = new QName(ODF_OFFICE_NS, "boolean-value");
	static final QName QNAME_PARAGRAPH = new QName(ODF_TEXT_NS, "p");
	static final QName QNAME_ANNOTATION = new QName(ODF_OFFICE_NS, "annotation");
	
	private boolean initialized = false;
	private boolean done = false;
	private int firstRow = 0;
	private int lastRow;
	private int rowIndex = 0;
	private int cellIndex = 0;
	
	private Row row;
	int row_repeat = 0;
	
	private XMLStreamReader reader = null;
	private Set<String> columns;
	private List<String> sheetList;
	private Set<Integer> sheetIndices;
	private int sheetIndex = 0;
	
	public ODFCellReaderImpl(String columnList, int firstRow, int lastRow) {
		this.columns = new HashSet<String>(Arrays.asList(columnList.split(",")));
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		sheetIndices = new HashSet<Integer>();
	}
	
	public void setContent(InputStream content) 
			throws CellReaderException, IOException {
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		
        SAXParser parser = new SAXParser();
        parser.setValidationMode(SAXParser.NONVALIDATING);
        parser.setPreserveWhitespace(false);
        
        SAXDocumentSerializer serializer = new SAXDocumentSerializer();
        serializer.setOutputStream(out);
        
        parser.setContentHandler(serializer);
        //parser.setProperty("http://xml.org/sax/properties/lexical-handler", serializer);
        
        try {
			parser.parse(content);
		} catch (Exception e) {
			throw new CellReaderException("Error while parsing content", e);
		}
        content.close();
        
        sheetList = serializer.getTableNames();
        
        XMLInputFactory factory = new com.sun.xml.fastinfoset.stax.factory.StAXInputFactory();
		InputStream is = new ByteArrayInputStream(out.toByteArray());
		try {
			reader = factory.createXMLStreamReader(is);
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error while creating XMLStreamReader instance", e);
		}		
	}
	
	public void close() throws CellReaderException {
		try {
			reader.close();
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error while closing CellReader", e);
		}
	}
	
	public int getColumnCount() {
		return columns.size();
	}
	
	public void addSheet(Sheet sheet) {
		sheetIndices.add(sheet.getIndex());
	}
	
	public List<String> getSheetList() {
		return sheetList;
	}
	
	public void nextSheet() throws CellReaderException {
		
		try {
			if (sheetIndices.isEmpty()) {
				done = true;
			} else {	
				while (reader.hasNext()) {
					int event = reader.next();
					if (event == XMLStreamReader.START_ELEMENT && reader.getLocalName().equals(QNAME_TABLE.getLocalPart())) {
						sheetIndex++;
						if (sheetIndices.contains(sheetIndex)) {
							rowIndex = 0;
							sheetIndices.remove(sheetIndex);
							break;
						}
					}
				}
			}
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error in nextSheet() method", e);
		}
		
	}
	
	public List<Row> readRows(int nrows) throws CellReaderException {
		
		if (!initialized) {
			nextSheet();
			initialized = true;
		}
		
		List<Row> rows = new ArrayList<Row>(nrows);
		row = null;

		try {

			while (!done && reader.hasNext() && nrows > 0) {

				int event = reader.next();

				switch (event) {
				case XMLStreamReader.START_ELEMENT:	
					if (reader.getLocalName().equals(QNAME_TROW.getLocalPart())) {

						do {
							rowIndex++;
							System.out.println(rowIndex);
							if (row_repeat == 0) {

								String row_repeat_attr = reader.getAttributeValue(null, "number-rows-repeated");
								row_repeat = ((row_repeat_attr==null)?1:Integer.parseInt(row_repeat_attr)) - 1;

								if (rowIndex >= firstRow - row_repeat) {
									row = readRow(rowIndex);
								}

							} else {
								row = Row.copyTo(row, rowIndex);
								row_repeat--;
							}

							if (rowIndex >= firstRow) {
								if (!row.isEmpty()) {
									rows.add(row);
									nrows--;
								} else if (row_repeat > 0) {
									// ignore empty repeating rows
									rowIndex += row_repeat;
									row_repeat = 0;
								}
								if (lastRow != -1 && rowIndex >= lastRow) {
									nextSheet();
									break;
								}
							}

						} while (row_repeat != 0);

					} 
					break;

				case XMLStreamReader.END_ELEMENT:
					if (reader.getLocalName().equals(QNAME_TABLE.getLocalPart())) {
						nextSheet();
					}
				}
			}

			return rows;

		} catch (XMLStreamException e) {
			throw new CellReaderException("Error in readRows method", e);
		}

	}
	
	private Row readRow(int rowIndex) throws XMLStreamException {
		
		Row row = new Row(rowIndex, this.columns.size());
	    Cell<?> cell = null;
	    cellIndex = 0;
	    int cell_repeat = 0;
		
		while (reader.hasNext()) {
			int event = reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (reader.getLocalName().equals(QNAME_TCELL.getLocalPart())) {
						
						do {						
							cellIndex++;
							CellRef cellRef = new CellRef(rowIndex, cellIndex);
							
							if (cell_repeat == 0) {
								String cell_repeat_attr = reader.getAttributeValue(null, "number-columns-repeated");
								cell_repeat = (cell_repeat_attr==null)?1:Integer.parseInt(cell_repeat_attr);
								cell = readCell(cellRef);
							} else if (cell != null) {
								cell = cell.copyToColumn(cellIndex);
							}
							
							if (cell != null && columns.contains(cellRef.getColRef())) {
								row.add(cell);
							}
							cell_repeat--;
						} while (cell_repeat != 0);

					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (reader.getLocalName().equals(QNAME_TROW.getLocalPart())) {
						return row;
					}
			}
			
		}
		
		throw new IllegalStateException();
		
	}	

	private Cell<?> readCell(CellRef cellRef) throws XMLStreamException {
		ODFCell cell = null;
		String valueTypeAttr = reader.getAttributeValue(QNAME_VTYPE.getNamespaceURI(), QNAME_VTYPE.getLocalPart());
		if (valueTypeAttr != null) {
			String cellValue = null;
			ODFValueType valueType = ODFValueType.get(valueTypeAttr);
			switch (valueType) {
			case FLOAT:
			case PERCENTAGE:
				cellValue = reader.getAttributeValue(QNAME_VALUE.getNamespaceURI(), QNAME_VALUE.getLocalPart());
				break;
				
			case STRING:
				cellValue = reader.getAttributeValue(QNAME_STRING_VALUE.getNamespaceURI(), QNAME_STRING_VALUE.getLocalPart());
				break;
				
			case DATE:
				cellValue = reader.getAttributeValue(QNAME_DATE_VALUE.getNamespaceURI(), QNAME_DATE_VALUE.getLocalPart());
				break;
				
			case BOOLEAN:
				cellValue = reader.getAttributeValue(QNAME_BOOLEAN_VALUE.getNamespaceURI(), QNAME_BOOLEAN_VALUE.getLocalPart());
				break;
				
			}
			
			cell = new ODFCell(cellRef, cellValue, valueType, sheetIndex);
			
			//read cell annotation and content
			readContent(cell);
			
		}
		return cell;
	}
	
	private void readContent(ODFCell cell) throws XMLStreamException {
		StringBuilder sb = new StringBuilder();
		int paragraphCount = 0;
		boolean inParagraph = false;
		readContent:
		while (reader.hasNext()) {
			int event = reader.next();
			QName name;
			switch (event) {
			case XMLStreamReader.START_ELEMENT:
				name = reader.getName();
				if (name.equals(QNAME_PARAGRAPH)) {
					if (paragraphCount != 0) {
						sb.append('\n');
					}
					paragraphCount++;
					inParagraph = true;
				}
				break;
				
			case XMLStreamReader.CHARACTERS:
				if (inParagraph) {
					sb.append(reader.getText());
				}
				break;
				
			case XMLStreamReader.END_ELEMENT:
				name = reader.getName();
				if (name.equals(QNAME_PARAGRAPH)) {
					inParagraph = false;
				} else if (name.equals(QNAME_TCELL)) {
					if (cell.isOf(ODFValueType.STRING)) {
						cell.setValue(sb.toString());
					}
					break readContent;
				} else if (name.equals(QNAME_ANNOTATION)) {
					cell.setAnnotation(sb.toString());
					//reset builder
					sb = new StringBuilder();
					paragraphCount = 0;
				}
			}

		}		
	}
	
}
