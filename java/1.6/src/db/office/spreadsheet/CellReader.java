package db.office.spreadsheet;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Blob;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

public class CellReader {

	private static final String TAG_SHEETDATA = "sheetData";
	private static final String TAG_ROW = "row";
	private static final String TAG_C = "c";
	private static final String TAG_V = "v";
	private static final String TAG_T = "t";
	//private static final String TAG_IS = "is";
	private static final String TAG_R = "r";
	//private static final String CT_SHAREDSTRING = "s";
	//private static final String CELLTYPE_STR = "str";
	//private static final String CT_INLINESTR = "inlineStr";
	
	private boolean done = false;
	private int firstRow = 0;
	private int lastRow;
	private InputStream source;
	private XMLInputFactory factory;
	private XMLStreamReader reader = null;
	private String[] strings;
	private Set<String> columns;
	private List<Sheet> sheets;
	private Iterator<Sheet> sheetIterator = null;
	private int sheetIndex = 0;
	
	public CellReader(InputStream sharedStrings, String columnList, int firstRow, int lastRow)
			throws CellReaderException, IOException {
		
		this.factory = XMLInputFactory.newInstance();
		if (sharedStrings != null) {
			this.strings = SharedStringsHandler.getStrings(sharedStrings);
		}
			
		this.columns = new HashSet<String>(Arrays.asList(columnList.split(",")));
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.sheets = new ArrayList<Sheet>();
		
	}
	
	public void close() throws CellReaderException, IOException {
		if (this.reader != null) {
			try {
				this.reader.close();
				this.reader = null;
			} catch (XMLStreamException e) {
				throw new CellReaderException("Error while closing CellReader", e);
			} finally {
				this.source.close();
			}
		}
	}
	
	public int getColumnCount() {
		return this.columns.size();
	}
	
	private void setReaderSource(InputStream is) throws CellReaderException {
		
		try {
			this.reader = factory.createXMLStreamReader(is);
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error while creating XMLStreamReader instance", e);
		}
		
	}
	
	public void initSheetIterator() throws CellReaderException {
		if (this.sheetIterator == null) {
			this.sheetIterator = this.sheets.iterator();
			this.nextSheet();
		}		
	}
	
	public void addSheet(int index, Blob content) {
		this.sheets.add(new Sheet(index, content));
	}
	
	public void nextSheet() throws CellReaderException {
		
		try {
			// close existing Reader instance
			this.close();
			
			// switch to next sheet
			if (this.sheetIterator.hasNext()) {
				Sheet sheet = this.sheetIterator.next();
				this.sheetIndex = sheet.getIndex();
				this.source = sheet.getContent().getBinaryStream();
				setReaderSource(this.source);
			} else {
				this.done = true;
			}
			
		} catch (Exception e) {
			if (e instanceof CellReaderException) {
				throw (CellReaderException) e;
			}
			else {
				throw new CellReaderException("Error in nextSheet() method", e);
			}
		}
		
	}
	
	public List<Row> readRows(int nrows) throws CellReaderException {
		
		List<Row> rows = null;

		try {

			rows = new ArrayList<Row>(nrows);
			
			while (!this.done && this.reader.hasNext() && nrows > 0) {
				
				int event = this.reader.next();
				
				switch (event) {
				case XMLStreamReader.START_ELEMENT:	
					if (this.reader.getLocalName().equals(TAG_ROW)) {
	
						int rowRef = Integer.parseInt(this.reader.getAttributeValue(null, TAG_R));
	
						if (rowRef >= this.firstRow) {
							rows.add(readRow(rowRef));
							nrows--;
							if (this.lastRow != -1 && rowRef >= this.lastRow) {
								this.nextSheet();
							}
						}
						
					} 
					break;
					
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_SHEETDATA)) {
						this.nextSheet();
					}
				}
			}
			
			return rows;
			
		} catch (XMLStreamException e) {
			throw new CellReaderException("Error in readRows method", e);
		}

	}	
	
	private Row readRow(int rowRef) throws XMLStreamException {
		
		Row row = new Row(rowRef, this.columns.size());
	    Cell cell;
		
		while (this.reader.hasNext()) {
			int event = this.reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_C)) {
						if ((cell = readCell()) != null) {
							row.add(cell);
						}
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_ROW)) {
						return row;
					}
			}
			
		}
		
		throw new IllegalStateException();
		
	}	
	
	private Cell readCell() throws XMLStreamException {

		CellRef cellRef = new CellRef(this.reader.getAttributeValue(null, TAG_R));
		Cell cell = null;
		
		if (columns.contains(cellRef.column)) {
		
			String cellType = this.reader.getAttributeValue(null, TAG_T);
			String cellValue = null;

			if (Cell.CT_INLINESTR.equals(cellType)) {
				// when t="inlineStr", <is> is the only child element allowed
				cellValue = readInlineStr();
			} else if (Cell.CT_SHAREDSTRING.equals(cellType)) {
				int idx = Integer.parseInt(readCellValue());
				cellValue = this.strings[idx];
			} else {
				cellValue = readCellValue();
			}

			cell = new Cell(cellRef, cellValue, cellType, this.sheetIndex);

		} else {
			cell = new Cell(cellRef, "", "", this.sheetIndex);
		}

		return cell;

	}
	
	private String readInlineStr() throws XMLStreamException {

		StringBuilder sb = new StringBuilder();
		
		while (this.reader.hasNext()) {
			int event = this.reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_T)) {
						sb.append(this.reader.getElementText());
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_C)) {
						return sb.toString();
					}
			}	
		}
		
		throw new IllegalStateException();
		
	}
	
	private String readCellValue() throws XMLStreamException {
		
		while (this.reader.hasNext()) {
			int event = this.reader.next();
			switch (event) {
				case XMLStreamReader.START_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_V)) {
						return this.reader.getElementText();
					}
					break;
				case XMLStreamReader.END_ELEMENT:
					if (this.reader.getLocalName().equals(TAG_C)) {
						return "";
					}
			}
			
		}
		
		throw new IllegalStateException();
		
	}
	
}
