
package db.office.spreadsheet.oox;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;

import db.office.spreadsheet.Cell;
import db.office.spreadsheet.CellReader;
import db.office.spreadsheet.CellReaderException;
import db.office.spreadsheet.CellRef;
import db.office.spreadsheet.Row;
import db.office.spreadsheet.Sheet;

public class OOXCellReaderImpl implements CellReader {

	private static final String TAG_SHEETDATA = "sheetData";
	private static final String TAG_ROW = "row";
	private static final String TAG_C = "c";
	private static final String TAG_V = "v";
	private static final String TAG_T = "t";
	private static final String TAG_R = "r";
	
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
	private Map<String,String> comments;
	
	public OOXCellReaderImpl(String columnList, int firstRow, int lastRow)
			throws CellReaderException {
		
		this.factory = XMLInputFactory.newInstance();
		this.columns = new HashSet<String>(Arrays.asList(columnList.split(",")));
		this.firstRow = firstRow;
		this.lastRow = lastRow;
		this.sheets = new ArrayList<Sheet>();
		comments = new HashMap<String,String>();
		
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
	
	private void initSheetIterator() throws CellReaderException {
		if (this.sheetIterator == null) {
			this.sheetIterator = this.sheets.iterator();
			this.nextSheet();
		}		
	}
	
	public void addSheet(Sheet sheet) {
		sheets.add(sheet);
	}
	
	public void addComments(int sheetIndex, InputStream is) throws CellReaderException, IOException {
		comments.putAll(OOXCommentReader.getComments(sheetIndex, is));
	}
	
	public void readStrings(InputStream is) throws IOException, CellReaderException {
		strings = SharedStringsHandler.getStrings(is);
	}
	
	public List<String> getSheetList() {
		return null;
	}
	
	private void nextSheet() throws CellReaderException {
		
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
		
		initSheetIterator();
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
							Row row = readRow(rowRef);
							if (!row.isEmpty()) {
								rows.add(row);
								nrows--;
							}			
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
	    Cell<?> cell;
		
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
	
	private Cell<?> readCell() throws XMLStreamException {

		String cellRefValue = this.reader.getAttributeValue(null, TAG_R);
		CellRef cellRef = new CellRef(cellRefValue);
		Cell<?> cell = null;
		
		if (columns.contains(cellRef.getColRef())) {
		
			String cellTypeAttr = this.reader.getAttributeValue(null, TAG_T);
			OOXCellType cellType = (cellTypeAttr!=null)?OOXCellType.get(cellTypeAttr):OOXCellType.NUMBER;
			String cellValue = null;

			switch (cellType) {
			case SHAREDSTRING:
				int idx = Integer.parseInt(readCellValue());
				cellValue = this.strings[idx];
				break;
				
			case INLINESTR:
				// when t="inlineStr", <is> is the only child element allowed
				cellValue = readInlineStr();
				break;
				
			default:
				cellValue = readCellValue();
			
			}

			if (cellValue != null) {
				cell = new OOXCell(cellRef, cellValue, cellType, this.sheetIndex);
				cell.setAnnotation(comments.get(String.valueOf(sheetIndex) + cellRefValue));
			}
			
		} else {
			//cell = new OOXCell(cellRef, null, OOXCellType.UNKNOWN, this.sheetIndex);
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
						return null;
					}
			}
			
		}
		
		throw new IllegalStateException();
		
	}
	
}
