package db.office.spreadsheet.oox;

import java.io.IOException;
import java.io.InputStream;

import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import db.office.spreadsheet.CellReaderException;
import oracle.xml.parser.v2.SAXParser;

public class SharedStringsHandler extends DefaultHandler {

	private StringBuilder sb;
	private boolean	in_text = false;
	private String[] strings;
	private int idx = 0;

	public SharedStringsHandler() {
		this.sb = new StringBuilder();	   
	}
	
	public String[] getStrings () {
		return this.strings;
	}

	public void startElement(String uri, String localName, String qName, Attributes attrs) 
			throws SAXException {
		
		if (localName.equals("t")) {
			this.in_text = true;
		} else if (localName.equals("sst")) {
			String uniqueCount = attrs.getValue("uniqueCount");
			if (uniqueCount != null) {
				int cnt = Integer.parseInt(uniqueCount);
				this.strings = new String[cnt];
			}
		}

	}

	public void endElement(String uri, String localName,
			String qName) throws SAXException {
		
		if (localName.equals("si")) {
			strings[idx++] = this.sb.toString();
			this.sb.setLength(0);
		} else if (localName.equals("t")) {
			this.in_text = false;
		}
	}

	public void characters(char[] ch, int start, int length) 
			throws SAXException {
		if (this.in_text) {
			this.sb.append(ch, start, length);
		}
	}
	
	public static String[] getStrings (InputStream is) 
			throws IOException, CellReaderException {
		
		SharedStringsHandler handler = null;
		try {
			handler = new SharedStringsHandler();
			SAXParser p = new SAXParser();
			p.setValidationMode(SAXParser.NONVALIDATING);
			p.setContentHandler(handler);
			p.parse(is);
		} catch (Exception e) {
			throw new CellReaderException("Error while parsing sharedStrings content", e);
		} finally {
			is.close();
		}
		
		return handler.getStrings();

	}
	
}

