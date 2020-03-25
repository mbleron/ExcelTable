package db.office.spreadsheet;

import java.io.IOException;
import java.sql.Array;
import java.sql.Blob;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Struct;
import java.util.ArrayList;
import java.util.List;

import db.office.spreadsheet.odf.ODFCellReaderImpl;
import db.office.spreadsheet.oox.OOXCellReaderImpl;
import oracle.CartridgeServices.ContextManager;
import oracle.CartridgeServices.CountException;
import oracle.CartridgeServices.InvalidKeyException;
import oracle.jdbc.OracleConnection;

public class ReadContext {
	
	public static int VC2_MAXSIZE;

	public CellReader reader;
	
	private static ReadContext get(int key) throws CellReaderException {
		ReadContext ctx;
		try {
			ctx = (ReadContext) ContextManager.getContext(key);
		} catch (InvalidKeyException e) {
			throw new CellReaderException("Invalid context key", e);
		}
		return ctx;
	}
	
	public ReadContext(String type, String columns, int firstRow, int lastRow)
			throws CellReaderException {

		if (type.equals("OOX")) {
			this.reader = new OOXCellReaderImpl(columns, firstRow, lastRow);
		} else if (type.equals("ODF")) {
			this.reader = new ODFCellReaderImpl(columns, firstRow, lastRow);
		} else {
			throw new IllegalArgumentException("Invalid context type : " + type);
		}

	}
	
	public static int initialize(String type, String columns, int firstRow, int lastRow, int vc2MaxSize)
			throws CellReaderException {
				
		VC2_MAXSIZE = vc2MaxSize; 
		ReadContext ctx = new ReadContext(type, columns, firstRow, lastRow);
		
		int key = 0;
		try {
			key = ContextManager.setContext(ctx);
		} catch (CountException e) {
			throw new CellReaderException("Error during context creation", e);
		}
		return key;

	}
	
	public static Array getSheetList(int key) throws SQLException, CellReaderException {
		OracleConnection conn = (OracleConnection) DriverManager.getConnection("jdbc:default:connection:");	
		ReadContext ctx = ReadContext.get(key);
		return conn.createOracleArray("EXCELTABLESHEETLIST", ctx.reader.getSheetList().toArray());
	}
	
	public static void setContent(int key, Blob content) 
			throws CellReaderException, IOException, SQLException {
		ReadContext ctx = ReadContext.get(key);
		((ODFCellReaderImpl) ctx.reader).setContent(content.getBinaryStream());
	}
	
	public static void setSharedStrings(int key, Blob sharedStrings) 
			throws CellReaderException, IOException, SQLException {
		ReadContext ctx = ReadContext.get(key);
		if (sharedStrings != null) {
			((OOXCellReaderImpl) ctx.reader).readStrings(sharedStrings.getBinaryStream());
		}
	}
	
	public static void addSheet(int key, int index, Blob content, Blob comments) 
			throws CellReaderException, SQLException, IOException {
		ReadContext ctx = ReadContext.get(key);
		ctx.reader.addSheet(new Sheet(index, content));
		if (comments != null) {
			((OOXCellReaderImpl) ctx.reader).addComments(index, comments.getBinaryStream());
		}
	}
	
	public static Array iterate(int key, int nrows) 
			throws SQLException, CellReaderException {
		
		OracleConnection conn = (OracleConnection) DriverManager.getConnection("jdbc:default:connection:");
		
		ReadContext ctx = ReadContext.get(key);
		
		int listSize = nrows * ctx.reader.getColumnCount();
		List<Struct> array = new ArrayList<Struct>(listSize);
		List<Row> rows = ctx.reader.readRows(nrows);
		for (Row r : rows) {
			for (Cell<?> c : r) {
				array.add(conn.createStruct("EXCELTABLECELL", c.getOraData(conn)));
			}
		}
		
		return conn.createOracleArray("EXCELTABLECELLLIST", array.toArray());
	}
	
	public static void terminate(int key) 
			throws CellReaderException, IOException {
		
		ReadContext ctx;
		try {
			ctx = (ReadContext) ContextManager.clearContext(key);
			ctx.reader.close();
		} catch (InvalidKeyException e) {
			throw new CellReaderException("Invalid context key", e);
		}
		
	}
	
}
