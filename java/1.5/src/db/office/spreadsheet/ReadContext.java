package db.office.spreadsheet;

import java.io.IOException;
import java.io.InputStream;
import java.sql.Array;
import java.sql.Blob;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Struct;
import java.util.ArrayList;
import java.util.List;

import oracle.CartridgeServices.ContextManager;
import oracle.CartridgeServices.CountException;
import oracle.CartridgeServices.InvalidKeyException;
import oracle.jdbc.OracleConnection;
import oracle.sql.STRUCT;
import oracle.sql.StructDescriptor;

public class ReadContext {
	
	public static int VC2_MAXSIZE;
	private static StructDescriptor cellTypeDesc;
	public CellReader reader;
	
	private static Connection getConnection() throws SQLException {
		return DriverManager.getConnection("jdbc:default:connection:");
	}
	
	public ReadContext(InputStream worksheet, InputStream sharedStrings, String columns, int firstRow, int lastRow)
			throws IOException, CellReaderException {

		this.reader = new CellReader(worksheet, sharedStrings, columns, firstRow, lastRow);

	}
	
	public static int initialize(Blob worksheet, Blob sharedStrings, String columns, int firstRow, int lastRow, int vc2MaxSize)
			throws IOException, CellReaderException, SQLException {
				
		VC2_MAXSIZE = vc2MaxSize;
		cellTypeDesc = StructDescriptor.createDescriptor("EXCELTABLECELL", getConnection());
		ReadContext ctx = new ReadContext(worksheet.getBinaryStream(), (sharedStrings!=null)?sharedStrings.getBinaryStream():null, columns, firstRow, lastRow);
		
		int key = 0;
		try {
			key = ContextManager.setContext(ctx);
		} catch (CountException e) {
			throw new CellReaderException("Error during context creation", e);
		}
		return key;

	}
	
	public static Array iterate(int key, int nrows) 
			throws SQLException, CellReaderException {
		
		OracleConnection conn = (OracleConnection) getConnection();
		
		ReadContext ctx;
		try {
			ctx = (ReadContext) ContextManager.getContext(key);
		} catch (InvalidKeyException e) {
			throw new CellReaderException("Invalid context key", e);
		}
		
		int listSize = nrows * ctx.reader.getColumnCount();
		List<Struct> array = new ArrayList<Struct>(listSize);
		List<Row> rows = ctx.reader.readRows(nrows, false, 0);
		for (Row r : rows) {
			for (Cell c : r) {
				//array.add(conn.createStruct("EXCELTABLECELL", c.getOraData(conn)));
				array.add(new STRUCT(cellTypeDesc, conn, c.getOraData(conn)));
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
