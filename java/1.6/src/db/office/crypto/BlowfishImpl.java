package db.office.crypto;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.InvalidAlgorithmParameterException;
import java.security.InvalidKeyException;
import java.security.NoSuchAlgorithmException;
import java.sql.SQLException;

import javax.crypto.Cipher;
import javax.crypto.CipherInputStream;
import javax.crypto.NoSuchPaddingException;
import javax.crypto.spec.IvParameterSpec;
import javax.crypto.spec.SecretKeySpec;

public class BlowfishImpl {

	public static void decrypt(java.sql.Blob input, byte[] key, byte[] iv, java.sql.Blob[] output) 
			throws NoSuchAlgorithmException, NoSuchPaddingException, InvalidKeyException, 
			InvalidAlgorithmParameterException, SQLException, IOException {

		SecretKeySpec keySpec = new SecretKeySpec(key,"Blowfish");
		IvParameterSpec ivSpec = new IvParameterSpec(iv);
		Cipher cipher = Cipher.getInstance("Blowfish/CFB/NoPadding");
		cipher.init(Cipher.DECRYPT_MODE, keySpec, ivSpec);
		InputStream is = new CipherInputStream(input.getBinaryStream(), cipher);
		OutputStream os = output[0].setBinaryStream(1L);
		byte[] buf = new byte[2048];
		int len;
		while ((len = is.read(buf)) != -1) {
			os.write(buf, 0, len);
		}
		os.close();
		is.close();
		
	}
	
}
