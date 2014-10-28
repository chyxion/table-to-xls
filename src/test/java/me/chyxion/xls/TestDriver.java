package me.chyxion.xls;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.commons.io.IOUtils;
import org.junit.Test;

/**
 * @version 0.1
 * @since 0.1
 * @author Shaun Chyxion <br />
 * chyxion@163.com <br />
 * Oct 24, 2014 2:07:51 PM
 */
public class TestDriver {

	@Test
	public void run() throws Exception {
		IOUtils.write(TableToXls.convert(IOUtils.toString(new FileInputStream("E:/table.html"))), 
				new FileOutputStream("E:/data.xls"));
	}

	@Test
	public void testSplit() {
		String[] ss = ":".split("\\s*\\:\\s*");
		System.err.println(ss.length);
		for (String s : ss) {
			System.err.println(s);
        }
		ss = "  a  :  b  ".split("\\s*\\:\\s*");
		System.err.println(ss.length);
		for (String s : ss) {
			System.err.println(s);
        }
	}
}
