package me.chyxion.xls;

import org.junit.Test;
import java.io.FileOutputStream;
import java.nio.charset.StandardCharsets;

/**
 * @author Shaun Chyxion
 * @date Oct 24, 2014 2:07:51 PM
 */
public class TestDriver {

	@Test
	public void run() throws Exception {
		TableToXls.process(getClass().getResourceAsStream("/sample.html"),
				StandardCharsets.UTF_8, "", new FileOutputStream("target/data.xlsx"));
	}
}
