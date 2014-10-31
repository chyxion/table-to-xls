package me.chyxion.xls;

import java.io.FileOutputStream;
import java.util.Scanner;
import org.junit.Test;

/**
 * @version 0.0.1
 * @since 0.0.1
 * @author Shaun Chyxion <br />
 * chyxion@163.com <br />
 * Oct 24, 2014 2:07:51 PM
 */
public class TestDriver {

	@Test
	public void run() throws Exception {
		StringBuilder html = new StringBuilder();
		Scanner s = new Scanner(getClass().getResourceAsStream("/sample.html"), "utf-8");
		while (s.hasNext()) {
			html.append(s.nextLine());
		}
		s.close();
		FileOutputStream fout = new FileOutputStream("data.xls");
		fout.write(TableToXls.process(html));
		fout.close();
	}
}
