package me.chyxion.xls.css.support;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;

import me.chyxion.xls.css.CssApplier;
import me.chyxion.xls.css.CssUtils;

/**
 * @version 0.0.1
 * @since 0.0.1
 * @author Shaun Chyxion <br />
 * chyxion@163.com <br />
 * Oct 24, 2014 5:14:22 PM
 */
public class WidthApplier implements CssApplier {

	/* (non-Javadoc)
	 * @see me.chyxion.xls.css.CssApplier#parse(java.util.Map)
	 */
    @Override
    public Map<String, String> parse(Map<String, String> style) {
    	Map<String, String> mapRtn = new HashMap<String, String>();
    	String width = style.get(WIDTH);
    	if (CssUtils.isNum(width)) {
    		mapRtn.put(WIDTH, width);
    	}
	    return mapRtn;
    }

	/* (non-Javadoc)
	 * @see me.chyxion.xls.css.CssApplier#apply(org.apache.poi.hssf.usermodel.HSSFCell, org.apache.poi.hssf.usermodel.HSSFCellStyle, java.util.Map)
	 */
    @Override
    public void apply(HSSFCell cell, HSSFCellStyle cellStyle, Map<String, String> style) {
    	int width = Math.round(CssUtils.getInt(style.get(WIDTH)) * 2048 / 8.43F);
    	HSSFSheet sheet = cell.getSheet();
    	int colIndex = cell.getColumnIndex();
    	if (width > sheet.getColumnWidth(colIndex)) {
    		if (width > 255 * 256) {
    			width = 255 * 256;
    		}
    		sheet.setColumnWidth(colIndex, width);
    	}
    }
}
