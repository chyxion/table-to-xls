package me.chyxion.xls.css.support;

import lombok.val;
import java.util.Map;
import java.util.HashMap;
import me.chyxion.xls.css.CssUtils;
import me.chyxion.xls.css.CssApplier;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * @author Shaun Chyxion
 * @date Oct 24, 2014 5:14:22 PM
 */
public class WidthApplier implements CssApplier {

	/**
	 * {@inheritDoc}
	 */
	@Override
	public Map<String, String> parse(final Map<String, String> style) {
    	val mapRtn = new HashMap<String, String>();
    	val width = style.get(WIDTH);
    	if (CssUtils.isNum(width)) {
    		mapRtn.put(WIDTH, width);
    	}
	    return mapRtn;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void apply(final XSSFCell cell,
					  final XSSFCellStyle cellStyle,
					  final Map<String, String> style) {
    	val width = Math.round(CssUtils.getInt(style.get(WIDTH)) * 2048 / 8.43F);
    	val sheet = cell.getSheet();
    	val colIndex = cell.getColumnIndex();
    	if (width > sheet.getColumnWidth(colIndex)) {
    		val maxWidth = 255 * 256;
    		sheet.setColumnWidth(colIndex, width > maxWidth ? maxWidth : width);
    	}
    }
}
