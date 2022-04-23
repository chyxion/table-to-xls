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
 * @date Oct 24, 2014 5:18:57 PM
 */
public class HeightApplier implements CssApplier {

	/**
	 * {@inheritDoc}
	 */
	@Override
	public Map<String, String> parse(final Map<String, String> style) {
    	val mapRtn = new HashMap<String, String>();
    	val height = style.get(HEIGHT);
    	if (CssUtils.isNum(height)) {
    		mapRtn.put(HEIGHT, height);
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
    	val height = Math.round(CssUtils.getInt(style.get(HEIGHT)) * 255 / 12.75F);
    	val row = cell.getRow();
    	if (height > row.getHeight()) {
    		row.setHeight((short) height);
    	}
    }
}
