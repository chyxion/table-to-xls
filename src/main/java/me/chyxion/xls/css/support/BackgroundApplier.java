package me.chyxion.xls.css.support;

import lombok.val;
import java.util.Map;
import java.util.HashMap;
import me.chyxion.xls.css.CssUtils;
import me.chyxion.xls.css.CssApplier;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * @author Shaun Chyxion
 * @date Oct 24, 2014 5:03:32 PM
 */
public class BackgroundApplier implements CssApplier {

	/**
	 * {@inheritDoc}
	 */
	@Override
    public Map<String, String> parse(final Map<String, String> style) {
    	val mapRtn = new HashMap<String, String>();
    	String bg = style.get(BACKGROUND);
    	String bgColor = null;
    	if (StringUtils.isNotBlank(bg)) {
    		for (val bgAttr : bg.split("(?<=\\)|\\w|%)\\s+(?=\\w)")) {
    			if ((bgColor = CssUtils.processColor(bgAttr)) != null) {
    				mapRtn.put(BACKGROUND_COLOR, bgColor);
    				break;
    			}
    		}
    	}

    	bg = style.get(BACKGROUND_COLOR);
    	if (StringUtils.isNotBlank(bg) &&
    			(bgColor = CssUtils.processColor(bg)) != null) {
    		mapRtn.put(BACKGROUND_COLOR, bgColor);
    	}
    	if (bgColor != null) {
    		bgColor = mapRtn.get(BACKGROUND_COLOR);
    		if ("#ffffff".equals(bgColor)) {
    			mapRtn.remove(BACKGROUND_COLOR);
    		}
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

    	val bgColor = style.get(BACKGROUND_COLOR);
    	if (StringUtils.isNotBlank(bgColor)) {
    		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    		cellStyle.setFillForegroundColor(CssUtils.parseColor(bgColor));
    	}
    }
}
