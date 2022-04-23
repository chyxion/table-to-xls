package me.chyxion.xls.css.support;

import lombok.val;
import java.util.Map;
import java.util.HashMap;
import me.chyxion.xls.css.CssApplier;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * @author Shaun Chyxion
 * @date Oct 24, 2014 2:29:17 PM
 */
public class AlignApplier implements CssApplier {

	/**
	 * {@inheritDoc}
	 */
	@Override
    public Map<String, String> parse(final Map<String, String> style) {
    	val mapRtn = new HashMap<String, String>();
    	String align = style.get(TEXT_ALIGN);
    	if (!ArrayUtils.contains(new String[] {LEFT, CENTER, RIGHT, JUSTIFY}, align)) {
    		align = LEFT;
    	}
    	mapRtn.put(TEXT_ALIGN, align);
    	align = style.get(VETICAL_ALIGN);
    	if (!ArrayUtils.contains(new String[] {TOP, MIDDLE, BOTTOM}, align)) {
    		align = MIDDLE;
    	}
    	mapRtn.put(VETICAL_ALIGN, align);
	    return mapRtn;
    }

    /**
     * {@inheritDoc}
     */
	@Override
	public void apply(final XSSFCell cell,
					  final XSSFCellStyle cellStyle,
					  final Map<String, String> style) {

    	// text align
    	val shAlign = style.get(TEXT_ALIGN);
		HorizontalAlignment sAlign = HorizontalAlignment.LEFT;
    	if (RIGHT.equals(shAlign)) {
    		sAlign = HorizontalAlignment.RIGHT;
    	}
    	else if (CENTER.equals(shAlign)) {
    		sAlign = HorizontalAlignment.CENTER;
    	}
    	else if (JUSTIFY.equals(shAlign)) {
    		sAlign = HorizontalAlignment.JUSTIFY;
    	}
    	cellStyle.setAlignment(sAlign);

    	// vertical align
    	val svAlign = style.get(VETICAL_ALIGN);

		VerticalAlignment vAlign = VerticalAlignment.CENTER;
    	if (TOP.equals(svAlign)) {
			vAlign = VerticalAlignment.TOP;
    	}
    	else if (BOTTOM.equals(svAlign)) {
    		vAlign = VerticalAlignment.BOTTOM;
    	}
    	else if (JUSTIFY.equals(svAlign)) {
    		vAlign = VerticalAlignment.JUSTIFY;
    	}
    	cellStyle.setVerticalAlignment(vAlign);
    }
}
