package me.chyxion.xls.css.support;

import lombok.val;
import java.util.Map;
import java.util.HashMap;
import lombok.extern.slf4j.Slf4j;
import me.chyxion.xls.css.CssUtils;
import java.util.function.BiConsumer;
import me.chyxion.xls.css.CssApplier;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * border[-[pos][-attr]]: [border-width] || [border-style] || [border-color]; <br>
 * border-style: none | hidden | dotted | dashed | solid | double
 *
 * @author Shaun Chyxion
 * @date Oct 24, 2014 5:21:51 PM
 */
@Slf4j
public class BorderApplier implements CssApplier {

	private static final String NONE = "none";
	private static final String HIDDEN = "hidden";
	private static final String SOLID = "solid";
	private static final String DOUBLE = "double";
	private static final String DOTTED = "dotted";
	private static final String DASHED = "dashed";
	// border styles
	private final static String[] BORDER_STYLES = new String[] {
        // Specifies no border	 
         NONE,
        // The same as "none", except in border conflict resolution for table elements
         HIDDEN,
        // Specifies a dotted border	 
         DOTTED,
        // Specifies a dashed border	 
         DASHED,
        // Specifies a solid border	 
         SOLID,
        // Specifies a double border	 
         DOUBLE
	};

	private static final Map<String, BiConsumer<XSSFCellStyle, XSSFColor>> CELL_BORDER_COLOR_SETTERS = new HashMap<>();
	private static final Map<String, BiConsumer<XSSFCellStyle, BorderStyle>> CELL_BORDER_STYLE_SETTERS = new HashMap<>();

	static {
		CELL_BORDER_COLOR_SETTERS.put(TOP, XSSFCellStyle::setTopBorderColor);
		CELL_BORDER_COLOR_SETTERS.put(RIGHT, XSSFCellStyle::setRightBorderColor);
		CELL_BORDER_COLOR_SETTERS.put(BOTTOM, XSSFCellStyle::setBottomBorderColor);
		CELL_BORDER_COLOR_SETTERS.put(LEFT, XSSFCellStyle::setLeftBorderColor);

		CELL_BORDER_STYLE_SETTERS.put(TOP, XSSFCellStyle::setBorderTop);
		CELL_BORDER_STYLE_SETTERS.put(RIGHT, XSSFCellStyle::setBorderRight);
		CELL_BORDER_STYLE_SETTERS.put(BOTTOM, XSSFCellStyle::setBorderBottom);
		CELL_BORDER_STYLE_SETTERS.put(LEFT, XSSFCellStyle::setBorderLeft);
	}

	/**
	 * {@inheritDoc}
	 */
	@Override
    public Map<String, String> parse(final Map<String, String> style) {
    	val mapRtn = new HashMap<String, String>();
    	for (String pos : new String[] {null, TOP, RIGHT, BOTTOM, LEFT}) {
    		// border[-attr]
    		if (pos == null) {
    			setBorderAttr(mapRtn, pos, style.get(BORDER));
    			setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + COLOR));
    			setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + WIDTH));
    			setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + STYLE));
    		}
    		// border-pos[-attr]
    		else {
    			setBorderAttr(mapRtn, pos, style.get(BORDER + "-" + pos));
    			for (String attr : new String[] {COLOR, WIDTH, STYLE}) {
    				val attrName = BORDER + "-" + pos + "-" + attr;
    				val attrValue = style.get(attrName);
    				if (StringUtils.isNotBlank(attrValue)) {
    					mapRtn.put(attrName, attrValue);
    				}
    			}
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

    	for (val pos : ALL_SIDES) {
    		// color
    		val colorAttr = BORDER + "-" + pos + "-" + COLOR;
    		val poiColor = CssUtils.parseColor(style.get(colorAttr));

    		if (poiColor != null) {
				CELL_BORDER_COLOR_SETTERS.get(pos).accept(cellStyle, poiColor);
    		}

			// border style
			val borderStyle = getBorderStyle(pos, style);
    		if (borderStyle != null) {
				CELL_BORDER_STYLE_SETTERS.get(pos).accept(cellStyle, borderStyle);
    		}
    	}
    }


	// --
    // private methods

	BorderStyle getBorderStyle(final String pos, final Map<String, String> style) {
		// width
		val width = CssUtils.getInt(style.get(BORDER + "-" + pos + "-" + WIDTH));
		val styleAttr = BORDER + "-" + pos + "-" + STYLE;
		val styleValue = style.get(styleAttr);

		// empty or solid
		if (StringUtils.isBlank(styleValue) || "solid".equals(styleValue)) {
			if (width > 2) {
				return BorderStyle.THICK;
			}
			if (width > 1) {
				return BorderStyle.MEDIUM;
			}
			return BorderStyle.THIN;
		}

		if (ArrayUtils.contains(new String[] {NONE, HIDDEN}, styleValue)) {
			return BorderStyle.NONE;
		}

		if (DOUBLE.equals(styleValue)) {
			return BorderStyle.DOUBLE;
		}

		if (DOTTED.equals(styleValue)) {
			return BorderStyle.DOTTED;
		}

		if (DASHED.equals(styleValue)) {
			if (width > 1) {
				return BorderStyle.MEDIUM_DASHED;
			}
			return BorderStyle.DASHED;
		}

		return null;
	}

    private void setBorderAttr(final Map<String, String> mapBorder,
							   final String pos,
							   final String value) {

    	if (StringUtils.isNotBlank(value)) {
    		String borderColor = null;
    		for (String borderAttr : value.split("\\s+")) {
    			if ((borderColor = CssUtils.processColor(borderAttr)) != null) {
    				setBorderAttr(mapBorder, pos, COLOR, borderColor);
    			}
    			else if (CssUtils.isNum(borderAttr)) {
    				setBorderAttr(mapBorder, pos, WIDTH, borderAttr);
    			}
    			else if (isStyle(borderAttr)) {
    				setBorderAttr(mapBorder, pos, STYLE, borderAttr);
    			}
    			else {
    				log.info("Border Attr [{}] Is Not Suppoted.", borderAttr);
    			}
    		}
    	}
    }

    private void setBorderAttr(final Map<String, String> mapBorder,
							   final String pos,
							   final String attr,
							   final String value) {

    	if (StringUtils.isNotBlank(pos)) {
    		mapBorder.put(BORDER + "-" + pos + "-" + attr, value);
    	}
    	else {
    		for (val side : ALL_SIDES) {
    			mapBorder.put(BORDER + "-" + side + "-" + attr, value);
    		}
    	}
    }
    
    private boolean isStyle(final String value) {
    	return ArrayUtils.contains(BORDER_STYLES, value);
    }
}
