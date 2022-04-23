package me.chyxion.xls.css;

import lombok.val;
import java.util.Map;
import java.util.HashMap;
import java.util.regex.Pattern;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.ss.usermodel.ExtendedColor;

/**
 * @author Shaun Chyxion
 * @date Oct 24, 2014 4:29:26 PM
 */
@Slf4j
public class CssUtils {
	// matches #rgb
	private static final String COLOR_PATTERN_VALUE_SHORT = 
			"^(#(?:[a-f]|\\d){3})$";
	// matches #rrggbb
	private static final String COLOR_PATTERN_VALUE_LONG = 
			"^(#(?:[a-f]|\\d{2}){3})$";
	// matches #rgb(r, g, b)
	private static final String COLOR_PATTERN_RGB = 
			"^(rgb\\s*\\(\\s*(.+)\\s*,\\s*(.+)\\s*,\\s*(.+)\\s*\\))$";
	// color name -> POI Color
	private static Map<String, Color> colors = new HashMap<>();

	// static init
	static {
		for (val colorPredefined : HSSFColor.HSSFColorPredefined.values()) {
			colors.put(colorName(colorPredefined.name()), convertHSSFColorToXSSFColor(colorPredefined.getColor()));
		}

		// light gray
		val colorLightgray = convertHSSFColorToXSSFColor(HSSFColor.HSSFColorPredefined.GREY_25_PERCENT.getColor());
		colors.put("lightgray", colorLightgray);
		colors.put("lightgrey", colorLightgray);
		// silver
		colors.put("silver", convertHSSFColorToXSSFColor(HSSFColor.HSSFColorPredefined.GREY_40_PERCENT.getColor()));
		// darkgray
		val colorDarkgray = convertHSSFColorToXSSFColor(HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getColor());
		colors.put("darkgray", colorDarkgray);
		colors.put("darkgrey", colorDarkgray);
		// gray
		val colorGray = convertHSSFColorToXSSFColor(HSSFColor.HSSFColorPredefined.GREY_80_PERCENT.getColor());
		colors.put("gray", colorGray);
		colors.put("grey", colorGray);
	}

	/**
	 * get color name
	 * @param color HSSFColor
	 * @return color name
	 */
    private static String colorName(final String color) {
	    return color.replace("_", "").toLowerCase();
    }
   
    /**
     * get int value of string
     * @param strValue string value
     * @return int value
     */
    public static int getInt(final String strValue) {
    	if (StringUtils.isNotBlank(strValue)) {
    		val m = Pattern.compile("^(\\d+)(?:\\w+|%)?$").matcher(strValue);
    		if (m.find()) {
    			return Integer.parseInt(m.group(1));
    		}
    	}
    	return 0;
    }
   
    /**
     * check number string 
     * @param strValue string
     * @return true if string is number
     */
    public static boolean isNum(final String strValue) {
    	return StringUtils.isNotBlank(strValue) && strValue.matches("^\\d+(\\w+|%)?$");
    }
   
    /**
     * process color
     * @param color color to process
     * @return color after process
     */
    public static String processColor(final String color) {
    	log.info("Process color [{}].", color);

    	if (StringUtils.isBlank(color)) {
    		return null;
		}

		// #rgb -> #rrggbb
		if (color.matches(COLOR_PATTERN_VALUE_SHORT)) {
			log.debug("Short Hex color [{}] found.", color);
			val sbColor = new StringBuffer();
			val m = Pattern.compile("([a-f]|\\d)").matcher(color);
			while (m.find()) {
				m.appendReplacement(sbColor, "$1$1");
			}
			val colorRtn = sbColor.toString();
			log.debug("Translate short HEX color [{}] to [{}].", color, colorRtn);
			return colorRtn;
		}

		// #rrggbb
		if (color.matches(COLOR_PATTERN_VALUE_LONG)) {
			log.debug("Hex color [{}] found, return.", color);
			return color;
		}

		// rgb(r, g, b)
		if (color.matches(COLOR_PATTERN_RGB)) {
			val m = Pattern.compile(COLOR_PATTERN_RGB).matcher(color);
			if (m.matches()) {
				log.debug("RGB color [{}] found.", color);
				val colorRtn = convertColor(calcColorValue(m.group(2)),
							calcColorValue(m.group(3)),
							calcColorValue(m.group(4)));
				log.debug("Translate RGB color [{}] to HEX [{}].", color, colorRtn);
				return colorRtn;
			}
		}

		val poiColor = getColor(color);
		// color name, red, green, ...
		if (poiColor != null) {
			log.debug("Color name [{}] found.", color);

			val rgb = convertToIntArray(((ExtendedColor) poiColor).getRGB());
			val colorRtn = convertColor(rgb[0], rgb[1], rgb[2]);
			log.debug("Translate color name [{}] to HEX [{}].", color, colorRtn);
			return colorRtn;
		}

    	return null;
    }

    /**
     * parse color
     * @param color string color
     * @return HSSFColor 
     */
    public static XSSFColor parseColor(final String color) {
    	if (StringUtils.isNotBlank(color)) {
    		val awtColor = java.awt.Color.decode(color);
    		if (awtColor != null) {
    			return new XSSFColor(new byte[]{(byte) awtColor.getRed(), (byte) awtColor.getGreen(), (byte) awtColor.getBlue()}, null);
    		}
    	}
    	return null;
    }

	/**
	 * if color is black
	 *
	 * @param color color
	 * @return true if color is black
	 */
	public static boolean isBlack(final XSSFColor color) {
		for (val b : color.getRGB()) {
			if (b != 0) {
				return false;
			}
		}
		return true;
	}

    // --
    // private methods

    private static Color getColor(String color) {
    	return colors.get(color.replace("_", ""));
    }

    private static String convertColor(final int r, final int g, final int b) {
    	return String.format("#%02x%02x%02x", r, g, b);
    }

    private static int calcColorValue(final String color) {
    	// matches 64 or 64%
		val m = Pattern.compile("^(\\d*\\.?\\d+)\\s*(%)?$").matcher(color);
		if (m.matches()) {
			// % not found
			if (m.group(2) == null) {
				return Math.round(Float.parseFloat(m.group(1))) % 256;
			}
			return Math.round(Float.parseFloat(m.group(1)) * 255 / 100) % 256;
		}
		return 0;
    }

    static XSSFColor convertHSSFColorToXSSFColor(final HSSFColor color) {
		val rgb = color.getTriplet();
		return new XSSFColor(new byte[]{(byte) rgb[0], (byte) rgb[1], (byte) rgb[2]}, null);
	}

	static int[] convertToIntArray(byte[] input) {
		val ret = new int[input.length];
		int i = 0;
		for (val b : input) {
			// Range 0 to 255, not -128 to 127
			ret[i++] = b & 0xff;
		}
		return ret;
	}
}
