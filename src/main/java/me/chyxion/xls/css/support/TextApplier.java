package me.chyxion.xls.css.support;

import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import me.chyxion.xls.css.CssApplier;
import me.chyxion.xls.css.CssUtils;

/**
 * @version 0.1
 * @since 0.1
 * @author Shaun Chyxion <br />
 * chyxion@163.com <br />
 * Oct 24, 2014 5:21:30 PM
 */
public class TextApplier implements CssApplier {
	private static final Logger log = LoggerFactory.getLogger(TextApplier.class);

	/* (non-Javadoc)
	 * @see me.chyxion.xls.css.CssApplier#parse(java.util.Map)
	 */
    @Override
    public Map<String, String> parse(Map<String, String> style) {
    	log.debug("Parse Font Style.");
    	Map<String, String> mapRtn = new HashMap<String, String>();
    	// color
    	String color = CssUtils.processColor(style.get(COLOR));
    	if (StringUtils.isNotBlank(color)) {
    		log.debug("Text Color [{}] Found.", color);
    		if (Integer.parseInt(color.substring(1), 16) > 
    			Integer.parseInt("444444", 16)) {
    			mapRtn.put(COLOR, color);
    		}
    		else {
    			log.debug("Color [{}] Is Familiar To Black, Ignore.", color);
    		}
    	}
    	// font
    	String font = style.get(FONT);
    	if (StringUtils.isNotBlank(font)) {
    		log.debug("Parse Font Attr [{}].", font);
    		String[] ignoreStyle = new String[] {
    			"normal",
    			"small\\-caps",
    			"caption",
    			"icon",
    			"menu", 
    			"message\\-box",
    			"small\\-caption",
    			"status-bar",
    			// font weight
    			"[1-3]00"
    		};
    		StringBuffer sbFont = new StringBuffer(
    			font.replaceAll("^|\\s*" + StringUtils.join(ignoreStyle, "|") + "\\s+|$", " "));
    		log.debug("Font Attr [{}] After Process Ingore.", sbFont);
    		// style
    		Matcher m = Pattern.compile("(?:^|\\s+)(italic|oblique)(?:\\s+|$)")
    						.matcher(sbFont.toString());
    		if (m.find()) {
    			sbFont.setLength(0);
    			if (log.isDebugEnabled()) {
    				log.debug("Font Style [{}] Found.", m.group(1));
    			}
    			mapRtn.put(FONT_STYLE, ITALIC);
    			m.appendReplacement(sbFont, " ");
    			m.appendTail(sbFont);
    		}
    		// weight
    		m = Pattern.compile("(?:^|\\s+)(bold|[4-9]00)(?:\\s+|$)")
    				.matcher(sbFont.toString());
    		if (m.find()) {
    			sbFont.setLength(0);
    			if (log.isDebugEnabled()) {
    				log.debug("Font Weight [{}] Found.", m.group(1));
    			}
    			mapRtn.put(FONT_WEIGHT, BOLD);
    			m.appendReplacement(sbFont, " ");
    			m.appendTail(sbFont);
    		}
    		// size xx-small | x-small | small | medium | large | x-large | xx-large | 18px [/2]
    		m = Pattern.compile(
    				// before blank or start
    				new StringBuilder("(?:^|\\s+)")
    				// font size
    				.append("(xx-small|x-small|small|medium|large|x-large|xx-large|")
    				.append("(?:")
    				.append(PATTERN_LENGTH)
    				.append("))")
    				// line height
    				.append("(?:\\s*\\/\\s*(")
    				.append(PATTERN_LENGTH)
    				.append("))?")
    				// after blank or end
    				.append("(?:\\s+|$)")
    				.toString())
    				.matcher(sbFont.toString());
    		if (m.find()) {
    			sbFont.setLength(0);
    			log.debug("Font Size[/line-height] [{}] Found.", m.group());
    			String fontSize = m.group(1);
    			if (StringUtils.isNotBlank(fontSize)) {
    				fontSize = StringUtils.deleteWhitespace(fontSize);
    				log.debug("Font Size [{}].", fontSize);
    				if (fontSize.matches(PATTERN_LENGTH)) {
    					mapRtn.put(FONT_SIZE, fontSize);
    				}
    				else {
    					log.warn("Font Size [{}] Not Suppoted, Ignore.", fontSize);
    				}
    			}
    			String lineHeight = m.group(2);
    			if (StringUtils.isNotBlank(lineHeight)) {
    				lineHeight = StringUtils.deleteWhitespace(lineHeight);
    				log.debug("Line Height [{}].", lineHeight);
    				mapRtn.put(LINE_HEIGHT, lineHeight);
    			}
    			m.appendReplacement(sbFont, " ");
    			m.appendTail(sbFont);
    		}
    		// font family
    		if (sbFont.length() > 0) {
    			log.debug("Font Families [{}].", sbFont);
    			// trim & remove '"
    			String fontFamily = sbFont.toString()
    					.split("\\s*,\\s*")[0].trim().replaceAll("'|\"", "");
    			log.debug("Use First Font Family [{}].", fontFamily);
    			mapRtn.put(FONT_FAMILY, fontFamily);
    		}
    	}
	    return mapRtn;
    }

	/* (non-Javadoc)
	 * @see me.chyxion.xls.css.CssApplier#apply(org.apache.poi.hssf.usermodel.HSSFCell, org.apache.poi.hssf.usermodel.HSSFCellStyle, java.util.Map)
	 */
    @Override
    public void apply(HSSFCell cell, HSSFCellStyle cellStyle, Map<String, String> style) {
    	HSSFWorkbook workBook = cell.getSheet().getWorkbook();
    	HSSFFont font = null;
    	if (ITALIC.equals(style.get(FONT_STYLE))) {
    		font = getFont(cell, font);
    		font.setItalic(true);
    	}
    	int fontSize = CssUtils.getInt(style.get(FONT_SIZE));
    	if (fontSize > 0) {
    		font = getFont(cell, font);
    		font.setFontHeightInPoints((short) fontSize);
    	}
    	if (BOLD.equals(style.get(FONT_WEIGHT))) {
    		font = getFont(cell, font);
    		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
    	}
    	String fontFamily = style.get(FONT_FAMILY);
    	if (StringUtils.isNotBlank(fontFamily)) {
    		font = getFont(cell, font);
    		font.setFontName(fontFamily);
    	}
    	HSSFColor color = CssUtils.parseColor(workBook, style.get(COLOR));
    	if (color != null) {
    		font = getFont(cell, font);
    		font.setColor(color.getIndex());
    	}
    	if (font != null) {
    		cellStyle.setFont(font);
    	}
    }

    HSSFFont getFont(HSSFCell cell, HSSFFont font) {
    	if (font == null) {
    		font = cell.getSheet().getWorkbook().createFont();
    	}
    	return font;
    }
}
