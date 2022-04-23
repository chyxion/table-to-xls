package me.chyxion.xls.css.support;

import lombok.extern.slf4j.Slf4j;
import lombok.val;
import java.util.Map;
import org.slf4j.Logger;
import java.util.HashMap;
import java.util.regex.Pattern;
import org.slf4j.LoggerFactory;
import me.chyxion.xls.css.CssUtils;
import me.chyxion.xls.css.CssApplier;
import org.apache.poi.ss.usermodel.Font;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * supports: <br>
 * color: name | #rgb | #rrggbb | rgb(r, g, b) <br>
 * text-decoration: underline; <br>
 * font-style: italic | oblique; <br>
 * font-weight:  bold | bolder | 700 | 800 | 900; <br>
 * font-size: length; length unit will be ignored, 
 * 	[xx-small|x-small|small|medium|large|x-large|xx-large] will be ignored. <br>
 * font：[[ font-style || font-variant || font-weight ]? font-size [/line-height]? font-family] 
 * | caption | icon | menu | message-box | small-caption | status-bar;
 * [font-variant, line-height, caption, icon, menu, message-box, small-caption, status-bar] will be ignored.
 *
 * @author Shaun Chyxion <br>
 * @date Oct 24, 2014 5:21:30 PM
 */
@Slf4j
public class TextApplier implements CssApplier {

	private static final String TEXT_DECORATION = "text-decoration";
	private static final String UNDERLINE = "underline"; 

	/**
	 * {@inheritDoc}
	 */
	@Override
    public Map<String, String> parse(final Map<String, String> style) {
    	log.debug("Parse font style.");
    	val mapRtn = new HashMap<String, String>(8);
    	// color
    	val color = CssUtils.processColor(style.get(COLOR));
    	if (StringUtils.isNotBlank(color)) {
    		log.debug("Text color [{}] found.", color);
    		mapRtn.put(COLOR, color);
    	}
    	// font
    	parseFontAttr(style, mapRtn);
    	// text text-decoration
    	if (UNDERLINE.equals(style.get(TEXT_DECORATION))) {
    		mapRtn.put(TEXT_DECORATION, UNDERLINE);
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

    	XSSFFont font = null;
    	if (ITALIC.equals(style.get(FONT_STYLE))) {
    		font = createFontIfNull(cell, font);
    		font.setItalic(true);
    	}

    	val fontSize = CssUtils.getInt(style.get(FONT_SIZE));
    	if (fontSize > 0) {
    		font = createFontIfNull(cell, font);
    		font.setFontHeightInPoints((short) fontSize);
    	}
    	if (BOLD.equals(style.get(FONT_WEIGHT))) {
    		font = createFontIfNull(cell, font);
    		font.setBold(true);
    	}

    	val fontFamily = style.get(FONT_FAMILY);
    	if (StringUtils.isNotBlank(fontFamily)) {
    		font = createFontIfNull(cell, font);
    		font.setFontName(fontFamily);
    	}

    	val color = CssUtils.parseColor(style.get(COLOR));
    	if (color != null) {
    		if (!CssUtils.isBlack(color)) {
    			font = createFontIfNull(cell, font);
    			font.setColor(color);
    		}
    		else {
				val removedColor = style.remove(COLOR);
				log.info("Text color [{}] is black or familiar to black, ignore.", removedColor);
    		}
    	}

    	// text-decoration
    	val textDecoration = style.get(TEXT_DECORATION);
    	if (UNDERLINE.equals(textDecoration)) {
    		font = createFontIfNull(cell, font);
    		font.setUnderline(Font.U_SINGLE);
    	}

    	if (font != null) {
    		cellStyle.setFont(font);
    	}
    }

    // --
    // private methods

    private Map<String, String> parseFontAttr(Map<String, String> style, Map<String, String> mapRtn) {
    	// font
    	val font = style.get(FONT);
    	if (StringUtils.isNotBlank(font) && 
    			!ArrayUtils.contains(new String[] {
    				"small-caps", "caption",
    				"icon", "menu", "message-box", 
    				"small-caption", "status-bar"
    			}, font)) {
    		log.debug("Parse font attr [{}].", font);
    		val ignoreStyles = new String[] {
    			"normal",
    			// font weight normal
    			"[1-3]00"
    		};
    		val sbFont = new StringBuffer(
    			font.replaceAll("^|\\s*" + StringUtils.join(ignoreStyles, "|") + "\\s+|$", " "));
    		log.debug("Font attr [{}] after process ignore.", sbFont);
    		// style
    		val matcherStyle = Pattern.compile("(?:^|\\s+)(italic|oblique)(?:\\s+|$)")
    						.matcher(sbFont.toString());
    		if (matcherStyle.find()) {
    			sbFont.setLength(0);
    			if (log.isDebugEnabled()) {
    				log.debug("Font style [{}] found.", matcherStyle.group(1));
    			}
    			mapRtn.put(FONT_STYLE, ITALIC);
    			matcherStyle.appendReplacement(sbFont, " ");
    			matcherStyle.appendTail(sbFont);
    		}

    		// weight
    		val matcherWeight = Pattern.compile("(?:^|\\s+)(bold(?:er)?|[7-9]00)(?:\\s+|$)")
    				.matcher(sbFont.toString());
    		if (matcherWeight.find()) {
    			sbFont.setLength(0);
    			if (log.isDebugEnabled()) {
    				log.debug("Font weight [{}](bold) found.", matcherWeight.group(1));
    			}
    			mapRtn.put(FONT_WEIGHT, BOLD);
    			matcherWeight.appendReplacement(sbFont, " ");
    			matcherWeight.appendTail(sbFont);
    		}

    		// size xx-small | x-small | small | medium | large | x-large | xx-large | 18px [/2]
    		val matcherSize = Pattern.compile(
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
    		if (matcherSize.find()) {
    			sbFont.setLength(0);
    			log.debug("Font size[/line-height] [{}] found.", matcherSize.group());
    			String fontSize = matcherSize.group(1);
    			if (StringUtils.isNotBlank(fontSize)) {
    				fontSize = StringUtils.deleteWhitespace(fontSize);
    				log.debug("Font size [{}].", fontSize);
    				if (fontSize.matches(PATTERN_LENGTH)) {
    					mapRtn.put(FONT_SIZE, fontSize);
    				}
    				else {
    					log.info("Font size [{}] not supported, ignore.", fontSize);
    				}
    			}

    			val lineHeight = matcherSize.group(2);
    			if (StringUtils.isNotBlank(lineHeight)) {
    				log.warn("Line height [{}] not supported, ignore.", lineHeight);
    			}
    			matcherSize.appendReplacement(sbFont, " ");
    			matcherSize.appendTail(sbFont);
    		}
    		// font family
    		if (sbFont.length() > 0) {
    			log.debug("Font families [{}].", sbFont);
    			// trim & remove '"
    			String fontFamily = sbFont.toString()
    					.split("\\s*,\\s*")[0].trim().replaceAll("'|\"", "");
    			log.debug("Use first font family [{}].", fontFamily);
    			mapRtn.put(FONT_FAMILY, fontFamily);
    		}
    	}

    	val fontStyle = style.get(FONT_STYLE);
    	if (ArrayUtils.contains(new String[] {ITALIC, "oblique"}, fontStyle)) {
    		log.debug("Font italic [{}] found.", fontStyle);
    		mapRtn.put(FONT_STYLE, ITALIC);
    	}

    	val fontWeight = style.get(FONT_WEIGHT);
    	if (StringUtils.isNotBlank(fontWeight) &&
    			Pattern.matches("^bold(?:er)?|[7-9]00$", fontWeight)) {
    		log.debug("Font weight [{}](bold) found.", fontWeight);
    		mapRtn.put(FONT_WEIGHT, BOLD);
    	}

    	val fontSize = style.get(FONT_SIZE);
    	if (CssUtils.isNum(fontSize)) {
    		log.debug("Font size [{}] found.", fontSize);
    		mapRtn.put(FONT_SIZE, fontSize);
    	}

    	val fontFamily = style.get(FONT_FAMILY);
    	if (StringUtils.isNotBlank(fontFamily)) {
    		log.debug("Font family [{}] found.", fontFamily);
    		mapRtn.put(FONT_FAMILY, fontFamily);
    	}
    	return mapRtn;
    }

    XSSFFont createFontIfNull(final XSSFCell cell, final XSSFFont font) {
    	if (font != null) {
			return font;
		}
		return cell.getSheet().getWorkbook().createFont();
	}
}
