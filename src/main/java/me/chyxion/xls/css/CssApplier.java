package me.chyxion.xls.css;

import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;

/**
 * @version 0.1
 * @since 0.1
 * @author Shaun Chyxion <br />
 * chyxion@163.com <br />
 * Oct 24, 2014 2:10:28 PM
 */
public interface CssApplier {
	// constants
	String PATTERN_LENGTH = "\\d*\\.?\\d+\\s*(?:em|ex|cm|mm|q|in|pt|pc|px)?";
	String STYLE = "style";
	// direction
	String TOP = "top";
	String RIGHT = "right";
	String BOTTOM = "bottom";
	String LEFT = "left";
	String WIDTH = "width";
	String HEIGHT = "height";
	String COLOR = "color";
	String BORDER = "border";
	String CENTER = "center";
	String JUSTIFY = "justify";
	String MIDDLE = "middle";
	String FONT = "font";
	String FONT_MS_YAHEI = "Microsoft YaHei";
	String FONT_STYLE = "font-style";
	String FONT_VARIANT = "font-variant";
	String FONT_WEIGHT = "font-weight";
	String FONT_SIZE = "font-size";
	String LINE_HEIGHT = "line-height";
	String FONT_FAMILY = "font-family";
	String ITALIC = "italic";
	String BOLD = "bold";
	String NORMAL = "normal";
	String DEFAULT_VALUE = "-1";
	String TEXT_ALIGN = "text-align";
	String VETICAL_ALIGN = "vertical-align";
	String BACKGROUND = "background";
	String BACKGROUND_COLOR = "background-color";
	String[] BORDER_STYLES = new String[] {
        // Default value Specifies no border	 
         "none",
        // The same as "none", except in border conflict resolution for table elements	 
         "hidden",
        // Specifies a dotted border	 
         "dotted",
        // Specifies a dashed border	 
         "dashed",
        // Specifies a solid border	 
         "solid",
        // Specifies a double border	 
         "double",
        // Specifies a 3D grooved border  
         "groove",
        // Specifies a 3D ridged border  
         "ridge",
        // Specifies a 3D inset border  
         "inset",
        // Specifies a 3D outset border  
         "outset",
        // Sets this property to its default value 
         "initial",
        // Inherits this property from its parent element 
         "inherit"
	};
	// methods

	/**
	 * parse css styles
	 * @param style
	 * @return
	 */
	Map<String, String> parse(Map<String, String> style);

	/**
	 * apply styles
	 * @param cell
	 * @param cellStyle
	 * @param style
	 */
	void apply(HSSFCell cell, HSSFCellStyle cellStyle, Map<String, String> style);
}
