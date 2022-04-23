package me.chyxion.xls;

import lombok.val;
import java.util.Map;
import java.util.List;

import org.apache.poi.xssf.usermodel.*;
import org.jsoup.Jsoup;
import java.util.Arrays;
import org.slf4j.Logger;
import java.util.HashMap;
import lombok.SneakyThrows;
import java.io.IOException;
import java.io.InputStream;
import java.util.LinkedList;
import java.io.OutputStream;
import org.slf4j.LoggerFactory;
import org.jsoup.nodes.Element;
import java.nio.charset.Charset;
import me.chyxion.xls.css.CssApplier;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.commons.lang3.StringUtils;
import me.chyxion.xls.css.support.TextApplier;
import me.chyxion.xls.css.support.WidthApplier;
import org.apache.poi.ss.util.CellRangeAddress;
import me.chyxion.xls.css.support.AlignApplier;
import org.apache.poi.ss.usermodel.BorderStyle;
import me.chyxion.xls.css.support.BorderApplier;
import me.chyxion.xls.css.support.HeightApplier;
import me.chyxion.xls.css.support.BackgroundApplier;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

/**
 * @author Shaun Chyxion
 * @date Oct 24, 2014 2:09:02 PM
 */
public class TableToXls {
	private static final Logger log = 
		LoggerFactory.getLogger(TableToXls.class);
	private static final List<CssApplier> STYLE_APPLIERS = 
		new LinkedList<CssApplier>();
	// static init
	static {
		STYLE_APPLIERS.add(new AlignApplier());
		STYLE_APPLIERS.add(new BackgroundApplier());
		STYLE_APPLIERS.add(new WidthApplier());
		STYLE_APPLIERS.add(new HeightApplier());
		STYLE_APPLIERS.add(new BorderApplier());
		STYLE_APPLIERS.add(new TextApplier());
	}
	private XSSFWorkbook workBook = new XSSFWorkbook();
	private XSSFSheet sheet;
	private Map<String, Object> cellsOccupied = new HashMap<>(64);
	private Map<String, XSSFCellStyle> cellStyles = new HashMap<>(64);
	private XSSFCellStyle defaultCellStyle;
	private int maxRow = 0;
	// init
	{
		sheet = workBook.createSheet();
		defaultCellStyle = workBook.createCellStyle();
		defaultCellStyle.setWrapText(true);
		defaultCellStyle.setAlignment(HorizontalAlignment.CENTER);
		defaultCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		// border
		val black = HSSFColor.HSSFColorPredefined.BLACK.getIndex();
		val thin = BorderStyle.THIN;
		// top
		defaultCellStyle.setBorderTop(thin);
		defaultCellStyle.setTopBorderColor(black);
		// right
		defaultCellStyle.setBorderRight(thin);
		defaultCellStyle.setRightBorderColor(black);
		// bottom
		defaultCellStyle.setBorderBottom(thin);
		defaultCellStyle.setBottomBorderColor(black);
		// left
		defaultCellStyle.setBorderLeft(thin);
		defaultCellStyle.setLeftBorderColor(black);
	}

	/**
	 * @param inputStream html input stream
     *
	 * @param charset charset
	 * @param baseUrl html base url
	 * @param output output stream
	 */
	@SneakyThrows
	public static void process(
			final InputStream inputStream,
			final Charset charset,
			final String baseUrl,
			final OutputStream output) {
		new TableToXls().doProcess(inputStream, charset, baseUrl, output);
	}

	// --
	// private methods

	private void processTable(final Element table) {
		int rowIndex = 0;
		if (maxRow > 0) {
			// blank row
			maxRow += 2;
			rowIndex = maxRow;
		}

		log.info("Iterate table rows.");
		for (val row : table.select("tr")) {
			log.info("Parse table row [{}]. row index [{}].", row, rowIndex);
			int colIndex = 0;
			for (val td : row.select("td, th")) {
				// skip occupied cell
				while (cellsOccupied.get(rowIndex + "_" + colIndex) != null) {
					log.info("Cell [{}][{}] has been occupied, skip.", rowIndex, colIndex);
					++colIndex;
				}
				log.info("Parse col [{}], col index [{}].", td, colIndex);
				int rowSpan = 0;
				val strRowSpan = td.attr("rowspan");
				if (StringUtils.isNotBlank(strRowSpan) && 
						StringUtils.isNumeric(strRowSpan)) {
					log.info("Found row span [{}].", strRowSpan);
					rowSpan = Integer.parseInt(strRowSpan);
				}

				int colSpan = 0;
				val strColSpan = td.attr("colspan");
				if (StringUtils.isNotBlank(strColSpan) && 
						StringUtils.isNumeric(strColSpan)) {
					log.info("Found col span [{}].", strColSpan);
					colSpan = Integer.parseInt(strColSpan);
				}
				// col span & row span
				if (colSpan > 1 && rowSpan > 1) {
					spanRowAndCol(td, rowIndex, colIndex, rowSpan, colSpan);
					colIndex += colSpan;
				}
				// col span only
				else if (colSpan > 1) {
					spanCol(td, rowIndex, colIndex, colSpan);
					colIndex += colSpan;
				}
				// row span only
				else if (rowSpan > 1) {
					spanRow(td, rowIndex, colIndex, rowSpan);
					++colIndex;
				}
				// no span
				else {
					createCell(td, getOrCreateRow(rowIndex), colIndex).setCellValue(td.text());
					++colIndex;
				}
			}
			++rowIndex;
		}
	}

	private void doProcess(final InputStream inputStream,
						   final Charset charset,
						   final String baseUrl,
						   final OutputStream output) throws IOException {
		for (val table : Jsoup.parse(inputStream, charset.name(), baseUrl).select("table")) {
	        processTable(table);
        }
		workBook.write(output);
	}

    private void spanRow(Element td, int rowIndex, int colIndex, int rowSpan) {
    	log.info("Span row , from row [{}], span [{}].", rowIndex, rowSpan);
    	mergeRegion(rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex);
    	for (int i = 0; i < rowSpan; ++i) {
			val row = getOrCreateRow(rowIndex + i);
    		createCell(td, row, colIndex);
    		cellsOccupied.put((rowIndex + i) + "_" + colIndex, true);
    	}
    	getOrCreateRow(rowIndex).getCell(colIndex).setCellValue(td.text());
    }

    private void spanCol(Element td, int rowIndex, int colIndex, int colSpan) {
    	log.info("Span col, from col [{}], span [{}].", colIndex, colSpan);
    	mergeRegion(rowIndex, rowIndex, colIndex, colIndex + colSpan - 1);
    	val row = getOrCreateRow(rowIndex);
    	for (int i = 0; i < colSpan; ++i) {
    		createCell(td, row, colIndex + i);
    	}
    	row.getCell(colIndex).setCellValue(td.text());
    }

    private void spanRowAndCol(Element td, int rowIndex, int colIndex,
            int rowSpan, int colSpan) {
    	log.info("Span row and col, from row [{}], span [{}].", rowIndex, rowSpan);
    	log.info("From col [{}], span [{}].", colIndex, colSpan);
    	mergeRegion(rowIndex, rowIndex + rowSpan - 1, colIndex, colIndex + colSpan - 1);
    	for (int i = 0; i < rowSpan; ++i) {
			val row = getOrCreateRow(rowIndex + i);
    		for (int j = 0; j < colSpan; ++j) {
    			createCell(td, row, colIndex + j);
    			cellsOccupied.put((rowIndex + i) + "_" + (colIndex + j), true);
    		}
    	}
    	getOrCreateRow(rowIndex).getCell(colIndex).setCellValue(td.text());
    }

    private XSSFCell createCell(final Element td, final XSSFRow row, final int colIndex) {
		XSSFCell cell = row.getCell(colIndex);
    	if (cell == null) {
    		log.debug("Create cell [{}][{}].", row.getRowNum(), colIndex);
    		cell = row.createCell(colIndex);
    	}
    	return applyStyle(td, cell);
    }

    private XSSFCell applyStyle(final Element td, final XSSFCell cell) {
    	val style = td.attr(CssApplier.STYLE);
    	XSSFCellStyle cellStyle = null;
    	if (StringUtils.isNotBlank(style)) {
    		if (cellStyles.size() < 4000) {
				val mapStyle = parseStyle(style.trim());
				val mapStyleParsed = new HashMap<String, String>();
				for (CssApplier applier : STYLE_APPLIERS) {
					mapStyleParsed.putAll(applier.parse(mapStyle));
				}
				cellStyle = cellStyles.get(styleStr(mapStyleParsed));
				if (cellStyle == null) {
					log.debug("No Cell Style Found In Cache, Parse New Style.");
					cellStyle = workBook.createCellStyle();
					cellStyle.cloneStyleFrom(defaultCellStyle);
					for (val applier : STYLE_APPLIERS) {
						applier.apply(cell, cellStyle, mapStyleParsed);
					}
					// cache style
					cellStyles.put(styleStr(mapStyleParsed), cellStyle);
				}
    		}
    		else {
    			log.info("Custom cell style exceeds 4000, could not create new style, use default style.");
    			cellStyle = defaultCellStyle;
    		}
    	}
    	else {
    		log.debug("Use default cell style.");
    		cellStyle = defaultCellStyle;
    	}
    	cell.setCellStyle(cellStyle);
	    return cell;
    }

    private String styleStr(Map<String, String> style) {
    	log.debug("Build style string, style [{}].", style);
    	val sbStyle = new StringBuilder();
    	val keys = style.keySet().toArray();
    	Arrays.sort(keys);
    	for (val key : keys) {
    		sbStyle.append(key)
    		.append(':')
    		.append(style.get(key))
    		.append(';');
        }
    	log.debug("Style string result [{}].", sbStyle);
    	return sbStyle.toString();
    }

    private Map<String, String> parseStyle(String style) {
    	log.debug("Parse style string [{}] to map.", style);
    	val mapStyle = new HashMap<String, String>();

    	for (val s : style.split("\\s*;\\s*")) {
    		if (StringUtils.isNotBlank(s)) {
    			val ss = s.split("\\s*\\:\\s*");
    			if (ss.length == 2 &&
    					StringUtils.isNotBlank(ss[0]) &&
    					StringUtils.isNotBlank(ss[1])) {
    				val attrName = ss[0].toLowerCase();
    				String attrValue = ss[1];
    				// do not change font name
    				if (!CssApplier.FONT.equals(attrName) && 
    					!CssApplier.FONT_FAMILY.equals(attrName)) {
    					attrValue = attrValue.toLowerCase();
    				}
    				mapStyle.put(attrName, attrValue);
    			}
    		}
    	}
    	log.debug("Style map result [{}].", mapStyle);
	    return mapStyle;
    }

    private XSSFRow getOrCreateRow(int rowIndex) {
		XSSFRow row = sheet.getRow(rowIndex);
    	if (row == null) {
    		log.info("create new row [{}].", rowIndex);
    		row = sheet.createRow(rowIndex);
    		if (rowIndex > maxRow) {
    			maxRow = rowIndex;
    		}
    	}
	    return row;
    }

    private void mergeRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
    	log.debug("merge region, from row [{}], to [{}].", firstRow, lastRow);
    	log.debug("from col [{}], to [{}].", firstCol, lastCol);
    	sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }
}
