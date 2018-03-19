package com.excel.template;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class RowCopyBack {
	public static void copyRow(XSSFRow srcRow, XSSFRow newRow) {
		List<CellStyle> styleMap = new ArrayList<>();
		newRow.setHeight(srcRow.getHeight());
		int j = srcRow.getFirstCellNum();
		if (j < 0) {
			j = 0;
		}
		for (; j <= srcRow.getLastCellNum(); j++) {
			XSSFCell oldCell = srcRow.getCell(j);
			XSSFCell newCell = newRow.getCell(j);
			if (oldCell != null) {
				if (newCell == null) {
					newCell = newRow.createCell(j);
				}
				copyCell(oldCell, newCell, styleMap);
			}
		}
		System.out.println("');

	}

	private static void copyCell(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
		if (styleList != null) {
			if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
				newCell.setCellStyle(oldCell.getCellStyle());
			} else {
				DataFormat newDataFormat = newCell.getSheet().getWorkbook().createDataFormat();

				CellStyle newCellStyle = getSameCellStyle(oldCell, newCell, styleList);
				if (newCellStyle == null) {
					Font oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
					Font newFont = newCell.getSheet().getWorkbook().findFont(oldFont.getBold(), oldFont.getColor(),
							oldFont.getFontHeight(), oldFont.getFontName(), oldFont.getItalic(), oldFont.getStrikeout(),
							oldFont.getTypeOffset(), oldFont.getUnderline());
					if (newFont == null) {
						newFont = newCell.getSheet().getWorkbook().createFont();
						newFont.setBold(oldFont.getBold());
						newFont.setColor(oldFont.getColor());
						newFont.setFontHeight(oldFont.getFontHeight());
						newFont.setFontName(oldFont.getFontName());
						newFont.setItalic(oldFont.getItalic());
						newFont.setStrikeout(oldFont.getStrikeout());
						newFont.setTypeOffset(oldFont.getTypeOffset());
						newFont.setUnderline(oldFont.getUnderline());
						newFont.setCharSet(oldFont.getCharSet());
					}

					short newFormat = newDataFormat.getFormat(oldCell.getCellStyle().getDataFormatString());
					newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
					newCellStyle.setFont(newFont);
					newCellStyle.setDataFormat(newFormat);
					newCellStyle.setHidden(oldCell.getCellStyle().getHidden());
					newCellStyle.setLocked(oldCell.getCellStyle().getLocked());
					newCellStyle.setWrapText(oldCell.getCellStyle().getWrapText());
					newCellStyle.setBottomBorderColor(oldCell.getCellStyle().getBottomBorderColor());
					newCellStyle.setFillBackgroundColor(oldCell.getCellStyle().getFillBackgroundColor());
					newCellStyle.setFillForegroundColor(oldCell.getCellStyle().getFillForegroundColor());
					newCellStyle.setIndention(oldCell.getCellStyle().getIndention());
					newCellStyle.setLeftBorderColor(oldCell.getCellStyle().getLeftBorderColor());
					newCellStyle.setRightBorderColor(oldCell.getCellStyle().getRightBorderColor());
					newCellStyle.setRotation(oldCell.getCellStyle().getRotation());
					newCellStyle.setTopBorderColor(oldCell.getCellStyle().getTopBorderColor());
					styleList.add(newCellStyle);
				}
				newCell.setCellStyle(newCellStyle);
			}
		}
		switch (oldCell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
			newCell.setCellValue(oldCell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_NUMERIC:
			newCell.setCellValue(oldCell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_BLANK:
			newCell.setCellType(Cell.CELL_TYPE_BLANK);
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			newCell.setCellValue(oldCell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR:
			newCell.setCellErrorValue(oldCell.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA:
			newCell.setCellFormula(oldCell.getCellFormula());

			break;
		default:
			break;
		}
	}

	private static CellStyle getSameCellStyle(Cell oldCell, Cell newCell, List<CellStyle> styleList) {
		CellStyle styleToFind = oldCell.getCellStyle();
		CellStyle currentCellStyle = null;
		CellStyle returnCellStyle = null;
		Iterator<CellStyle> iterator = styleList.iterator();
		Font oldFont = null;
		Font newFont = null;
		while (iterator.hasNext() && returnCellStyle == null) {
			currentCellStyle = iterator.next();

			if (currentCellStyle.getAlignment() != styleToFind.getAlignment()) {
				continue;
			}
			if (currentCellStyle.getHidden() != styleToFind.getHidden()) {
				continue;
			}
			if (currentCellStyle.getLocked() != styleToFind.getLocked()) {
				continue;
			}
			if (currentCellStyle.getWrapText() != styleToFind.getWrapText()) {
				continue;
			}
			if (currentCellStyle.getBorderBottom() != styleToFind.getBorderBottom()) {
				continue;
			}
			if (currentCellStyle.getBorderLeft() != styleToFind.getBorderLeft()) {
				continue;
			}
			if (currentCellStyle.getBorderRight() != styleToFind.getBorderRight()) {
				continue;
			}
			if (currentCellStyle.getBorderTop() != styleToFind.getBorderTop()) {
				continue;
			}
			if (currentCellStyle.getBottomBorderColor() != styleToFind.getBottomBorderColor()) {
				continue;
			}
			if (currentCellStyle.getFillBackgroundColor() != styleToFind.getFillBackgroundColor()) {
				continue;
			}
			if (currentCellStyle.getFillForegroundColor() != styleToFind.getFillForegroundColor()) {
				continue;
			}
			if (currentCellStyle.getFillPattern() != styleToFind.getFillPattern()) {
				continue;
			}
			if (currentCellStyle.getIndention() != styleToFind.getIndention()) {
				continue;
			}
			if (currentCellStyle.getLeftBorderColor() != styleToFind.getLeftBorderColor()) {
				continue;
			}
			if (currentCellStyle.getRightBorderColor() != styleToFind.getRightBorderColor()) {
				continue;
			}
			if (currentCellStyle.getRotation() != styleToFind.getRotation()) {
				continue;
			}
			if (currentCellStyle.getTopBorderColor() != styleToFind.getTopBorderColor()) {
				continue;
			}
			if (currentCellStyle.getVerticalAlignment() != styleToFind.getVerticalAlignment()) {
				continue;
			}

			oldFont = oldCell.getSheet().getWorkbook().getFontAt(oldCell.getCellStyle().getFontIndex());
			newFont = newCell.getSheet().getWorkbook().getFontAt(currentCellStyle.getFontIndex());

			if (newFont.getBold() == oldFont.getBold()) {
				continue;
			}
			if (newFont.getColor() == oldFont.getColor()) {
				continue;
			}
			if (newFont.getFontHeight() == oldFont.getFontHeight()) {
				continue;
			}
			if (newFont.getFontName() == oldFont.getFontName()) {
				continue;
			}
			if (newFont.getItalic() == oldFont.getItalic()) {
				continue;
			}
			if (newFont.getStrikeout() == oldFont.getStrikeout()) {
				continue;
			}
			if (newFont.getTypeOffset() == oldFont.getTypeOffset()) {
				continue;
			}
			if (newFont.getUnderline() == oldFont.getUnderline()) {
				continue;
			}
			if (newFont.getCharSet() == oldFont.getCharSet()) {
				continue;
			}
			if (oldCell.getCellStyle().getDataFormatString().equals(currentCellStyle.getDataFormatString())) {
				continue;
			}

			returnCellStyle = currentCellStyle;
		}
		return returnCellStyle;
	}

}
