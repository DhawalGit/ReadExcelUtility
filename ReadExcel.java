package com.appliedselenium.utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Calendar;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {
	public String path;
	public FileInputStream fis;
	public FileOutputStream fos;
	private XSSFWorkbook workbook;
	private XSSFSheet sheet;
	private XSSFRow row;
	private XSSFCell cell;

	// Constructor accepts String value, which is the path where excel is stored
	public ReadExcel(String path) {
		fis = null;
		fos = null;
		workbook = null;
		sheet = null;
		row = null;
		cell = null;

		this.path = path;
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);
			fis.close();
		} catch (Exception e) {

			e.printStackTrace();
		}
	}

	// get total number of rows in the excel
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			return 0;
		}
		sheet = workbook.getSheetAt(index);
		return sheet.getLastRowNum() + 1;
	}

	// get particular cell value. Accepts sheet name, column name and row number as
	// argument
	public String getCellData(String sheetName, String colName, int rowNum) {
		try {
			if (rowNum <= 0) {
				return "";
			}
			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1) {
				return "";
			}

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {

				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}
			if (col_Num == -1) {
				return "";
			}
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);

			if (cell == null) {
				return "";
			}

			// Retreive the cell value based on its data type
			if (cell.getCellTypeEnum() == CellType.STRING)
				return cell.getStringCellValue();
			if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {

					double d = cell.getNumericCellValue();
					Calendar calendar = Calendar.getInstance();
					calendar.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = String.valueOf(calendar.get(1)).substring(2);
					cellText = String.valueOf(calendar.get(5)) + "/" + calendar.get(2) + '\001' + "/" + cellText;
				}

				return cellText;
			}
			if (cell.getCellTypeEnum() == CellType.BLANK) {
				return "";
			}
			return String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {

			e.printStackTrace();
			return "Row " + rowNum + " or column " + colName + " does not exist";
		}
	}

	// get cell value from the excel. Method accepts sheet name, column number and
	// row number as argument
	public String getCellData(String sheetName, int colNum, int rowNum) {
		try {
			if (rowNum <= 0) {
				return "";
			}
			int index = workbook.getSheetIndex(sheetName);

			if (index == -1) {
				return "";
			}
			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(colNum);
			if (cell == null)
				return "";
			XSSFFormulaEvaluator xssfFormulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();

			if (cell.getCellTypeEnum() == CellType.STRING) {
				return cell.getStringCellValue();
			}

			if (cell.getCellTypeEnum() == CellType.NUMERIC) {

				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {

					double num = cell.getNumericCellValue();

					Calendar calendar = Calendar.getInstance();
					calendar.setTime(HSSFDateUtil.getJavaDate(num));
					cellText = String.valueOf(calendar.get(1)).substring(2);
					cellText = String.valueOf(calendar.get(2) + 1) + "/" + calendar.get(5) + "/" + cellText;
				}

				return cellText;
			}
			if (cell.getCellTypeEnum() == CellType.BLANK) {
				return "";
			}
			if (cell.getCellTypeEnum() == CellType.FORMULA) {
				return xssfFormulaEvaluator.evaluate(cell).getStringValue();
			}

			return String.valueOf(cell.getBooleanCellValue());
		} catch (Exception e) {

			e.printStackTrace();
			return "Row " + rowNum + " or column " + colNum + " does not exist";
		}
	}

	// set value in a cell
	public boolean setCellData(String sheetName, String colName, int rowNum, String data) {
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			if (rowNum <= 0) {
				return false;
			}
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1) {
				return false;
			}
			sheet = workbook.getSheetAt(index);

			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {

				if (row.getCell(i).getStringCellValue().trim().equals(colName))
					colNum = i;
			}
			if (colNum == -1) {
				return false;
			}
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null) {
				row = sheet.createRow(rowNum - 1);
			}
			cell = row.getCell(colNum);
			if (cell == null) {
				cell = row.createCell(colNum);
			}
			cell.setCellValue(data);

			fos = new FileOutputStream(path);

			workbook.write(fos);

			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// set cell value #2
	public boolean setCellData(String sheetName, String colName, int rowNum, String data, String url) {
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);

			if (rowNum <= 0) {
				return false;
			}
			int index = workbook.getSheetIndex(sheetName);
			int colNum = -1;
			if (index == -1) {
				return false;
			}
			sheet = workbook.getSheetAt(index);

			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {

				if (row.getCell(i).getStringCellValue().trim().equalsIgnoreCase(colName)) {
					colNum = i;
				}
			}
			if (colNum == -1)
				return false;
			sheet.autoSizeColumn(colNum);
			row = sheet.getRow(rowNum - 1);
			if (row == null) {
				row = sheet.createRow(rowNum - 1);
			}
			cell = row.getCell(colNum);
			if (cell == null) {
				cell = row.createCell(colNum);
			}
			cell.setCellValue(data);
			XSSFCreationHelper createHelper = workbook.getCreationHelper();

			XSSFCellStyle xSSFCellStyle = workbook.createCellStyle();
			XSSFFont hlink_font = workbook.createFont();
			hlink_font.setUnderline((byte) 1);
			hlink_font.setColor(IndexedColors.BLUE.getIndex());
			xSSFCellStyle.setFont(hlink_font);

			XSSFHyperlink link = createHelper.createHyperlink(HyperlinkType.FILE);
			link.setAddress(url);
			cell.setHyperlink(link);
			cell.setCellStyle(xSSFCellStyle);

			fos = new FileOutputStream(path);
			workbook.write(fos);

			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// add a new sheet and return true if add is successful
	public boolean addSheet(String sheetname) {
		try {
			workbook.createSheet(sheetname);
			FileOutputStream fos = new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// remove an existing sheet and return true if deleted successfully
	public boolean removeSheet(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			return false;
		}

		try {
			workbook.removeSheetAt(index);
			FileOutputStream fos = new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// add a new column to the existing sheet and return true if column is added
	// successfully
	public boolean addColumn(String sheetName, String colName) {
		try {
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			int index = workbook.getSheetIndex(sheetName);
			if (index == -1) {
				return false;
			}
			XSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

			sheet = workbook.getSheetAt(index);

			row = sheet.getRow(0);
			if (row == null) {
				row = sheet.createRow(0);
			}
			if (row.getLastCellNum() == -1) {
				cell = row.createCell(0);
			} else {
				cell = row.createCell(row.getLastCellNum());
			}
			cell.setCellValue(colName);
			cell.setCellStyle(style);

			fos = new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// remove an existing column from a sheet and return true if column is removed
	public boolean removeColumn(String sheetName, int colNum) {
		try {
			if (!isSheetExist(sheetName))
				return false;
			fis = new FileInputStream(path);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheet(sheetName);
			XSSFCellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex());
			XSSFCreationHelper createHelper = workbook.getCreationHelper();
			style.setFillPattern(FillPatternType.NO_FILL);

			for (int i = 0; i < getRowCount(sheetName); i++) {
				row = sheet.getRow(i);
				if (row != null) {
					cell = row.getCell(colNum);
					if (cell != null) {
						cell.setCellStyle(style);
						row.removeCell(cell);
					}
				}
			}
			fos = new FileOutputStream(path);
			workbook.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	// return true if sheetName exist
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1) {
				return false;
			}
			return true;
		}
		return true;
	}

	// get total number of columns
	public int getColumnCount(String sheetName) {
		if (!isSheetExist(sheetName)) {
			return -1;
		}
		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(0);

		if (row == null) {
			return -1;
		}
		return row.getLastCellNum();
	}

	// add a hyperlink and return true if successfully added
	public boolean addHyperLink(String sheetName, String screenShotColName, String testCaseName, int index, String url,
			String message) {
		url = url.replace('\\', '/');
		if (!isSheetExist(sheetName)) {
			return false;
		}
		sheet = workbook.getSheet(sheetName);

		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, 0, i).equalsIgnoreCase(testCaseName)) {

				setCellData(sheetName, screenShotColName, i + index, message, url);

				break;
			}
		}
		return true;
	}

	// get the value of a particular cell
	public int getCellRowNum(String sheetName, String colName, String cellValue) {
		for (int i = 2; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, colName, i).equalsIgnoreCase(cellValue)) {
				return i;
			}
		}
		return -1;
	}
}
