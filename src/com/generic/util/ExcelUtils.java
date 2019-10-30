package com.generic.util;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.*;

import com.generic.setup.ExcelSheetsPaths;
import com.generic.setup.SheetVariables;

import java.io.*;
import java.util.Calendar;

public class ExcelUtils {
	public String path;
	public FileInputStream fis = null;
	public FileOutputStream fileOut = null;
	public XSSFWorkbook workbook = null;
	private XSSFSheet sheet = null;
	private XSSFRow row = null;
	private XSSFCell cell = null;

	private XSSFCellStyle my_style = null;
	private XSSFFont my_font = null;

	public ExcelUtils(String mode) {

		if (mode.toLowerCase().equals("loaddata")) {
			this.path = ExcelSheetsPaths.getLoadDatasPath();
		} else {
			this.path = ExcelSheetsPaths.getDataAnalysisPath();
		}
		System.out.println("Inputs file: " + this.path);

		try {
			fis = new FileInputStream(path);
			ZipSecureFile.setMinInflateRatio(0);
			try {
				workbook = new XSSFWorkbook(fis);
				sheet = workbook.getSheetAt(0);
			} finally {
				fis.close();
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// returns the row count in a sheet
	public int getRowCount(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1)
			return 0;
		else {
			sheet = workbook.getSheetAt(index);
			int number = sheet.getLastRowNum() + 1;
			return number;
		}

	}

	// returns the data from a cell
	@SuppressWarnings("deprecation")
	public String getCellData(String sheetName, String colName, int rowNum) {
		try {
			if (rowNum <= 0)
				return "";

			int index = workbook.getSheetIndex(sheetName);
			int col_Num = -1;
			if (index == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(0);
			for (int i = 0; i < row.getLastCellNum(); i++) {
				if (row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
					col_Num = i;
			}
			if (col_Num == -1)
				return "";

			sheet = workbook.getSheetAt(index);
			row = sheet.getRow(rowNum - 1);
			if (row == null)
				return "";
			cell = row.getCell(col_Num);

			if (cell == null)
				return "";
			if (cell.getCellTypeEnum() == CellType.STRING)
				return ((cell.getStringCellValue().equals("null")) ? "" : cell.getStringCellValue());
			else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				String cellText = String.valueOf(cell.getNumericCellValue());
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in the form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + cal.get(Calendar.MONTH) + 1 + "/" + cellText;
				}

				return cellText;
			} else if (cell.getCellTypeEnum() == CellType.BLANK)
				return "";
			else
				return cell.getStringCellValue();

		} catch (Exception e) {

			e.printStackTrace();
			return "row " + rowNum + " or column " + colName + " does not exist in xls";
		}
	}

	// returns the data from a cell
	@SuppressWarnings("deprecation")
	public String getCellData(String sheetName, int colNum, int rowNum) {
		try {
			if (rowNum <= 0) {
				return "";
			}
			// logs.debug("getting data from sheet: " + sheetName);
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

			if (cell.getCellTypeEnum() == CellType.STRING)
				return cell.getStringCellValue();
			else if (cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA) {
				// Abeer change: use NumberToTextConverter to get the correct number value as
				// written in Excel sheet
				String cellText = NumberToTextConverter.toText(cell.getNumericCellValue());
				/*
				 * String cellText = String.valueOf(cell.getNumericCellValue()); if
				 * (cellText.contains(".0")) { cellText = cellText.replace(".0", ""); }
				 */
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// format in form of M/D/YY
					double d = cell.getNumericCellValue();

					Calendar cal = Calendar.getInstance();
					cal.setTime(HSSFDateUtil.getJavaDate(d));
					cellText = (String.valueOf(cal.get(Calendar.YEAR))).substring(2);
					cellText = cal.get(Calendar.MONTH) + 1 + "/" + cal.get(Calendar.DAY_OF_MONTH) + "/" + cellText;
				}

				return cellText;
			} else if (cell.getCellTypeEnum() == CellType.BOOLEAN) {
				cell.setCellType(CellType.STRING);
				return String.valueOf(cell.getBooleanCellValue());
			} else if (cell.getCellTypeEnum() == CellType.BLANK)
				return "";
			else {
				cell.setCellType(CellType.STRING);
				return String.valueOf(cell.getStringCellValue());
			}
		} catch (Exception e) {

			e.printStackTrace();
			return "row " + rowNum + " or column " + colNum + " does not exist  in xls";
		}
	}// getCellData()

	// Get ColName
	public int getColNum(String SheetName, String colName) {
		sheet = workbook.getSheet(SheetName);
		row = sheet.getRow(0);
		int colNum = -1;
		String CellValue;

		for (int i = 0; i < row.getLastCellNum(); i++) {
			try {
				CellValue = row.getCell(i).getStringCellValue().trim();
			} catch (Exception e) {
				CellValue = "X";
			}
			if (CellValue.equals(colName)) {
				colNum = i;
				break;
			}
		}
		return colNum;
	}// getColNum

	public int getRowNumber(String sheetName, String ColName, int guidCol) {
		sheet = workbook.getSheet(sheetName);
		int rowNum = -1;
		String CellValue;

		for (int rowNumber = 0; rowNumber < getRowCount(sheetName); rowNumber++) {
			row = sheet.getRow(rowNumber);
			try {
				CellValue = row.getCell(guidCol).getStringCellValue().trim().toLowerCase();
			} catch (Exception e) {
				CellValue = "X";
			}

			if (CellValue.equals(ColName.toLowerCase())) {
				rowNum = rowNumber;
				break;
			}

		}

		// logs.debug("Row number is: "+rowNum +" for rowName: "+rowName);
		return rowNum + 1;
	}// getRowNumber

	// returns true if data is set successfully else false - write plant valid / not
	// valid
	public boolean setOnload(String SheetName, String release_name, String data, int rowNumber, boolean valid) {
		int guideCol = getColNum(SheetName, release_name) + 2;
		return setValid(data, guideCol, rowNumber, valid);
	}

	// returns true if data is set successfully else false - write valid / not valid
	public boolean setValid(String data, int guideCol, int guidRow, boolean valid) {

		String[] Sheets = { SheetVariables.BD.toString(), SheetVariables.FG.toString(), SheetVariables.GR.toString(),
				SheetVariables.GH.toString(), SheetVariables.RY.toString() };
		boolean writeUsersatus = true;

		for (String sheet : Sheets) {
			writeUsersatus = writeUsersatus & setCellData(sheet, guideCol, guidRow, data, valid);
		}
		return writeUsersatus;

	}

	// returns true if data is set successfully else false
	public boolean setCellData(String sheetName, int colNumber, int rowNumber, String data, boolean valid) {
		try {

			int rowNum = rowNumber;
			int sheetIndex = workbook.getSheetIndex(sheetName);

			sheet = workbook.getSheetAt(sheetIndex);

			if (sheetIndex == -1 || rowNum == 0) {
				return false;
			} else {
				row = sheet.getRow(rowNum - 1);
				if (row == null) {
					row = sheet.createRow(rowNum - 1);
				} else {

					cell = row.getCell(colNumber);
					if (cell == null) {
						cell = row.createCell(colNumber);
					}

					my_style = workbook.createCellStyle();

					if ((data.startsWith("R") && (data.length() >= 2 && data.length() <= 3)) || data.equals("pageNname")
							|| data.equals("pageName") || data.equals("pageHits") || data.equals("onload")
							|| data.contains("onload diff ") || data.contains("Run #")) {
						my_font = workbook.createFont();
						my_font.setFamily(XSSFFont.DEFAULT_FONT_SIZE);
						my_font.setFontName("Arial");
						my_font.setBold(true);
						my_style.setFont(my_font);
					}
					else
					{
						my_font = workbook.createFont();
						my_font.setFamily(XSSFFont.DEFAULT_FONT_SIZE);
						my_font.setFontName("Arial");
						my_style.setFont(my_font);
					}

					cell.setCellStyle(my_style);
					cell.setCellValue(data);
				}

			}
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

		return true;
	}// setCell

	// returns true if data is set successfully else false
	public boolean setCellData(String sheetName, int colNumber, int rowNumber, double data, boolean valid) {
		try {

			int rowNum = rowNumber;
			int sheetIndex = workbook.getSheetIndex(sheetName);

			sheet = workbook.getSheetAt(sheetIndex);

			if (sheetIndex == -1 || rowNum == 0) {
				return false;
			} else {
				row = sheet.getRow(rowNum - 1);
				if (row == null) {
					row = sheet.createRow(rowNum - 1);
				} else {

					cell = row.getCell(colNumber);
					if (cell == null) {
						cell = row.createCell(colNumber);
					}

					my_style = workbook.createCellStyle();

					if (data >= 0) {
						my_font = workbook.createFont();
						my_font.setColor(IndexedColors.BLACK.getIndex());// my_font.setColor(new XSSFColor(new
																			// Color(50,73,38)));
						my_font.setFamily(XSSFFont.DEFAULT_FONT_SIZE);
						my_font.setFontName("Arial");
						my_style.setFont(my_font);
						my_style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
						my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}

					else if (data < 0 && data > -1) {
						my_font = workbook.createFont();
						my_font.setColor(IndexedColors.DARK_YELLOW.getIndex());// my_font.setColor(new XSSFColor(new
																				// Color(50,73,38)));
						my_font.setFamily(XSSFFont.DEFAULT_FONT_SIZE);
						my_font.setFontName("Arial");
						my_style.setFont(my_font);
						my_style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
						my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					} else if (data <= -1) {
						my_font = workbook.createFont();
						my_font.setColor(IndexedColors.RED.getIndex());// my_font.setColor(new XSSFColor(new
																		// Color(50,73,38)));
						my_font.setFamily(XSSFFont.DEFAULT_FONT_SIZE);
						my_font.setFontName("Arial");
						my_style.setFont(my_font);
						my_style.setFillForegroundColor(IndexedColors.ROSE.getIndex());
						my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					} else {
						my_font = workbook.createFont();
						my_font.setColor(IndexedColors.BLACK.getIndex());// my_font.setColor(new XSSFColor(new
																			// Color(50,73,38)));
						my_font.setFamily(XSSFFont.DEFAULT_FONT_SIZE);
						my_font.setFontName("Arial");
						my_style.setFont(my_font);
						my_style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
						my_style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
					}

					// maintain boarders
					my_style.setBorderBottom(BorderStyle.THIN);
					my_style.setBottomBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
					my_style.setBorderRight(BorderStyle.THIN);
					my_style.setRightBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
					my_style.setBorderLeft(BorderStyle.THIN);
					my_style.setLeftBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());
					my_style.setBorderTop(BorderStyle.THIN);
					my_style.setTopBorderColor(IndexedColors.GREY_25_PERCENT.getIndex());

					cell.setCellStyle(my_style);
					cell.setCellValue(data);
				}

			}
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		}

		return true;
	}// setCell

	public void writeExcelFile() throws FileNotFoundException, IOException, InterruptedException {
		try {
			System.out.println("Writing data to excel sheet");
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
			fileOut.flush();
			fileOut.close();
			System.out.println("Writing Done");
		} catch (Exception e) {
			System.out.println("closing the file failed");
			e.printStackTrace();
		} finally {
			System.out.println("try to clean file");
			if (fileOut != null) {
				try {
					Thread.sleep(4000);
					System.out.println("closing file again");
					fileOut.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	public void readExcelFile() throws FileNotFoundException, IOException {
		fis = new FileInputStream(path);
		workbook = new XSSFWorkbook(fis);
	}

	// find whether sheets exists
	public boolean isSheetExist(String sheetName) {
		int index = workbook.getSheetIndex(sheetName);
		if (index == -1) {
			index = workbook.getSheetIndex(sheetName.toUpperCase());
			if (index == -1)
				return false;
			else
				return true;
		} else
			return true;
	}

	// returns number of columns in a sheet
	public int getColumnCount(String sheetName) {
		// check if sheet exists
		if (!isSheetExist(sheetName))
			return -1;

		sheet = workbook.getSheet(sheetName);
		row = sheet.getRow(2);

		if (row == null)
			return -1;

		return row.getLastCellNum();

	}

	public int getCellRowNum(String sheetName, String colName, String cellValue) {

		for (int i = 1; i <= getRowCount(sheetName); i++) {
			if (getCellData(sheetName, colName, i).equalsIgnoreCase(cellValue)) {
				return i;
			}
		}
		return -1;

	}

}
