package com.generic.util;

import java.io.FileNotFoundException;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Sheet;
import com.generic.page.LoadTestResult;
import com.generic.setup.ReleaseDetails;

public class LoadTestDataAnalysis {

	public static void main(String[] args) throws FileNotFoundException, IOException, InterruptedException {
		// TODO Auto-generated method stub

		String DatabaseName;
		String ReleasesTable;
		String CBIBrandsPagesTable;
		String BrandsLoadTestTable;

		DatabaseName = "CBILoadTest.db";
		ReleasesTable = "CBI_RELEASES";
		CBIBrandsPagesTable = "CBI_BRANDS_PAGES";
		BrandsLoadTestTable = "_LOAD_TEST_DETAILS";

		// Iterating over all the sheets in the workbook to be loaded to database tables
		loadResultsToTables(DatabaseName, ReleasesTable, CBIBrandsPagesTable, BrandsLoadTestTable);

		// Iterating over all the sheets in the workbook to do the data analysis
		// 1- Write R10 data to excel sheet
		writeReleaseResultToExcel("R10", DatabaseName, BrandsLoadTestTable);
		// 2- Do the data analysis (onload difference)
		dataAnaysis(DatabaseName, BrandsLoadTestTable, "R10", "R9");

	}

	public static void loadResultsToTables(String DatabaseName, String ReleasesTable, String CBIBrandsPagesTable,
			String BrandsLoadTestTable) {

		ExcelUtils ReadData;
		LoadTestResult details;
		String SheetName;

		ReadData = new ExcelUtils("LoadData"); // Load data pages and load test details to DB tables

		SqLiteUtils.insertCBIRelease(ReleasesTable, DatabaseName);

		System.out.println("Retrieving Sheets using for-each loop");
		for (Sheet sheet : ReadData.workbook) {
			System.out.println("##### Sheet Name => " + sheet.getSheetName());

			SheetName = sheet.getSheetName();

			System.out.println("##### Sheet Row Count => " + ReadData.getRowCount(SheetName));

				// Load all the brands pages then Load all the load test results (Read from
				// excel sheet and insert to db)
				for (int i = 3; i <= ReadData.getRowCount(SheetName); i++) {

					details = new LoadTestResult("LoadData", SheetName,0, i);
					
					SqLiteUtils.insertCBIBrandsPages(details, CBIBrandsPagesTable, DatabaseName);
					SqLiteUtils.insertLoadTestResult(details, SheetName + BrandsLoadTestTable, DatabaseName);
				}
		}

	}// loadResultsToTables

	public static void writeReleaseResultToExcel(String ReleaseName, String DatabaseName, String BrandsLoadTestTable)
			throws FileNotFoundException, IOException, InterruptedException {

		ExcelUtils WriteData;
		String SheetName;
		int SheetColNum;

		WriteData = new ExcelUtils("DataAnalysis"); // DataAnalysis mode means read the data from tables and load them
													// to excel sheet and do the data analysis

		System.out.println("Retrieving Sheets using for-each loop");
		for (Sheet sheet : WriteData.workbook) {
			System.out.println("##### Sheet Name => " + sheet.getSheetName());

			SheetName = sheet.getSheetName();

			System.out.println("##### Sheet Row Count  => " + WriteData.getRowCount(SheetName));

				SheetColNum = WriteData.getColumnCount(SheetName);

				System.out.println("~~~ Sheet column count => " + SheetColNum);

				SqLiteUtils ReleasResult = new SqLiteUtils();
				ReleasResult.getReleaseResults(ReleaseName, SheetName + BrandsLoadTestTable, DatabaseName);
				
				WriteData.setCellData(SheetName, SheetColNum + 1, 1, ReleaseDetails.getReleaseName(), true);
				WriteData.setCellData(SheetName, SheetColNum + 2, 1, "Run #" + ReleaseDetails.getRunNum(), true);
				WriteData.setCellData(SheetName, SheetColNum + 1, 2, "pageName", true);
				WriteData.setCellData(SheetName, SheetColNum + 2, 2, "pageHits", true);
				WriteData.setCellData(SheetName, SheetColNum + 3, 2, "onload", true);

				for (int i = 0; i <= ReleasResult.PageName.length - 1; i++) {
					if (ReleasResult.PageName[i] == null) {
						break;
					}
					
					WriteData.setCellData(SheetName, SheetColNum + 1, i + 3, ReleasResult.PageName[i], true);
					WriteData.setCellData(SheetName, SheetColNum + 2, i + 3, ReleasResult.PageHits[i], true);
					WriteData.setCellData(SheetName, SheetColNum + 3, i + 3, ReleasResult.Onload[i], true);
				}

				WriteData.writeExcelFile();

			
		}

	}// writeReleaseResultToExcel

	public static void dataAnaysis(String DatabaseName, String BrandsLoadTestTable, String CurrentReleaseName,
			String PreviousReleaseName) throws FileNotFoundException, IOException, InterruptedException {

		ExcelUtils WriteData;
		String SheetName;
		int SheetColNum;
		LoadTestResult ReleaseData;
		double PreviousOnload;
		double CurrentOnload;
		double FirstReleaseOnload;
		double OnloadDiff1;
		double OnloadDiff2;

		WriteData = new ExcelUtils("DataAnalysis"); // DataAnalysis mode means read the data from tables and load them
													// to excel sheet and do the data analysis

		System.out.println("Retrieving Sheets using for-each loop");
		for (Sheet sheet : WriteData.workbook) {
			System.out.println("##### Sheet Name => " + sheet.getSheetName());

			SheetName = sheet.getSheetName();

			System.out.println("##### Sheet Row Count  => " + WriteData.getRowCount(SheetName));

			if (SheetName.toUpperCase().equals("RY")) {

				SheetColNum = WriteData.getColumnCount(SheetName);
				System.out.println("~~~ Sheet column count => " + SheetColNum);
				WriteData.setCellData(SheetName, SheetColNum + 1, 2,
						"onload diff " + PreviousReleaseName + "/" + CurrentReleaseName, true);
				WriteData.setCellData(SheetName, SheetColNum + 3, 2, "onload diff R1/" + CurrentReleaseName, true);
				for (int i = 3; i <= WriteData.getRowCount(SheetName); i++) {

					ReleaseData = new LoadTestResult("DataAnalysis", SheetName,0, i);
					try {
						PreviousOnload = Double.valueOf(
								ReleaseData.getOnload(WriteData.getColNum(SheetName, PreviousReleaseName) + 2).trim());
					} catch (Exception e) {
						PreviousOnload = 0;
					}
					try {
						CurrentOnload = Double.valueOf(
								ReleaseData.getOnload(WriteData.getColNum(SheetName, CurrentReleaseName) + 2).trim());
					} catch (Exception e) {
						CurrentOnload = 0;
					}
					try {
						FirstReleaseOnload = Double
								.valueOf(ReleaseData.getOnload(WriteData.getColNum(SheetName, "R1") + 2).trim());
					} catch (Exception e) {
						FirstReleaseOnload = 0;
					}

					OnloadDiff1 = (PreviousOnload - CurrentOnload) / 1000;
					OnloadDiff2 = (FirstReleaseOnload - CurrentOnload) / 1000;

					WriteData.setCellData(SheetName, SheetColNum + 1, i, OnloadDiff1, true);
					WriteData.setCellData(SheetName, SheetColNum + 3, i, OnloadDiff2, true);
				}

				WriteData.writeExcelFile();

			}
		}

	}// dataAnaysis

}
