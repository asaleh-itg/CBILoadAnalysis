package com.generic.page;

import com.generic.util.*;

public class LoadTestResult {

	// variable
	public String mode;
	public String SheetName;
	public String Brand;
	public int releaseColNum;
	public int rowNum;
	public String pageName;
	public String pageHits;
	public String onload;

	// constructor
	public LoadTestResult(String mode, String SheetName,int releaseColNum, int rowNum) {
		this.mode = mode;
		this.SheetName = SheetName;
		this.releaseColNum = releaseColNum;
		this.rowNum = rowNum;
	}

	public LoadTestResult getLoadTestDetails() {
		LoadTestResult obj = new LoadTestResult(this.mode, this.SheetName,this.releaseColNum, this.rowNum);
		ExcelUtils data = new ExcelUtils(this.mode);
		obj.Brand = this.SheetName.toUpperCase();
		obj.pageName = data.getCellData(this.SheetName, this.releaseColNum, this.rowNum);
		obj.pageHits = data.getCellData(this.SheetName, this.releaseColNum+1, this.rowNum);
		obj.onload = data.getCellData(this.SheetName, this.releaseColNum+2, this.rowNum);

		return obj;
	}
	
	public String getOnload(int colNum) {
		String onload;
		ExcelUtils data = new ExcelUtils(this.mode);
		onload = data.getCellData(this.SheetName,colNum, this.rowNum);

		return onload;
	}

}
