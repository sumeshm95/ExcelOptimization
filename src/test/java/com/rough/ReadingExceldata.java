package com.rough;

import com.utilityes.Constants;
import com.utilityes.ExcelReader;

public class ReadingExceldata {
	
	public static void main(String[] args) {
		ExcelReader excel= new ExcelReader("C:\\Users\\smohandass\\eclipse-workspace\\ExcelOptimization\\src\\test\\resources\\Excel\\TestCases.xlsx");
	int rowCount = excel.getRowCount(Constants.sheetname);
	//find the test case starts
	String testName="openAccount" ;
int testCaseRowNum=1;
for(testCaseRowNum=1;testCaseRowNum<=rowCount;testCaseRowNum++) {
	String cellData = excel.getCellData(Constants.sheetname, 0, testCaseRowNum);
	if(cellData.equalsIgnoreCase(testName)) {
		break;
	}
}
	//System.out.println(testCaseRowNum);
	
	//checking total rows in test case
		int dataStartRowNo = testCaseRowNum +2;
		int row=0;
		while(!excel.getCellData(Constants.sheetname, 0, dataStartRowNo+row).equals("")) {
			row++;
		}
	//System.out.println(row);
	//checking total columns in test case
		
		  int StartColumnNo=testCaseRowNum+1; int colm=0; while
		  (!excel.getCellData(Constants.sheetname, colm, StartColumnNo).equals("")) {
		  
		  
			colm++;
		} /* System.out.println(colm); */
		 
	

	
	//printing whole test case data
	
	for(int rnum=dataStartRowNo;rnum<(dataStartRowNo+row);rnum++) {
		for(int cnum=0;cnum<colm;cnum++) {
			System.out.println(excel.getCellData(Constants.sheetname, cnum, rnum));
		}
	}
	
	
	

}
}
