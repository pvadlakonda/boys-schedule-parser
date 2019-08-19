package com.excelparser;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	private static final String READ_FILE = "Full Schedule Summer 2018_Final.xlsx";
	private static final String WRITE_FILE = "BoyzSummer2018Schedule.xlsx";
	static final String KEY = "Hyderabad Boyz";

	public static void main(String[] args) {
		try {
			FileInputStream excelFile = new FileInputStream(new File(READ_FILE));
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet datatypeSheet = workbook.getSheetAt(0);

			ExcelUtils.filterRows(datatypeSheet);
			ExcelUtils.removeEmptyRows(datatypeSheet);

			FileOutputStream fileOut = new FileOutputStream(WRITE_FILE);
			workbook.write(fileOut);
			fileOut.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
