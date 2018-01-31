package com.vyoms.whatsapp.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriteOperation {

	public static void main(String[] args) {

try {

    FileInputStream file = new FileInputStream(new File("C:\\Users\\Administrator\\Desktop\\BankAcoountData.xlsx"));
	
    XSSFWorkbook workbook = new XSSFWorkbook(file);

    XSSFSheet sheet = workbook.getSheetAt(0);

    Cell cell = null;

    //Update the value of cell

    cell = sheet.getRow(1).getCell(3);

    cell.setCellValue(20);

    cell = sheet.getRow(2).getCell(3);

    cell.setCellValue(30);

    cell = sheet.getRow(3).getCell(3);

    cell.setCellValue(50);

     

    file.close();

     
    FileOutputStream outFile =new FileOutputStream(new File("C:\\Users\\Administrator\\Desktop\\BankAcoountData.xlsx"));
    

    workbook.write(outFile);
    System.out.println("data written successfully");

    outFile.close();

     

} catch (FileNotFoundException e) {

    e.printStackTrace();

} catch (IOException e) {

    e.printStackTrace();

}
	}

}
