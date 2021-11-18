package org.example;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class imp {

    public void save() {

        try {
        	FileInputStream inputStream = new FileInputStream("task1.xlsx");
        	Workbook workbook = WorkbookFactory.create(inputStream);
        	Sheet sheet = workbook.getSheet("master");
        	
//            XSSFWorkbook workbook = new XSSFWorkbook("task1.xlsx");
//            XSSFSheet sheet = workbook.getSheet("master");

            Row firstRow = sheet.createRow(0);
            firstRow.createCell(0,CellType.STRING).setCellValue("Sno");
            firstRow.createCell(1,CellType.STRING).setCellValue("Project Name");
            firstRow.createCell(2,CellType.STRING).setCellValue("Status");

            System.out.println("Current sheet row number: " + sheet.getLastRowNum());
            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.createCell(0, CellType.NUMERIC).setCellValue(1);
            row.createCell(1, CellType.STRING).setCellValue("First");
            row.createCell(2, CellType.STRING).setCellValue("Completed");

            System.out.println(sheet.getLastRowNum());

//            Row row1 = sheet.createRow(sheet.getLastRowNum()+2);
//            row.createCell(0, CellType.NUMERIC).setCellValue(2);
//            row.createCell(1, CellType.STRING).setCellValue("Second");
//            row.createCell(2, CellType.STRING).setCellValue("On going");

//            cell0.setCellValue(1);
//            cell1.setCellValue("First");
//            cell2.setCellValue("Completed");


//            System.out.println(sheet.getRow(2).getCell(2));
//
////            System.out.println(row.getLastCellNum());
//            System.out.println(row.getCell(2).getColumnIndex());
//            System.out.println(row.getCell(2).getStringCellValue());
//            System.out.println("New sheet Row number: " + sheet.getLastRowNum());

//            Iterator<Row> rowIterator = sheet.iterator();
//            Row rowNew = rowIterator.next();
//
//            System.out.println(rowNew.getRowNum());

            FileOutputStream outputStream = new FileOutputStream("task1.xlsx");
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            



        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        }
        }
    }
}