package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.time.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;

public class implementation {
  LocalDate date = LocalDate.now();
    int SNo;
    String projectNAme;
    int target;
    int activity;

    public void input() {

        Scanner sc = new Scanner(System.in);

        System.out.println("Enter serial number");
        SNo = sc.nextInt();

        System.out.println("Enter the project name");
        projectNAme = sc.next();

        System.out.println("Enter target value");
        target = sc.nextInt();

        System.out.println("Enter activity value");
        activity = sc.nextInt();     

    }

    public String toExcel() {

    	try {
        	FileInputStream inputStream = new FileInputStream("project.xlsx");
        	Workbook workbook = WorkbookFactory.create(inputStream);
        	Sheet sheet = workbook.getSheet("master");
        	Sheet newSheet = workbook.getSheet("details");
        	
//        	System.out.println(sheet.getLastRowNum());
        	
        	Row row = sheet.createRow(sheet.getLastRowNum()+1);
        	Row newRow = newSheet.createRow(sheet.getLastRowNum());
        	
//        	System.out.println(sheet.getLastRowNum());
        	
        	newRow.createCell(0).setCellValue(SNo);
        	newRow.createCell(1).setCellValue(projectNAme);
        	newRow.createCell(2).setCellValue(target);
        	newRow.createCell(3).setCellValue(activity);
        		
        	row.createCell(0).setCellValue(SNo);
        	row.createCell(1).setCellValue(projectNAme);
        	
        	if((target>1 && target <10) || (activity>1 && activity<10)) {
        		row.createCell(2).setCellValue("On Going");
        		newRow.createCell(4).setCellValue("On Going");
        	}
        	else if(target==10 && activity ==10) {
        		row.createCell(2).setCellValue("Completed");
        		newRow.createCell(4).setCellValue("Completed");
        	}
        	else if(activity==0 && target>1 &&target<9) {
        		row.createCell(2).setCellValue("Pending");
        		newRow.createCell(4).setCellValue("Pending");
        	}
        	else if(activity ==0 || target==0) {
        		row.createCell(2).setCellValue("Cancelled");
        		newRow.createCell(4).setCellValue("Cancelled");
        	}
        	else {
        		row.createCell(2).setCellValue("NA");
        		newRow.createCell(4).setCellValue("NA");
        	}
    	


//            System.out.println(row.getCell(2));

            FileOutputStream outputStream = new FileOutputStream("project.xlsx");
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            
            return "Successful";


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
//        } catch (InvalidFormatException e) {
//            e.printStackTrace();
//        }
        }
               
        return null;
        }

    }

