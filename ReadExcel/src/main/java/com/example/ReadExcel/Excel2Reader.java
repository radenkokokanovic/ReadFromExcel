package com.example.ReadExcel;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class Excel2Reader {

	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
		  final String SAMPLE_XLSX_FILE_PATH = "BxSablon.xlsx";

		   

		        // Creating a Workbook from an Excel file (.xls or .xlsx)
		        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

		        // Retrieving the number of sheets in the Workbook
		        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

		        /*
		           =============================================================
		           			Iterating over all the sheets in the workbook 
		           =============================================================
		        */

		 
		        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
		        workbook.forEach(sheet -> {
		            System.out.println("=> " + sheet.getSheetName());
		        });

		        /*
		           ==================================================================
		           			Iterating over all the rows and columns in a Sheet 
		           ==================================================================
		        */

		        // Getting the Sheet at index zero
		        Sheet sheet = workbook.getSheetAt(0);

		        // Create a DataFormatter to format and get each cell's value as String
		        DataFormatter dataFormatter = new DataFormatter();

		    
		      int i=1;
		        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
		        sheet.forEach(row -> {
		        	
		            row.forEach(cell -> {
		            	
		                printCellValue(cell);
		            });
		            System.out.println();
		           
		        });

		        // Closing the workbook
		        workbook.close();
		    }

		    private static void printCellValue(Cell cell) {
		        switch (cell.getCellTypeEnum()) {
		            case BOOLEAN:
		                System.out.print(cell.getBooleanCellValue());
		                break;
		            case STRING:
		                System.out.print(cell.getRichStringCellValue().getString());
		                break;
		            case NUMERIC:
		                if (DateUtil.isCellDateFormatted(cell)) {
		                    System.out.print(cell.getDateCellValue());
		                } else {
		                    System.out.print(cell.getNumericCellValue());
		                }
		                break;
		            case FORMULA:
		                System.out.print(cell.getCellFormula());
		                break;
		            case BLANK:
		                System.out.print("");
		                break;
		            default:
		                System.out.print("");
		        }

		        System.out.print("\t");
		    }
}
