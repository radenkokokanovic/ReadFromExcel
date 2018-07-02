package com.example.ReadExcel;




import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class ReadExcel {
	
	 private static final String FILE_NAME = "BxSablon.xlsx";

	    public static void main(String[] args) {

	        try {

	            FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
	            Workbook workbook = new XSSFWorkbook(excelFile);
	            Sheet datatypeSheet = workbook.getSheetAt(0);
	            Iterator<Row> iterator = datatypeSheet.iterator();
	            
	            while (iterator.hasNext()) {

	                Row currentRow = iterator.next();
	                Iterator<Cell> cellIterator = currentRow.iterator();
	                
	              
	                
	                while (cellIterator.hasNext()) {

	                    Cell currentCell = cellIterator.next();
	                    //getCellTypeEnum shown as deprecated for version 3.15
	                    //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
	                    if (currentCell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.STRING) {
	                        System.out.println(currentCell.getStringCellValue() + "--");
	                    } else if (currentCell.getCellTypeEnum() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
	                        System.out.println(currentCell.getNumericCellValue() + "--");
	                    }

	                }
	                System.out.println();

	            }
	        } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }

	    }
	

}
