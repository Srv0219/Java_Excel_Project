package com.java.FileIo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;
import java.util.Scanner;

public class SearchFile {
    public void searchCells(XSSFSheet sheet, XSSFWorkbook wb) {
        Scanner sc = new Scanner(System.in);
        System.out.print("Search : ");
        String search = sc.nextLine();
        for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
            sheet = wb.getSheetAt(sheetIndex);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                XSSFRow row = sheet.getRow(rowIndex);
                // System.out.println(row.getCell(0).getStringCellValue());
                if (row != null && row.getCell(0).getStringCellValue().equalsIgnoreCase(search)) {
                    /*Row getRow = sheet.getRow(row.getCell(1).getRowIndex());
                    Iterator<Row> rowIterator = sheet.iterator();
                     getRow = rowIterator.next();*/
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        System.out.print(cell.getStringCellValue() + "\t\t\t");
                    }
                    System.out.println("");

                }
            }
        }
    }
}
