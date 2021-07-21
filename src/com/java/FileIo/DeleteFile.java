package com.java.FileIo;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Scanner;

public class DeleteFile {
    public void deleteFile(XSSFSheet sheet, XSSFWorkbook wb) {
        Scanner sc = new Scanner(System.in);
        System.out.print("Delete : ");
        String search = sc.nextLine();
        for (int sheetIndex = 0; sheetIndex < wb.getNumberOfSheets(); sheetIndex++) {
            sheet = wb.getSheetAt(sheetIndex);
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                // System.out.println(row.getCell(0).getStringCellValue());
                if (row != null && row.getCell(0).getStringCellValue().equalsIgnoreCase(search)) {
                    sheet.removeRow(row);

                }
            }
        }

    }
}
