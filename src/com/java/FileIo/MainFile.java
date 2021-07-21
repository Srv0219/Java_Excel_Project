package com.java.FileIo;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

public class MainFile {

    public static void main(String[] args) throws IOException {
        File excel = new File("D:\\demo\\students.xlsx");
        FileInputStream fis = new FileInputStream(excel);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);

        Scanner sc = new Scanner(System.in);

        DeleteFile delete = new DeleteFile();
        SearchFile search = new SearchFile();
        ReadFile read = new ReadFile();
        WriteFile write = new WriteFile();

        int number;
        do {
            System.out.println("1\t Read xlsx file");
            System.out.println("2\t Write xlsx file");
            System.out.println("3\t Search ");
            System.out.println("4\t Delete file");
            System.out.println("5\t Exit file");
            System.out.println("Please enter your choice:");
            number = sc.nextInt();
            switch (number) {

                case 1:
                    read.readData(sheet);
                    break;
                case 2:
                    write.writeData(sheet, excel, fis, wb);
                    break;
                case 3:
                    search.searchCells(sheet, wb);
                    break;
                case 4:
                    delete.deleteFile(sheet, wb);
                    break;
                case 5:
                    System.exit(0);
                    break;
                default:
                    read.readData(sheet);
            }
        } while (number != 3);
    }
}

 /* int row = sheet.getLastRowNum();
        int col =sheet.getRow(1).getLastCellNum();
        for (int r=0;r<=rows;r++){
            XSSFRow row = sheet.getRow(r);
            for (int c=0;c<=cols;c++){
                XSSFCell cell = row.getCell(c);
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        System.out.print(cell.getStringCellValue() + "\t\t\t");break;
                    case Cell.CELL_TYPE_NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t\t\t");break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t\t\t");break;
                    default:

                }
            }
        }*/
