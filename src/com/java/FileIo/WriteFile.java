package com.java.FileIo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Date;
import java.util.HashMap;
import java.util.Scanner;
import java.util.Set;

public class WriteFile {
    public void writeData(XSSFSheet sheet, File excel, FileInputStream fis, XSSFWorkbook wb) {

        int rownum1 = sheet.getLastRowNum() + 1;
        Scanner sc = new Scanner(System.in);
        System.out.println("Id: ");
        String id = Integer.toString(rownum1);
        System.out.println(id);
        System.out.println("firstName: ");
        String firstName = sc.nextLine();
        System.out.println("lastName: ");
        String lastName = sc.nextLine();
        System.out.println("city: ");
        String city = sc.nextLine();
        System.out.println("state: ");
        String state = sc.nextLine();
        System.out.println("phoneNo: ");
        String phoneNo = sc.nextLine();
        System.out.println("email: ");
        String email = sc.nextLine();
        HashMap<String, Object[]> newData = new HashMap<String, Object[]>();
        newData.put(id, new Object[]{firstName, lastName, city, state, phoneNo, email});

        Set<String> keyset = newData.keySet();
        int rownum = sheet.getLastRowNum() + 1;
        for (String key : keyset) {
            Row row = sheet.createRow(rownum++);
            Object[] objArr = newData.get(key);
            int cellnum = 0;
            for (Object obj : objArr) {
                Cell cell = row.createCell(cellnum++);
                if (obj instanceof String) {
                    cell.setCellValue((String) obj);
                } else if (obj instanceof Boolean) {
                    cell.setCellValue((Boolean) obj);
                } else if (obj instanceof Date) {
                    cell.setCellValue((Date) obj);
                } else if (obj instanceof Double) {
                    cell.setCellValue((Double) obj);
                }
            }
        }
        try {
            FileOutputStream os = new FileOutputStream(excel);
            wb.write(os);
            os.close();
            wb.close();
            fis.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            System.out.println("Data Added successfull in xlsx file");
        }
    }
}
