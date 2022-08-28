package com.uc.Excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

public class MakeDuplicateTable {

    public static void main(String[] args) {
        String path = "H://upvc.xlsx";
        int numberOfCopies=76;
        //Selection
        int top=0;
        int bottom=14;
        //Open excel file
        FileInputStream fileInputStream = readExcel(path);
        //Create duplicate tables
        //performCopy(fileInputStream,numberOfCopies, top,bottom);
        removeTables(fileInputStream,numberOfCopies,top,bottom);

    }

    private static FileInputStream readExcel(String path) {
        FileInputStream fileInputStream = null;
        try {
            fileInputStream = new FileInputStream(path);


        } catch (FileNotFoundException e) {
            System.out.println("file not found \n" + e);
        }
        return fileInputStream;
    }

    private static void performCopy(FileInputStream fileInputStream, int numberOfCopies, int top, int bottom) {

        XSSFWorkbook workbook;
        int startPoint=bottom+2;
        int roomNumber=901;
        Row row;

        try {
            workbook = new XSSFWorkbook(fileInputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
//Get the Desired sheet
        XSSFSheet sheet = workbook.getSheetAt(1);
        CellStyle style = workbook.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);

        for(int i=1;i<=numberOfCopies;i++){
//            sheet.removeMergedRegion(1);
            sheet.copyRows(top,bottom,startPoint, new CellCopyPolicy());
            row=sheet.getRow(startPoint);
            Cell cell=row.createCell(1);
            roomNumber++;
            cell.setCellValue(roomNumber);
            cell.setCellStyle(style);
            sheet.removeMergedRegion(1);
            startPoint+=bottom+2;
        }

        try {
            FileOutputStream out = new FileOutputStream(
                    "H://upvc.xlsx");
            workbook.write(out);
            out.close();
            fileInputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        //Increment over rows
//        for (Row row : sheet) {
//            //Iterate and get the cells from the row
//            Iterator cellIterator = row.cellIterator();
//        }
    }

    private static void removeTables(FileInputStream fileInputStream, int numberOfCopies, int top, int bottom) {

        XSSFWorkbook workbook;
        int startPoint=bottom+2;
        int roomNumber=901;
        int total=startPoint+13;
        Row row;

        try {
            workbook = new XSSFWorkbook(fileInputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        //Get the Desired sheet
        XSSFSheet sheet = workbook.getSheetAt(1);
//        CellStyle style = workbook.createCellStyle();
//        style.setBorderTop(BorderStyle.THIN);
//        style.setAlignment(HorizontalAlignment.CENTER);

        List<XSSFTable> list=sheet.getTables();
        sheet.removeTable(new XSSFTable(){});
        for(int i=1;i<=numberOfCopies;i++){
//            sheet.removeMergedRegion(1);
           // sheet.removeTable();//copyRows(top,bottom,startPoint, new CellCopyPolicy());
            row=sheet.getRow(total);
            Double value=row.getCell(4).getNumericCellValue();
            if(value<1){
                System.out.println("Matching found "+i);
            }
//            Cell cell=row.createCell(1);
            roomNumber++;
//            cell.setCellValue(roomNumber);
//            cell.setCellStyle(style);
            sheet.removeMergedRegion(1);
            startPoint+=bottom+2;
        }

        try {
            FileOutputStream out = new FileOutputStream(
                    "H://upvc.xlsx");
           // workbook.write(out);
            out.close();
            fileInputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        //Increment over rows
//        for (Row row : sheet) {
//            //Iterate and get the cells from the row
//            Iterator cellIterator = row.cellIterator();
//        }
    }
}
