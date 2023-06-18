package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelDataWriter {
    public static void main(String[] args) {
        String sourceFilePath = "miniproject/seat/src/main/java/com/example/SortedExcel.xlsx";
        String destinationFilePath = "sample2.xlsx";

        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             Workbook sourceWorkbook = new XSSFWorkbook(fis);
             Workbook destinationWorkbook = new XSSFWorkbook(new FileInputStream("miniproject/seat/src/main/java/com/example/sample.xlsx"));
             FileOutputStream fos = new FileOutputStream(destinationFilePath)) {

            Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
            Sheet destinationSheet = destinationWorkbook.getSheet("AAA"); // Modify the sheet name as per your template

            int rowIndex = 6; // Starting row index in the destination sheet
            int columnIndex = 1; // Starting column index in the destination sheet (B column)

            for (int i = 0; i < 18; i++) {
                Row sourceRow = sourceSheet.getRow(i);
                Cell sourceCell = sourceRow.getCell(0); // Assuming data is in the first column of the source sheet

                Row destinationRow = destinationSheet.getRow(rowIndex);
                if (destinationRow == null) {
                    destinationRow = destinationSheet.createRow(rowIndex);
                }

                Cell destinationCell = destinationRow.getCell(columnIndex);
                if (destinationCell == null) {
                    destinationCell = destinationRow.createCell(columnIndex);
                }

                destinationCell.setCellValue(getCellValueAsString(sourceCell));

                rowIndex++;
                if (rowIndex > 11) {
                    rowIndex = 6; // Reset row index to 6 (B7) when reaching the end of the range
                    columnIndex += 3; // Move to the next column (E or G)
                }
            }
            for (int i = 18; i < 38; i++) {
                 columnIndex = 0;
                Row sourceRow = sourceSheet.getRow(i);
                Cell sourceCell = sourceRow.getCell(0); // Assuming data is in the first column of the source sheet
                rowIndex = 27;
                Row destinationRow = destinationSheet.getRow(rowIndex);
                if (destinationRow == null) {
                    destinationRow = destinationSheet.createRow(rowIndex);
                }

                Cell destinationCell = destinationRow.getCell(columnIndex);
                if (destinationCell == null) {
                    destinationCell = destinationRow.createCell(columnIndex);
                }

                destinationCell.setCellValue(getCellValueAsString(sourceCell));

                rowIndex++;
                if (rowIndex > 33) {
                    rowIndex = 27; // Reset row index to 27 (B7) when reaching the end of the range
                    columnIndex += 3; // Move to the next column (E or G)
                }
            }

            destinationWorkbook.write(fos);
            System.out.println("Data written to the destination file successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}