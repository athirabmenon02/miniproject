package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelDataWriter2 {
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

            for (int i = 0; i < 28; i++) { // Loop through rows 1 to 28 in the source sheet
                Row sourceRow = sourceSheet.getRow(i);

                // Copy to column B (index 1)
                Cell sourceCell = sourceRow.getCell(0); // Assuming data is in the first column of the source sheet
                Row destinationRow = destinationSheet.getRow(rowIndex);
                if (destinationRow == null) {
                    destinationRow = destinationSheet.createRow(rowIndex);
                }
                Cell destinationCell = destinationRow.createCell(1);
                destinationCell.setCellValue(getCellValueAsString(sourceCell));

                // Copy to column E (index 4)
                sourceCell = sourceRow.getCell(1); // Assuming data is in the second column of the source sheet
                destinationCell = destinationRow.createCell(4);
                destinationCell.setCellValue(getCellValueAsString(sourceCell));

                // Copy to column H (index 7)
                sourceCell = sourceRow.getCell(2); // Assuming data is in the third column of the source sheet
                destinationCell = destinationRow.createCell(7);
                destinationCell.setCellValue(getCellValueAsString(sourceCell));

                rowIndex++;
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
