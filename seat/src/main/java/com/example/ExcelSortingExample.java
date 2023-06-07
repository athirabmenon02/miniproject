//register number sort
package com.example;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSortingExample {
    public static void main(String[] args) {
        String inputFilePath = "miniproject/seat/src/main/java/com/example/Excel.xlsx";
        String outputFilePath = "miniproject/seat/src/main/java/com/example/SortedExcel.xlsx";

        try (FileInputStream fis = new FileInputStream(inputFilePath);
             Workbook inputWorkbook = new XSSFWorkbook(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            for (Sheet inputSheet : inputWorkbook) {
                Sheet outputSheet = outputWorkbook.createSheet(inputSheet.getSheetName());

                // Get the register numbers from the first column
                ColumnData columnData = getColumnData(inputSheet, 0);

                // Sort the register numbers in ascending order
                columnData.sort();

                // Write the sorted register numbers to the output sheet
                writeColumnData(outputSheet, columnData);
            }

            // Save the output workbook to a new Excel file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                outputWorkbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static ColumnData getColumnData(Sheet sheet, int columnIndex) {
        ColumnData columnData = new ColumnData();

        for (Row row : sheet) {
            Cell cell = row.getCell(columnIndex);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                columnData.add(cell.getStringCellValue());
            }
        }

        return columnData;
    }

    private static void writeColumnData(Sheet sheet, ColumnData columnData) {
        int rowIndex = 0;
        for (String registerNumber : columnData.getSortedValues()) {
            Row row = sheet.createRow(rowIndex++);
            Cell cell = row.createCell(0);
            cell.setCellValue(registerNumber);
        }
    }

    private static class ColumnData {
        private List<String> registerNumbers = new ArrayList<>();

        public void add(String registerNumber) {
            registerNumbers.add(registerNumber);
        }

        public void sort() {
            Collections.sort(registerNumbers);
        }

        public List<String> getSortedValues() {
            return registerNumbers;
        }
    }
}

