package com.example;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelSorter {

    public static void main(String[] args) {
        String filename = "miniproject/seat/src/main/java/com/example/Excel.xlsx";
        try {
            sortExcelSheets(filename);
            System.out.println("Excel sheets sorted successfully.");
        } catch (IOException e) {
            System.out.println("Error occurred while sorting Excel sheets: " + e.getMessage());
        }
    }

    public static void sortExcelSheets(String filename) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(filename);
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        // Iterate over each sheet in the workbook
        for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            // Get the rows from the sheet
            Iterator<Row> rowIterator = sheet.iterator();
            List<Row> rows = new ArrayList<>();
            while (rowIterator.hasNext()) {
                rows.add(rowIterator.next());
            }

            // Sort the rows based on the first column
            Collections.sort(rows, new Comparator<Row>() {
                @Override
                public int compare(Row row1, Row row2) {
                    Cell cell1 = row1.getCell(0);
                    Cell cell2 = row2.getCell(0);
                    if (cell1 == null && cell2 == null) {
                        return 0;
                    } else if (cell1 == null) {
                        return 1;
                    } else if (cell2 == null) {
                        return -1;
                    } else {
                        String value1 = cell1.getStringCellValue();
                        String value2 = cell2.getStringCellValue();
                        return value1.compareTo(value2);
                    }
                }
            });

            // Clear the existing rows in the sheet
            while (sheet.getLastRowNum() >= 0) {
                sheet.removeRow(sheet.getRow(0));
            }

            // Add the sorted rows back to the sheet
            for (Row row : rows) {
                sheet.createRow(sheet.getLastRowNum() + 1);
                for (int i = 0; i < row.getLastCellNum(); i++) {
                    Cell oldCell = row.getCell(i);
                    if (oldCell != null) {
                        Cell newCell = sheet.getRow(sheet.getLastRowNum()).createCell(i);
                        newCell.setCellStyle(oldCell.getCellStyle());
                      //  newCell.setCellType(oldCell.getCellType());
                        switch (oldCell.getCellType()) {
                            case STRING:
                                newCell.setCellValue(oldCell.getStringCellValue());
                                break;
                            case NUMERIC:
                                newCell.setCellValue(oldCell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                newCell.setCellValue(oldCell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                newCell.setCellFormula(oldCell.getCellFormula());
                                break;
                            default:
                                break;
                        }
                    }
                }
            }
        }

        // Write the sorted workbook back to the file
        FileOutputStream fileOutputStream = new FileOutputStream(filename);
        workbook.write(fileOutputStream);
        workbook.close();
        fileOutputStream.close();
    }
}
