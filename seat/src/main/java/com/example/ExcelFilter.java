package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFilter {

    public static void main(String[] args) throws IOException {
        String excelFilePath = "src/main/java/com/example/appearence.xlsx";
        String sheetName = "AppearingStudentEligibilityRepo";
        String branchColumnName = "Branch Name";
        String slotColumnName = "Slot";

        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);

        Map<String, List<Row>> branchSlotMap = new HashMap<>();
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell branchCell = row.getCell(getColumnIndex(sheet, branchColumnName));
            Cell slotCell = row.getCell(getColumnIndex(sheet, slotColumnName));
            String branch = branchCell.getStringCellValue().trim();
            String slot = slotCell.getStringCellValue().trim();
            String key = getBranchAbbreviation(branch) + "-" + slot;
            List<Row> rows = branchSlotMap.getOrDefault(key, new ArrayList<Row>());
            rows.add(row);
            branchSlotMap.put(key, rows);
        }
        
        for (String key : branchSlotMap.keySet()) {
            String[] parts = key.split("-");
            String branch = parts[0];
            String slot = parts[1];
            String newSheetName = generateUniqueSheetName(workbook, branch, slot);
            Sheet newSheet = workbook.createSheet(newSheetName);
            Row headerRow = newSheet.createRow(0);
            Row newRow = newSheet.createRow(1);
            int columnIndex = 0;
            for (Cell cell : sheet.getRow(1)) {
                headerRow.createCell(columnIndex).setCellValue(cell.getStringCellValue());
                columnIndex++;
            }
            for (Row row : branchSlotMap.get(key)) {
                int newRowNum = newRow.getRowNum();
                newRow = newSheet.createRow(newRowNum + 1);
                columnIndex = 0;
                for (Cell cell : row) {
                    newRow.createCell(columnIndex).setCellValue(cell.getStringCellValue());
                    columnIndex++;
                }
            }
        }

        inputStream.close();
        FileOutputStream outputStream = new FileOutputStream("out2.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(1); // Get the header row (second row)
        for (int i = 0; i < headerRow.getLastCellNum(); i++) {
            Cell cell = headerRow.getCell(i);
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return i;
            }
        }
        throw new IllegalArgumentException("Column " + columnName + " not found");
    }

    private static String generateUniqueSheetName(Workbook workbook, String branch, String slot) {
        String cleanedBranch = getBranchAbbreviation(branch);
        String sheetName = cleanedBranch + "-" + slot;
        String uniqueSheetName = sheetName;
        int suffix = 1;
        while (workbook.getSheet(uniqueSheetName) != null) {
            uniqueSheetName = sheetName + "-" + suffix;
            suffix++;
        }
        return uniqueSheetName;
    }

    private static String getBranchAbbreviation(String branch) {
        switch (branch) {
            case "COMPUTER SCIENCE & ENGINEERING":
                return "CS";
            case "INFORMATION TECHNOLOGY":
                return "IT";
            case "ELECTRONICS & COMMUNICATION ENGG":
                return "EC";
            case "APPLIED ELECTRONICS & INSTRUMENTATION ENGINEERING":
                return "AEI";
            case "CIVIL ENGINEERING":
                return "CE";
            default:
                return branch;
        }
    }
}
