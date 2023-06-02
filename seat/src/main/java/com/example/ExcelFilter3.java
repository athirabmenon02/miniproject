package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFilter3{

    public static void main(String[] args) throws IOException {
        String excelFilePath = "src/main/java/com/example/appearance.xlsx";
        String sheetName = "AppearingStudentEligibilityRepo";
        String branchColumnName = "Branch Name";
        String slotColumnName = "Slot";

        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheet(sheetName);

        Map<String, List<Row>> branchSlotMap = new HashMap<>();
        for (int i = 2; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            Cell branchCell = row.getCell(getColumnIndex(sheet, branchColumnName));
            Cell slotCell = row.getCell(getColumnIndex(sheet, slotColumnName));
            String branch = branchCell.getStringCellValue().trim();
            String slot = slotCell.getStringCellValue().trim();
            String key = getBranchAbbreviation(branch) + "-" + slot;
            List<Row> rows = branchSlotMap.getOrDefault(key, new ArrayList<>());
            rows.add(row);
            branchSlotMap.put(key, rows);
        }

        // Remove the original sheet
        workbook.removeSheetAt(workbook.getSheetIndex(sheet));

        for (String key : branchSlotMap.keySet()) {
            String[] parts = key.split("-");
            String branch = parts[0];
            String slot = parts[1];
            String newSheetName = generateUniqueSheetName(workbook, branch, slot);
            Sheet newSheet = workbook.createSheet(newSheetName);
            Row headerRow = newSheet.createRow(0);
            int columnIndex = 0;
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);
            for (Cell cell : sheet.getRow(1)) {
                String headerValue = cell.getStringCellValue();
                if (columnIndex == 0) {
                    Cell headerCell = headerRow.createCell(columnIndex);
                    headerCell.setCellValue(headerValue);
                    headerCell.setCellStyle(style);
                }
                columnIndex++;
            }
            for (Row row : branchSlotMap.get(key)) {
                int newRowNum = newSheet.getLastRowNum() + 1;
                Row newRow = newSheet.createRow(newRowNum);
                Cell studentCell = row.getCell(0);
                String studentValue = studentCell.getStringCellValue();
                Cell newCell = newRow.createCell(0);
                newCell.setCellValue(studentValue);
            }

            // Adjust column widths based on content
            newSheet.autoSizeColumn(0);
        }

        inputStream.close();
        FileOutputStream outputStream = new FileOutputStream("Excelfilter3.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(1); // Get the header row (first row)
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
