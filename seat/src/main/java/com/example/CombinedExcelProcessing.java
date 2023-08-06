
//final part1
package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CombinedExcelProcessing {

    public static void main(String[] args) throws IOException {
        // Step 1: Select the input Excel file and perform modifications
        File inputFile = selectExcelFile();
        if (inputFile == null) {
            System.out.println("No input file selected. Exiting...");
            return;
        }

        File modifiedFile = modifyExcelFile(inputFile);
        if (modifiedFile == null) {
            System.out.println("Error occurred while modifying the Excel file. Exiting...");
            return;
        }

        // Step 2: Sort the modified file into different sheets in a new Excel file
        sortModifiedFile(modifiedFile);
         String filePath = "SortedExcel.xlsx";
         int rowToDelete = 0; 
         try (FileInputStream file = new FileInputStream(new File(filePath));
             Workbook workbook = WorkbookFactory.create(file)) {

            // Iterate over all sheets in the workbook
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);

                if (rowToDelete <= sheet.getLastRowNum()) {
                    sheet.shiftRows(rowToDelete + 1, sheet.getLastRowNum(), -1);
       
                }
            }

            // Write the modified workbook back to the file
            try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
                workbook.write(outputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
         try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (Sheet sheet : workbook) {
                // Get the register numbers from the first column
                ColumnData columnData = getColumnData(sheet, 0);

                // Sort the register numbers in ascending order
                columnData.sort();

                // Write the sorted register numbers to the same sheet
                writeColumnData(sheet, columnData);
            }

            // Save the updated workbook back to the original file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
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
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            Cell cell = row.createCell(0);
            cell.setCellValue(registerNumber);
            rowIndex++;
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
    
    
    private static File selectExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.home")));
        fileChooser.setDialogTitle("Select an Excel file");
        fileChooser.setFileFilter(new FileNameExtensionFilter("Excel files", "xlsx", "xls"));
        int result = fileChooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            return fileChooser.getSelectedFile();
        } else {
            return null;
        }
    }

    private static File modifyExcelFile(File inputFile) {
        try (FileInputStream file = new FileInputStream(inputFile);
             Workbook workbook = WorkbookFactory.create(file)) {

            Sheet sheet = workbook.getSheetAt(0);

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);

            // Create a new column called "Register No"
            Row headerRow = sheet.getRow(1);
            int registerNoColumnIndex = headerRow.getLastCellNum();
            Cell registerNoHeaderCell = headerRow.createCell(registerNoColumnIndex);
            registerNoHeaderCell.setCellValue("Register No");
            registerNoHeaderCell.setCellStyle(style);
            sheet.autoSizeColumn(4);

            // Copy data from the original sheet to the new sheet
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell firstCell = row.getCell(0);

                // Extract text within parentheses in the first column
                String firstCellValue = firstCell.getStringCellValue();
                int startIndex = firstCellValue.indexOf("(");
                int endIndex = firstCellValue.indexOf(")");
                if (startIndex != -1 && endIndex != -1) {
                    String textWithinParentheses = firstCellValue.substring(startIndex + 1, endIndex);
                    Cell registerNoCell = row.createCell(registerNoColumnIndex);
                    registerNoCell.setCellValue(textWithinParentheses);
                }
            }

            // Write the modified Excel file to a new file
            String filePath = inputFile.getAbsolutePath();
            String extension = filePath.substring(filePath.lastIndexOf("."));
            String newFilePath = filePath.replace(extension, "_modified" + extension);
            try (FileOutputStream outputStream = new FileOutputStream(newFilePath)) {
                workbook.write(outputStream);
            }

            System.out.println("Modified file created at " + newFilePath);
            return new File(newFilePath);
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private static void sortModifiedFile(File modifiedFile) throws IOException {
        String excelFilePath = modifiedFile.getAbsolutePath();
        String sheetName = "AppearingStudentEligibilityRepo";
        String branchColumnName = "Branch Name";
        String slotColumnName = "Slot";
        FileInputStream inputStream = new FileInputStream(modifiedFile);
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
                Cell studentCell = row.getCell(4);
                String studentValue = studentCell.getStringCellValue();
                Cell newCell = newRow.createCell(0);
                newCell.setCellValue(studentValue);
            }

            // Adjust column widths based on content
            newSheet.autoSizeColumn(0);
        }

        inputStream.close();
        FileOutputStream outputStream = new FileOutputStream("SortedExcel.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();

        System.out.println("Sorting completed. New Excel file created: Excel.xlsx");
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