package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.*;

public class ExcelFileManipulation {

    public static void main(String[] args) {
        // Create a file chooser dialog to allow the user to select an Excel file
        File selectedFile = selectExcelFile();
        if (selectedFile == null) {
            return;
        }

        try (FileInputStream file = new FileInputStream(selectedFile);
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
            String filePath = selectedFile.getAbsolutePath();
            String extension = filePath.substring(filePath.lastIndexOf("."));
            String newFilePath = filePath.replace(extension, "_modified" + extension);
            try (FileOutputStream outputStream = new FileOutputStream(newFilePath)) {
                workbook.write(outputStream);
            }

            System.out.println("New file created at " + newFilePath);
        } catch (IOException e) {
            e.printStackTrace();
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
}
