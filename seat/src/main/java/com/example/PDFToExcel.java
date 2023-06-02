package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PDFToExcel {
    public static void main(String[] args) {
        File pdfFile = new File("src/main/java/com/example/Adithya.pdf");
        File outputFile = new File("javatpoint.xlsx");
        try (InputStream inputStream = new FileInputStream(pdfFile);
             OutputStream outputStream = new FileOutputStream(outputFile)) {
            PDDocument document = PDDocument.load(inputStream);
            XSSFWorkbook workbook = new XSSFWorkbook();
            // Convert each page of the PDF to a new sheet in the Excel workbook
            for (int pageIndex = 0; pageIndex < document.getNumberOfPages(); pageIndex++) {
                String sheetName = "Page " + (pageIndex + 1);
                workbook.createSheet(sheetName);
                PDFToExcel.convertPageToSheet(document, pageIndex, workbook.getSheet(sheetName));
            }
            workbook.write(outputStream);
            workbook.close();
            document.close();
            System.out.println("PDF to Excel conversion completed successfully!");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void convertPageToSheet(PDDocument document, int pageIndex, org.apache.poi.ss.usermodel.Sheet sheet) {
        try {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setStartPage(pageIndex + 1);
            stripper.setEndPage(pageIndex + 1);
            String text = stripper.getText(document);
            String[] lines = text.split("\\r?\\n");
            for (int i = 0; i < lines.length; i++) {
                String[] columns = lines[i].split("\\t");
                Row row = sheet.createRow(i);
                for (int j = 0; j < columns.length; j++) {
                    Cell cell = row.createCell(j);
                    cell.setCellValue(columns[j]);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
