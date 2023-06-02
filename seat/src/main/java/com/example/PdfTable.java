//each sheet is created for each page-formatting better


package com.example;

import technology.tabula.ObjectExtractor;
import technology.tabula.Page;
import technology.tabula.Table;
import technology.tabula.extractors.BasicExtractionAlgorithm;
import technology.tabula.extractors.ExtractionAlgorithm;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class PdfTable {

    public static void main(String[] args) {
        try {
            // Load the PDF file
            PDDocument pdfFile = PDDocument.load(new File("src/main/java/com/example/timetable.pdf"));
            ObjectExtractor extractor = new ObjectExtractor(pdfFile);

            // Extract tables from each page
            ExtractionAlgorithm algorithm = new BasicExtractionAlgorithm();
            try (Workbook workbook = new XSSFWorkbook()) {
                int sheetIndex = 0;

                for (int pageIndex = 2; pageIndex <=9; pageIndex++) {
                    Page page = extractor.extract(pageIndex);
                    List<? extends Table> tables = algorithm.extract(page);

                    for (Table table : tables) {
                        Sheet sheet = workbook.createSheet("Table " + (sheetIndex + 1));
                        sheetIndex++;

                        for (int row = 0; row < table.getRowCount(); row++) {
                            Row excelRow = sheet.createRow(row);

                            for (int col = 0; col < table.getColCount(); col++) {
                                Cell excelCell = excelRow.createCell(col);
                                excelCell.setCellValue(table.getCell(row, col).getText());
                            }
                        }
                    }
                }

                // Write the Excel file
                FileOutputStream out = new FileOutputStream("output.xlsx");
                workbook.write(out);
                out.close();
            }
            // Close the extractor
            extractor.close();

            System.out.println("PDF converted to Excel successfully!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
