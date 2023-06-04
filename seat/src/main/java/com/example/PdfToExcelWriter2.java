package com.example;
import org.apache.pdfbox.pdmodel.*;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PdfToExcelWriter2 {

    public static void main(String[] args) {
        try (PDDocument document = PDDocument.load(new FileInputStream("miniproject/seat/src/main/java/com/example/timetable.pdf"))) {
            PDFTextStripper stripper = new PDFTextStripper();
            stripper.setStartPage(2);
            stripper.setEndPage(9);
            String tableData = stripper.getText(document);
            String[] rows = tableData.split("\n");
            int rowCount = rows.length;
            int colCount = rows[0].split("\t").length;
            try (Workbook workbook = new XSSFWorkbook()) {
              Sheet sheet = workbook.createSheet("Table Data");
              for (int i = 0; i < rowCount; i++) {
                  Row row = sheet.createRow(i);
                  String[] cells = rows[i].split("\t");
                  for (int j = 0; j < colCount; j++) {
                      Cell cell = row.createCell(j);
                      cell.setCellValue(cells[j]);
                  }
               

              }
              FileOutputStream out = new FileOutputStream("output.xlsx");
              workbook.write(out);
              out.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}