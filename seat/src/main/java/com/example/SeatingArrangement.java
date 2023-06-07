package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

public class SeatingArrangement {
    public static void main(String[] args) {
        // Path to the Excel file
        String filePath = "miniproject/seat/src/main/java/com/example/SortedExcel.xlsx";

        // Read register numbers from the Excel sheet
        List<String> registerNumbers = readRegisterNumbersFromExcel(filePath);

        // Define the number of students in each branch
        int branch1Count = 18;
        int branch2Count = 12;

        // Generate seating arrangement
        List<List<String>> seatingArrangement = generateSeatingArrangement(registerNumbers, branch1Count, branch2Count);

        // Print the seating arrangement
        printSeatingArrangement(seatingArrangement);
    }

    private static List<String> readRegisterNumbersFromExcel(String filePath) {
        List<String> registerNumbers = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Assuming the register numbers are in the first column of the first sheet
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0); // Assuming register number is in the first column
                String registerNumber = cell.getStringCellValue();
                registerNumbers.add(registerNumber);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return registerNumbers;
    }

    private static List<List<String>> generateSeatingArrangement(List<String> registerNumbers, int branch1Count, int branch2Count) {
        // Shuffle the register numbers randomly
      //  Collections.shuffle(registerNumbers);

        // Divide the register numbers into two lists for each branch
        List<String> branch1Students = registerNumbers.subList(0, branch1Count);
        List<String> branch2Students = registerNumbers.subList(branch1Count, branch1Count + branch2Count);

        // Generate the seating arrangement
        List<List<String>> seatingArrangement = new ArrayList<>();
        int totalColumns = 5;

        // Determine the number of rows required for each branch
        int branch1Rows = (int) Math.ceil((double) branch1Count / totalColumns);
        int branch2Rows = (int) Math.ceil((double) branch2Count / totalColumns);

        // Assign seats for branch 1
        int branch1Index = 0;
        for (int row = 0; row < branch1Rows; row++) {
            List<String> seatingRow = new ArrayList<>();
            for (int col = 0; col < totalColumns; col++) {
                if (branch1Index < branch1Count) {
                    seatingRow.add(branch1Students.get(branch1Index));
                    branch1Index++;
                }
            }
            seatingArrangement.add(seatingRow);
        }

        // Assign seats for branch 2
        int branch2Index = 0;
        for (int row = 0; row < branch2Rows; row++) {
            List<String> seatingRow = new ArrayList<>();
            for (int col = 0; col < totalColumns; col++) {
                if (branch2Index < branch2Count) {
                    seatingRow.add(branch2Students.get(branch2Index));
                    branch2Index++;
                }
            }
            seatingArrangement.add(seatingRow);
        }

        return seatingArrangement;
    }

    private static void printSeatingArrangement(List<List<String>> seatingArrangement) {
        for (List<String> seatingRow : seatingArrangement) {
            for (String student : seatingRow) {
                System.out.print(student + "\t");
            }
            System.out.println();
        }
    }
}
