package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;
import java.awt.Color;

public class SeatingAllocation {
    private static final String SEATING_PLAN_FILE_PATH = "miniproject/seat/src/main/java/com/example/Excelfilter3.xlsx";
    private static final String[] SEATING_PLAN_SHEET_NAMES = {"IT-A", "IT-D"};
    private static final int MAX_ROWS = 3;
    private static final int MAX_COLUMNS = 7;

    private Map<String, List<String>> studentMap;
    private Map<String, String> seatAllocationMap;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            SeatingAllocation seatingAllocation = new SeatingAllocation();
            seatingAllocation.loadStudentData();
            seatingAllocation.allocateSeats();
            seatingAllocation.showSeatingPlan();
        });
    }

    private void loadStudentData() {
        studentMap = new HashMap<>();

        try (FileInputStream file = new FileInputStream(new File(SEATING_PLAN_FILE_PATH));
             Workbook workbook = new XSSFWorkbook(file)) {
            for (String sheetName : SEATING_PLAN_SHEET_NAMES) {
                Sheet sheet = workbook.getSheet(sheetName);
                List<String> students = new ArrayList<>();

                for (Row row : sheet) {
                    Cell cell = row.getCell(0); // Assuming student names are in the first column
                    if (cell != null && cell.getCellType() == CellType.STRING) {
                        students.add(cell.getStringCellValue());
                    }
                }

                studentMap.put(sheetName, students);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void allocateSeats() {
        seatAllocationMap = new HashMap<>();

        int seatIndex = 0;
        int middleRowSeatIndex = 2; // Index of the seat in the middle row

        for (String sheetName : SEATING_PLAN_SHEET_NAMES) {
            List<String> students = studentMap.get(sheetName);
            if (students == null) {
                continue;
            }

            for (String student : students) {
                String seat = getSeatNumber(seatIndex, middleRowSeatIndex);
                seatAllocationMap.put(student, seat);
                seatIndex += 2;
            }

            seatIndex = 1;
        }
    }

    private String getSeatNumber(int seatIndex, int middleRowSeatIndex) {
        int rowNumber = 1 + (seatIndex / MAX_COLUMNS); // 1 represents the first row, adjust as per your sheet
        int seatInRow = seatIndex % MAX_COLUMNS;
    
        if (rowNumber == 2) {
            if (seatInRow >= middleRowSeatIndex) {
                seatInRow++;
            }
        } else if (rowNumber == 3) {
            seatInRow += 2;
        }
    
        return rowNumber + " " + seatInRow;
    }
    
    private void showSeatingPlan() {
        JFrame frame = new JFrame("Seating Plan");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    
        JPanel panel = new JPanel();
        panel.setLayout(new GridLayout(MAX_ROWS, MAX_COLUMNS, 10, 10));
    
        JLabel[][] seatLabels = new JLabel[MAX_ROWS][MAX_COLUMNS];
    
        for (int row = 0; row < MAX_ROWS; row++) {
            for (int col = 0; col < MAX_COLUMNS; col++) {
                JLabel seatLabel = new JLabel();
                seatLabel.setOpaque(true);
                seatLabel.setBackground(Color.WHITE);
                seatLabel.setBorder(BorderFactory.createLineBorder(Color.BLACK));
                panel.add(seatLabel);
                seatLabels[row][col] = seatLabel;
            }
        }
    
        for (Map.Entry<String, String> entry : seatAllocationMap.entrySet()) {
            String student = entry.getKey();
            String seat = entry.getValue();
    
            String[] seatParts = seat.split(" ");
            int row = Integer.parseInt(seatParts[0]) - 1;
            int col = Integer.parseInt(seatParts[1]) - 1;
    
            seatLabels[row][col].setText(student);
        }
    
        frame.getContentPane().add(panel);
        frame.pack();
        frame.setVisible(true);
    }
}
