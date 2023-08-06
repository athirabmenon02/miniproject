//integraeed gui with allocatrion.


package com.example.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
public class ExcelDataWriterWithGUI{
    static int Examsnum;
    private static Map<String, String> examSlots = new HashMap<>();
    private static List<String> selectedKeys = new ArrayList<>();

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }

    private static void createAndShowGUI() {
        // Create the main frame
        JFrame frame = new JFrame("Exam Selection");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 200);
        frame.setLayout(new BorderLayout());

        // Create the exam number label and text field
        JLabel examLabel = new JLabel("Number of Exams:");
        JTextField examTextField = new JTextField(10);

        // Create the checkboxes
        JCheckBox csCheckBox = new JCheckBox("CS");
        JCheckBox itCheckBox = new JCheckBox("IT");
        JCheckBox ecCheckBox = new JCheckBox("EC");
        JCheckBox aeiCheckBox = new JCheckBox("AEI");
        JCheckBox ereCheckBox = new JCheckBox("ERE");
        JCheckBox ceCheckBox = new JCheckBox("CE");

        // Create the slot selection boxes
        String[] slots = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J"};
        JComboBox<String> csSlotComboBox = new JComboBox<>(slots);
        JComboBox<String> itSlotComboBox = new JComboBox<>(slots);
        JComboBox<String> ecSlotComboBox = new JComboBox<>(slots);
        JComboBox<String> aeiSlotComboBox = new JComboBox<>(slots);
        JComboBox<String> ereSlotComboBox = new JComboBox<>(slots);
        JComboBox<String> ceSlotComboBox = new JComboBox<>(slots);
        
        // Create the submit button
        JButton submitButton = new JButton("Submit");
        submitButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int numExams = Integer.parseInt(examTextField.getText());

                selectedKeys.clear();
                if (csCheckBox.isSelected()) {
                    selectedKeys.add("CS-" + csSlotComboBox.getSelectedItem());
                }
                if (itCheckBox.isSelected()) {
                    selectedKeys.add("IT-" + itSlotComboBox.getSelectedItem());
                }
                if (ecCheckBox.isSelected()) {
                    selectedKeys.add("EC-" + ecSlotComboBox.getSelectedItem());
                }
                if (aeiCheckBox.isSelected()) {
                    selectedKeys.add("AEI-" + aeiSlotComboBox.getSelectedItem());
                }
                if (ereCheckBox.isSelected()) {
                    selectedKeys.add("ERE-" + ereSlotComboBox.getSelectedItem());
                }
                if (ceCheckBox.isSelected()) {
                    selectedKeys.add("CE-" + ceSlotComboBox.getSelectedItem());
                }

                if (selectedKeys.size() != numExams) {
                    JOptionPane.showMessageDialog(frame,
                            "Please select " + numExams + " exams.",
                            "Invalid Selection",
                            JOptionPane.ERROR_MESSAGE);
                    return;
                }
                Examsnum = numExams;
                frame.dispose();
                processExcelData();
            }
        });

        // Create the panel and add components
        JPanel panel = new JPanel();
        panel.setLayout(new GridBagLayout());
        GridBagConstraints constraints = new GridBagConstraints();
        constraints.fill = GridBagConstraints.HORIZONTAL;
        constraints.weightx = 1;
        constraints.weighty = 1;
        constraints.insets = new Insets(5, 5, 5, 5);

        constraints.gridx = 0;
        constraints.gridy = 0;
        panel.add(examLabel, constraints);

        constraints.gridx = 1;
        constraints.gridy = 0;
        panel.add(examTextField, constraints);

        constraints.gridx = 0;
        constraints.gridy = 1;
        panel.add(csCheckBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 1;
        panel.add(csSlotComboBox, constraints);

        constraints.gridx = 0;
        constraints.gridy = 2;
        panel.add(itCheckBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 2;
        panel.add(itSlotComboBox, constraints);

        constraints.gridx = 0;
        constraints.gridy = 3;
        panel.add(ecCheckBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 3;
        panel.add(ecSlotComboBox, constraints);

        constraints.gridx = 0;
        constraints.gridy = 4;
        panel.add(aeiCheckBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 4;
        panel.add(aeiSlotComboBox, constraints);

        constraints.gridx = 0;
        constraints.gridy = 5;
        panel.add(ereCheckBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 5;
        panel.add(ereSlotComboBox, constraints);

        constraints.gridx = 0;
        constraints.gridy = 6;
        panel.add(ceCheckBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 6;
        panel.add(ceSlotComboBox, constraints);

        constraints.gridx = 1;
        constraints.gridy = 7;
        panel.add(submitButton, constraints);

        // Add the panel to the frame
        frame.add(panel, BorderLayout.CENTER);

        // Display the frame
        frame.setVisible(true);
    }

    private static void processExcelData() {
        String sourceFilePath = "SortedExcel.xlsx";
        String destinationFilePath = "sample2.xlsx";

        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             Workbook sourceWorkbook = new XSSFWorkbook(fis);
             Workbook destinationWorkbook = new XSSFWorkbook(new FileInputStream("tholi2.xlsx"));
             FileOutputStream fos = new FileOutputStream(destinationFilePath)) {

            Map<String, List<String>> sheetColumnDataMap = new HashMap<>();

            for (int sheetIndex = 0; sheetIndex < sourceWorkbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sourceSheet = sourceWorkbook.getSheetAt(sheetIndex);
                String sheetName = sourceSheet.getSheetName();
                List<String> columnData = getColumnData(sourceSheet, 0);
                sheetColumnDataMap.put(sheetName, columnData);
            }

            int row = 6, count = 0, b1 = 0, b2 = 0, rowEnd = 11;
            int totalStudents = 0, k = 0, k2 =Examsnum-1;

            for (String key : selectedKeys) {
                List<String> columnData = sheetColumnDataMap.get(key);
                totalStudents += columnData.size();
            }

            int classrooms = (int) Math.ceil((double) totalStudents / 30.0);
            System.out.println(totalStudents);
            System.out.println(classrooms);

        for (int j = 0; j < classrooms; j++) {
                List<String> columnData;

                while (count < 18) {
                    System.out.println("b1");
                    
                    for (Map.Entry<String, List<String>> entry : sheetColumnDataMap.entrySet()) {
                        String sheetName = entry.getKey();
                        String desiredKey = selectedKeys.get(k);
                        columnData = sheetColumnDataMap.get(desiredKey);

                        Sheet destinationSheet = destinationWorkbook.getSheet("AAA");
                        if (destinationSheet == null) {
                            destinationSheet = destinationWorkbook.createSheet("AAA");
                        }

                        int rowIndex = row; // Starting row index in the destination sheet
                        int columnIndex = 1; // Starting column index in the destination sheet

                        for (int i = b1; i < columnData.size(); i++) {
                            String data = columnData.get(i);
                            Row destinationRow = destinationSheet.getRow(rowIndex);
                            if (destinationRow == null) {
                                destinationRow = destinationSheet.createRow(rowIndex);
                            }

                            Cell destinationCell = destinationRow.getCell(columnIndex);
                            if (destinationCell == null) {
                                destinationCell = destinationRow.createCell(columnIndex);
                            }
                            destinationCell.setCellValue(data);
                            count++;
                            b1++;
                            if (b1 == columnData.size()) {
                                if (k>=k2)
                                    break;
                                desiredKey = selectedKeys.get(++k);
                                columnData = sheetColumnDataMap.get(desiredKey);
                                b1 = i = 0;
                               // k++;
                            }

                            rowIndex++;
                            if (rowIndex > rowEnd) {
                                rowIndex = row; // Reset row index to 6 (B7) when reaching the end of the range
                                columnIndex += 3; // Move to the next column (E or G)
                            }
                            if (count >= 18)
                                break;
                        }
                        break;
                    }
                }
                count = 18;
                while (count < 30) {
                    System.out.println("b2");
                    for (Map.Entry<String, List<String>> entry : sheetColumnDataMap.entrySet()) {
                        String sheetName = entry.getKey();
                        String desiredKey = selectedKeys.get(k2);
                        columnData = sheetColumnDataMap.get(desiredKey);

                        Sheet destinationSheet = destinationWorkbook.getSheet("AAA");
                        if (destinationSheet == null) {
                            destinationSheet = destinationWorkbook.createSheet("AAA");
                        }

                        int rowIndex = row; // Starting row index in the destination sheet
                        int columnIndex = 2; // Starting column index in the destination sheet

                        for (int i = b2; i < columnData.size(); i++) {
                            String data = columnData.get(i);
                            Row destinationRow = destinationSheet.getRow(rowIndex);
                            if (destinationRow == null) {
                                destinationRow = destinationSheet.createRow(rowIndex);
                            }

                            Cell destinationCell = destinationRow.getCell(columnIndex);
                            if (destinationCell == null) {
                                destinationCell = destinationRow.createCell(columnIndex);
                            }

                            destinationCell.setCellValue(data);
                            count++;
                            b2++;
                            if (b2 == columnData.size()) {
                                if (k >= k2)
                                    break;
                                desiredKey = selectedKeys.get(--k2);
                                columnData = sheetColumnDataMap.get(desiredKey);
                                b2 = i = 0;
                              //  k2--;
                            }
                            rowIndex++;
                            if (rowIndex > rowEnd) {
                                rowIndex = row; // Reset row index to 6 (B7) when reaching the end of the range
                                columnIndex += 4; // Move to the next column (E or G)
                            }
                            if (count >= 30)
                                break;
                        }
                        break;
                    }
                }
                count = 0;
                row += 21;
                rowEnd += 21;
          
        }
            destinationWorkbook.write(fos);
            System.out.println("Data written to the destination file successfully.");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<String> getColumnData(Sheet sourceSheet, int columnIndex) {
        List<String> columnData = new ArrayList<>();

        for (int rowIndex = sourceSheet.getFirstRowNum(); rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
            Row row = sourceSheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    String data = cell.getStringCellValue();
                    columnData.add(data);
                }
            }
        }

        return columnData;
    }
}