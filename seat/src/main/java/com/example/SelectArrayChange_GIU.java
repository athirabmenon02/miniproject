package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class SelectArrayChange_GIU extends JFrame {
    private JButton selectFileButton;
    private JComboBox<String> sheetComboBox;
    private JButton displayButton;
    private String selectedFilePath;

    public SelectArrayChange_GIU() {
        super("Excel Reader");

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(400, 200);
        setLocationRelativeTo(null);

        selectFileButton = new JButton("Select File");
        selectFileButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int result = fileChooser.showOpenDialog(SelectArrayChange_GIU.this);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    selectedFilePath = selectedFile.getAbsolutePath();
                    populateSheetComboBox(selectedFilePath);
                }
            }
        });

        sheetComboBox = new JComboBox<>();
        displayButton = new JButton("Display");
        displayButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String selectedSheet = (String) sheetComboBox.getSelectedItem();
                if (selectedSheet != null) {
                    displaySelectedSheet(selectedSheet);
                }
            }
        });

        setLayout(new FlowLayout());
        add(selectFileButton);
        add(sheetComboBox);
        add(displayButton);
    }

    private void populateSheetComboBox(String filePath) {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            sheetComboBox.removeAllItems();
            for (int i = 0; i < numberOfSheets; i++) {
                sheetComboBox.addItem(workbook.getSheetName(i));
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void displaySelectedSheet(String sheetName) {
        if (selectedFilePath == null || selectedFilePath.isEmpty()) {
            System.out.println("No file selected.");
            return;
        }

        try (FileInputStream fis = new FileInputStream(selectedFilePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.out.println("Sheet not found: " + sheetName);
                return;
            }

            List<String> studentNames = extractStudentNames(sheet);
            System.out.println("Sheet " + sheetName + " student names: " + studentNames);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private List<String> extractStudentNames(Sheet sheet) {
        List<String> studentNames = new ArrayList<>();
        int rows = sheet.getPhysicalNumberOfRows();

        for (int i = 0; i < rows; i++) {
            Row row = sheet.getRow(i);
            Cell cell = row.getCell(0); // Assuming student names are in the first column

            if (cell != null && cell.getCellType() == CellType.STRING) {
                String studentName = cell.getStringCellValue();
                studentNames.add(studentName);
            }
        }

        return studentNames;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            SelectArrayChange_GIU gui = new SelectArrayChange_GIU();
            gui.setVisible(true);
        });
    }
}