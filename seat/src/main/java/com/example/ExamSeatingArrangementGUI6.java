package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.HashMap;

public class ExamSeatingArrangementGUI6 extends JFrame implements ActionListener {
    private JTextField examsField;
    private Map<JCheckBox, JComboBox<String>> branchSlotMap;
    private String filePath;

    public ExamSeatingArrangementGUI6() {
        setTitle("Exam Seating Arrangement");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new GridLayout(0, 2));

        JLabel examsLabel = new JLabel("Number of Exams: ");
        examsField = new JTextField(10);

        JLabel branchLabel = new JLabel("Branch: ");
        JPanel branchPanel = new JPanel(new GridLayout(0, 1));

        branchSlotMap = new HashMap<>();
        addBranchSlotSelection("CS", branchPanel);
        addBranchSlotSelection("IT", branchPanel);
        addBranchSlotSelection("AEI", branchPanel);
        addBranchSlotSelection("CE", branchPanel);
        addBranchSlotSelection("EC", branchPanel);

        JButton submitButton = new JButton("Submit");
        submitButton.addActionListener(this);

        add(examsLabel);
        add(examsField);
        add(branchLabel);
        add(branchPanel);
        add(submitButton);

        pack();
        setLocationRelativeTo(null);
        setVisible(true);
    }

    private void addBranchSlotSelection(String branch, JPanel branchPanel) {
        JCheckBox branchCheckbox = new JCheckBox(branch);
        JLabel branchLabel = new JLabel(branch + " Slot: ");
        JComboBox<String> slotDropdown = new JComboBox<>(generateSlotList());
        slotDropdown.setEnabled(false); // Initially disabled
        branchCheckbox.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                slotDropdown.setEnabled(branchCheckbox.isSelected());
            }
        });
        branchSlotMap.put(branchCheckbox, slotDropdown);
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        panel.add(branchCheckbox);
        panel.add(branchLabel);
        panel.add(slotDropdown);
        branchPanel.add(panel);
    }

    private String[] generateSlotList() {
        int numSlots = 26; // A to Z
        String[] slots = new String[numSlots];
        for (int i = 0; i < numSlots; i++) {
            slots[i] = String.valueOf((char) ('A' + i));
        }
        return slots;
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getActionCommand().equals("Submit")) {
            int numberOfExams = Integer.parseInt(examsField.getText());
            List<String> selectedBranches = new ArrayList<>();
            List<String> selectedSlots = new ArrayList<>();

            for (JCheckBox branchCheckbox : branchSlotMap.keySet()) {
                JComboBox<String> slotDropdown = branchSlotMap.get(branchCheckbox);
                if (branchCheckbox.isSelected()) {
                    String selectedSlot = slotDropdown.getSelectedItem().toString();
                    if (!selectedSlot.isEmpty()) {
                        selectedBranches.add(branchCheckbox.getText());
                        selectedSlots.add(selectedSlot);
                    }
                }
            }

            // Generate seating arrangement based on selected branches and slots
            for (int i = 0; i < selectedBranches.size(); i++) {
                String branch = selectedBranches.get(i);
                String slot = selectedSlots.get(i);

                // Assuming the Excel sheet contains a sheet named "branch-slot"
                String sheetName = branch + "-" + slot;

                // Read register numbers from the Excel sheet
                List<String> registerNumbers = readRegisterNumbersFromExcel(filePath, sheetName);

                // Define the number of students in each branch
                int branch1Count = 18;
                int branch2Count = 12;

                // Generate seating arrangement
                List<List<String>> seatingArrangement = generateSeatingArrangement(registerNumbers, branch1Count, branch2Count);

                // Print the seating arrangement
                System.out.println("Seating Arrangement for " + branch + " - " + slot);
                printSeatingArrangement(seatingArrangement);
                System.out.println();
            }
        }
    }

    private List<String> readRegisterNumbersFromExcel(String filePath, String sheetName) {
        List<String> registerNumbers = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Assuming the register numbers are in the first column of the specified sheet
            Sheet sheet = workbook.getSheet(sheetName);
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

    private List<List<String>> generateSeatingArrangement(List<String> registerNumbers, int branch1Count, int branch2Count) {
        // Shuffle the register numbers randomly
       // Collections.shuffle(registerNumbers);

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
        for (int col = 0; col < totalColumns; col++) {
            List<String> seatingColumn = new ArrayList<>();
            for (int row = 0; row < branch1Rows; row++) {
                if (branch1Index < branch1Count) {
                    seatingColumn.add(branch1Students.get(branch1Index));
                    branch1Index++;
                }
            }
            seatingArrangement.add(seatingColumn);
        }

        // Assign seats for branch 2
        int branch2Index = 0;
        for (int col = 0; col < totalColumns; col++) {
            List<String> seatingColumn = new ArrayList<>();
            for (int row = 0; row < branch2Rows; row++) {
                if (branch2Index < branch2Count) {
                    seatingColumn.add(branch2Students.get(branch2Index));
                    branch2Index++;
                }
            }
            seatingArrangement.add(seatingColumn);
        }

        return seatingArrangement;
    }

    private void printSeatingArrangement(List<List<String>> seatingArrangement) {
        int numRows = seatingArrangement.get(0).size();
        for (int row = 0; row < numRows; row++) {
            for (List<String> seatingColumn : seatingArrangement) {
                if (row < seatingColumn.size()) {
                    System.out.print(seatingColumn.get(row) + "\t");
                } else {
                    System.out.print("\t");
                }
            }
            System.out.println();
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                ExamSeatingArrangementGUI6 gui = new ExamSeatingArrangementGUI6();
                gui.filePath = "miniproject/seat/src/main/java/com/example/SortedExcel.xlsx";
            }
        });
    }
}
