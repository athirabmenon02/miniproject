package com.example;

import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExamSeatingArrangementGUI2 extends JFrame implements ActionListener {
    private JTextField examsField;
    private Map<JCheckBox, JComboBox<String>> branchSlotMap;

    public ExamSeatingArrangementGUI2() {
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

            // Perform further logic and functionality based on the selected options
            // Generate the exam seating arrangement accordingly
            generateExamSeatingArrangement(selectedBranches, selectedSlots, numberOfExams);
        }
    }

    private void generateExamSeatingArrangement(List<String> selectedBranches, List<String> selectedSlots, int numberOfExams) {
        try {
            try (Workbook workbook = new XSSFWorkbook()) {
                for (int i = 0; i < numberOfExams; i++) {
                    String branch = selectedBranches.get(i);
                    String slot = selectedSlots.get(i);
                    String sheetName = branch + "_" + slot;

                    Sheet sheet = workbook.createSheet(sheetName);

                    // Load student registration numbers from the Excel sheet
                    List<String> studentRegistrationNumbers = loadStudentRegistrationNumbers();

                    // Shuffle the list of student registration numbers
                    Collections.shuffle(studentRegistrationNumbers);

                    int rowNumber = 0;
                    int seatNumber = 1;

                    while (!studentRegistrationNumbers.isEmpty()) {
                        Row row = sheet.createRow(rowNumber);

                        // Add students to the row
                        int columnNumber = 0;

                        while (columnNumber < 5 && !studentRegistrationNumbers.isEmpty()) {
                            Cell cell = row.createCell(columnNumber);

                            // Get the next student registration number
                            String registrationNumber = studentRegistrationNumbers.remove(0);

                            // Set the student registration number in the cell
                            cell.setCellValue(registrationNumber);

                            // Increment the column number and seat number
                            columnNumber++;
                            seatNumber++;

                            // Skip the middle seat column if it is full
                            if (columnNumber == 2) {
                                columnNumber++;
                            }
                        }

                        rowNumber++;
                    }
                }

                // Save the workbook to a file
                FileOutputStream fileOut = new FileOutputStream("seating_arrangement.xlsx");
                workbook.write(fileOut);
                fileOut.close();
            }

            System.out.println("Seating arrangement saved to seating_arrangement.xlsx");

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private List<String> loadStudentRegistrationNumbers() {
        // Implement your logic to load student registration numbers from the Excel sheet
        // and return them as a list
        // Here, you can use any method or library that allows you to read from Excel files

        // Dummy implementation with hardcoded registration numbers for demonstration
        List<String> registrationNumbers = new ArrayList<>();
        registrationNumbers.add("202101");
        registrationNumbers.add("202102");
        registrationNumbers.add("202103");
        registrationNumbers.add("202104");
        registrationNumbers.add("202105");
        registrationNumbers.add("202106");
        registrationNumbers.add("202107");
        registrationNumbers.add("202108");
        registrationNumbers.add("202109");
        registrationNumbers.add("202110");
        registrationNumbers.add("202111");
        registrationNumbers.add("202112");
        registrationNumbers.add("202113");
        registrationNumbers.add("202114");
        registrationNumbers.add("202115");
        registrationNumbers.add("202116");
        registrationNumbers.add("202117");
        registrationNumbers.add("202118");
        registrationNumbers.add("202119");
        registrationNumbers.add("202120");
        return registrationNumbers;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                new ExamSeatingArrangementGUI2();
            }
        });
    }
}
