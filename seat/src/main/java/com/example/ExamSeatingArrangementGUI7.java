package com.example;

import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExamSeatingArrangementGUI7 extends JFrame implements ActionListener {
    private JTextField examsField;
    private Map<JCheckBox, JComboBox<String>> branchSlotMap;
    private String filePath;

    public ExamSeatingArrangementGUI7() {
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

            // Check if the total number of selected students is greater than the number of available rooms
            int totalSelectedStudents = selectedBranches.size() + selectedSlots.size();
            int totalRooms = 12; // Total number of available rooms

            if (totalSelectedStudents > totalRooms) {
                JOptionPane.showMessageDialog(this, "Not enough rooms available to allocate all the students.",
                        "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Generate seating arrangement based on selected branches and slots
            List<String> registerNumbers = new ArrayList<>();
            for (int i = 0; i < selectedBranches.size(); i++) {
                String branch = selectedBranches.get(i);
                String slot = selectedSlots.get(i);

                // Assuming the Excel sheet contains a sheet named "branch-slot"
                String sheetName = branch + "-" + slot;

                // Read register numbers from the Excel sheet
                List<String> sheetRegisterNumbers = readRegisterNumbersFromExcel(filePath, sheetName);
                registerNumbers.addAll(sheetRegisterNumbers);
            }

            // Check if the total number of selected students is greater than the number of available classrooms
            int totalStudents = registerNumbers.size();
            int totalClassrooms = 12; // Total number of available classrooms

            if (totalStudents > (totalClassrooms * 30)) {
                JOptionPane.showMessageDialog(this, "Not enough classrooms available to allocate all the students.",
                        "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // Allocate students to classrooms
            List<List<String>> classrooms = allocateStudentsToClassrooms(registerNumbers);

            // Print the seating arrangement
            System.out.println("Seating Arrangement:");
            for (int i = 0; i < classrooms.size(); i++) {
                System.out.println("Classroom " + (i + 1) + ":");
                List<String> classroom = classrooms.get(i);
                for (int j = 0; j < classroom.size(); j++) {
                    System.out.print(classroom.get(j) + "\t");
                    if ((j + 1) % 6 == 0) {
                        System.out.println();
                    }
                }
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

    private List<List<String>> allocateStudentsToClassrooms(List<String> registerNumbers) {
        List<List<String>> classrooms = new ArrayList<>();
        int totalClassrooms = 12; // Total number of available classrooms
        int studentsPerClassroom = 30; // Number of students per classroom
        int columnsPerClassroom = 5; // Number of columns per classroom
        int rowsPerClassroom = 6; // Number of rows per classroom
    
        // Calculate the total number of required classrooms based on the number of students
        int requiredClassrooms = (int) Math.ceil((double) registerNumbers.size() / studentsPerClassroom);
    
        // Adjust the total number of classrooms if necessary
        totalClassrooms = Math.min(totalClassrooms, requiredClassrooms);
    
        // Create classrooms
        for (int i = 0; i < totalClassrooms; i++) {
            classrooms.add(new ArrayList<>());
        }
    
        // Allocate students to classrooms
        int currentClassroomIndex = 0;
        int currentColumnIndex = 0;
        int currentRowIndex = 0;
        int requiredITStudents = 18;
        int requiredCSStudents = 12;
    
        // Allocate IT students
        int itCount = 0;
        while (itCount < requiredITStudents) {
            String registerNumber = registerNumbers.get(itCount);
            List<String> classroom = classrooms.get(currentClassroomIndex);
            classroom.add(registerNumber);
    
            currentRowIndex++;
            if (currentRowIndex == rowsPerClassroom) {
                currentRowIndex = 0;
                currentColumnIndex++;
            }
            if (currentColumnIndex == columnsPerClassroom) {
                currentColumnIndex = 0;
                currentClassroomIndex++;
            }
            itCount++;
        }
    
        // Allocate CS students
        int csCount = 0;
        while (csCount < requiredCSStudents) {
            String registerNumber = registerNumbers.get(itCount + csCount);
            List<String> classroom = classrooms.get(currentClassroomIndex);
            classroom.add(registerNumber);
    
            currentRowIndex++;
            if (currentRowIndex == rowsPerClassroom) {
                currentRowIndex = 0;
                currentColumnIndex++;
            }
            if (currentColumnIndex == columnsPerClassroom) {
                currentColumnIndex = 0;
                currentClassroomIndex++;
            }
            csCount++;
        }
    
        // Fill remaining classrooms with empty seats if necessary
        for (int i = 0; i < totalClassrooms; i++) {
            List<String> classroom = classrooms.get(i);
            int remainingSeats = studentsPerClassroom - classroom.size();
    
            for (int j = 0; j < remainingSeats; j++) {
                classroom.add("Empty Seat");
            }
        }
    
        return classrooms;
    }
    
    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                ExamSeatingArrangementGUI7 gui = new ExamSeatingArrangementGUI7();
                gui.filePath = "miniproject/seat/src/main/java/com/example/SortedExcel.xlsx";
            }
        });
    }
}
