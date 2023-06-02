package com.example;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExamSeatingArrangementGUI extends JFrame implements ActionListener {
    private JTextField examsField;
    private Map<JCheckBox, JComboBox<String>> branchSlotMap;

    public ExamSeatingArrangementGUI() {
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
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                new ExamSeatingArrangementGUI();
            }
        });
    }
}
