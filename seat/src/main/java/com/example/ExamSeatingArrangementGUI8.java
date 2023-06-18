package com.example;

import java.awt.Dimension;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExamSeatingArrangementGUI8 extends JFrame implements ActionListener {
    private JTextField examsField;
    private List<String> selectedSheets;

    public ExamSeatingArrangementGUI8() {
        setTitle("Exam Seating Arrangement");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new GridLayout(0, 2));

        JLabel examsLabel = new JLabel("Number of Exams: ");
        examsField = new JTextField(10);

        JButton selectSheetsButton = new JButton("Select Sheets");
        selectSheetsButton.addActionListener(this);

        JButton displayButton = new JButton("Display");
        displayButton.addActionListener(this);

        selectedSheets = new ArrayList<>();

        add(examsLabel);
        add(examsField);
        add(selectSheetsButton);
        add(displayButton);

        pack();
        setLocationRelativeTo(null);
        setVisible(true);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getActionCommand().equals("Select Sheets")) {
            JFileChooser fileChooser = new JFileChooser();
            int result = fileChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                String filePath = fileChooser.getSelectedFile().getAbsolutePath();
                selectedSheets = selectSheets(filePath);
            }
        } else if (e.getActionCommand().equals("Display")) {
            int numberOfExams = Integer.parseInt(examsField.getText());
            if (selectedSheets.isEmpty() || numberOfExams <= 0) {
                JOptionPane.showMessageDialog(this, "Please select sheets and provide a valid number of exams.");
                return;
            }

            List<String> seatingArrangement = generateSeatingArrangement(selectedSheets, numberOfExams);
            displaySeatingArrangement(seatingArrangement);
        }
    }

    private List<String> selectSheets(String filePath) {
        List<String> selectedSheets = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            String[] sheetNames = new String[numberOfSheets];
            for (int i = 0; i < numberOfSheets; i++) {
                sheetNames[i] = workbook.getSheetName(i);
            }

            String selectedOptions = (String) JOptionPane.showInputDialog(this,
                    "Select the sheets for seating arrangement:",
                    "Select Sheets", JOptionPane.PLAIN_MESSAGE, null,
                    sheetNames, sheetNames[0]);

            if (selectedOptions != null) {
                selectedSheets.add(selectedOptions);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return selectedSheets;
    }

    private List<String> generateSeatingArrangement(List<String> selectedSheets, int numberOfExams) {
        List<String> seatingArrangement = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream("miniproject/seat/src/main/java/com/example/SortedExcel.xlsx");
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (String sheetName : selectedSheets) {
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet != null) {
                    seatingArrangement.add(sheetName);
                    Iterator<Row> rowIterator = sheet.iterator();

                    while (rowIterator.hasNext()) {
                        Row row = rowIterator.next();
                        seatingArrangement.add(row.toString());
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return seatingArrangement;
    }

    private void displaySeatingArrangement(List<String> seatingArrangement) {
        StringBuilder sb = new StringBuilder();
        sb.append("Seating Arrangement:\n");

        for (String student : seatingArrangement) {
            sb.append(student).append("\n");
        }

        JTextArea textArea = new JTextArea(sb.toString());
        JScrollPane scrollPane = new JScrollPane(textArea);
        scrollPane.setPreferredSize(new Dimension(400, 300));

        JOptionPane.showMessageDialog(this, scrollPane, "Seating Arrangement", JOptionPane.PLAIN_MESSAGE);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                new ExamSeatingArrangementGUI8();

            }
        });
    }
}
