package com.example;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import javax.swing.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.List;
import java.awt.Color;
import javax.swing.border.AbstractBorder;
class RoundedBorder extends AbstractBorder {
    private int radius;
    private Color color;

    public RoundedBorder(int radius, Color color) {
        this.radius = radius;
        this.color = color;
    }

    @Override
    public void paintBorder(Component c, Graphics g, int x, int y, int width, int height) {
        Graphics2D g2d = (Graphics2D) g.create();
        g2d.setColor(color);
        g2d.drawRoundRect(x, y, width - 1, height - 1, radius, radius);
        g2d.dispose();
    }

    @Override
    public Insets getBorderInsets(Component c) {
        return new Insets(radius, radius, radius, radius);
    }

    @Override
    public Insets getBorderInsets(Component c, Insets insets) {
        insets.left = insets.right = insets.top = insets.bottom = radius;
        return insets;
    }
}
public class ExcelDataWriterGUI{
    static int Examsnum,hall;
     Workbook destinationWorkbook;
    private static Map<String, String> examSlots = new HashMap<>();
    private static List<String> selectedKeys = new ArrayList<>();

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowGUI();
            }
        });
    }
private static int totalHall(Sheet sheet) {
    int hallCount = 0;
    int startingRow = 6;
    int columnIndex = 1;

    Row row = sheet.getRow(startingRow);
    while (row != null) {
        Cell cell = row.getCell(columnIndex);
        if (cell != null && cell.getCellType() != CellType.BLANK) {
            hallCount++;
        }
        startingRow += 21;
        row = sheet.getRow(startingRow);
    }

    return hallCount;
}

private static void createAndShowGUI() {
    // Set the Nimbus look and feel for a modern appearance
    try {
        UIManager.setLookAndFeel("javax.swing.plaf.nimbus.NimbusLookAndFeel");
    } catch (Exception e) {
        e.printStackTrace();
    }

    // Create the main frame
    JFrame frame = new JFrame("Exam Selection");
    frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

    // Create the main panel with a blue background
    JPanel mainPanel = new JPanel(new GridBagLayout());
   // mainPanel.setBackground(new Color(200, 220, 240)); // Light blue background color
    mainPanel.setBackground(new Color(29,51,84)); 
    // Create a container panel to hold the components
    JPanel containerPanel = new JPanel(new GridBagLayout());
    containerPanel.setBackground(new Color(70,117,153));; // Set the background color of the container panel

    GridBagConstraints gbcContainer = new GridBagConstraints();
    gbcContainer.gridx = 0;
    gbcContainer.gridy = GridBagConstraints.RELATIVE;
    gbcContainer.insets = new Insets(10, 10, 10, 10);
    gbcContainer.fill = GridBagConstraints.HORIZONTAL;

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

    // Create the submit button with a gradient background
    JButton submitButton = new JButton("Submit");
    submitButton.setBackground(new Color(255,255,255));
    submitButton.setForeground(Color.BLACK);
    submitButton.setFocusPainted(false); // Remove the focus border
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
                processExcelData();
                openClassroomInputDialog(hall);
                frame.dispose();
            }
        });
        
       // Add components to the container panel
       GridBagConstraints gbcLabel = new GridBagConstraints();
       gbcLabel.anchor = GridBagConstraints.WEST;
       gbcLabel.gridx = 0;
       gbcLabel.gridy = GridBagConstraints.RELATIVE;
       gbcLabel.insets = new Insets(10, 10, 5, 10);
       gbcLabel.gridwidth = 1;

       GridBagConstraints gbcField = new GridBagConstraints();
       gbcField.anchor = GridBagConstraints.WEST;
       gbcField.gridx = 1;
       gbcField.gridy = GridBagConstraints.RELATIVE;
       gbcField.insets = new Insets(10, 0, 5, 10);
       gbcField.fill = GridBagConstraints.HORIZONTAL;

       GridBagConstraints gbcCheckbox = new GridBagConstraints();
       gbcCheckbox.anchor = GridBagConstraints.WEST;
       gbcCheckbox.gridx = 0;
       gbcCheckbox.gridy = GridBagConstraints.RELATIVE;
       gbcCheckbox.insets = new Insets(0, 0, 5, 10);

       GridBagConstraints gbcSlot = new GridBagConstraints();
       gbcSlot.anchor = GridBagConstraints.WEST;
       gbcSlot.gridx = 1;
       gbcSlot.gridy = GridBagConstraints.RELATIVE;
       gbcSlot.insets = new Insets(0, 0, 5, 10);

       GridBagConstraints gbcSubmitButton = new GridBagConstraints();
       gbcSubmitButton.gridx = 0;
       gbcSubmitButton.gridy = GridBagConstraints.RELATIVE;
       gbcSubmitButton.gridwidth = 2;
       gbcSubmitButton.insets = new Insets(10, 10, 10, 10);
       gbcSubmitButton.anchor = GridBagConstraints.CENTER;

       containerPanel.add(examLabel, gbcLabel);
       containerPanel.add(examTextField, gbcField);

       // Add spacing between the rows
       containerPanel.add(Box.createRigidArea(new Dimension(0, 10)));

       containerPanel.add(csCheckBox, gbcCheckbox);
       containerPanel.add(csSlotComboBox, gbcSlot);

       containerPanel.add(itCheckBox, gbcCheckbox);
       containerPanel.add(itSlotComboBox, gbcSlot);

       containerPanel.add(ecCheckBox, gbcCheckbox);
       containerPanel.add(ecSlotComboBox, gbcSlot);

       containerPanel.add(aeiCheckBox, gbcCheckbox);
       containerPanel.add(aeiSlotComboBox, gbcSlot);

       containerPanel.add(ereCheckBox, gbcCheckbox);
       containerPanel.add(ereSlotComboBox, gbcSlot);

       containerPanel.add(ceCheckBox, gbcCheckbox);
       containerPanel.add(ceSlotComboBox, gbcSlot);

       // Add spacing between the rows
       containerPanel.add(Box.createRigidArea(new Dimension(0, 20)));

       // Add the submit button with proper padding and rounded corners
       submitButton.setBorder(BorderFactory.createEmptyBorder(10, 25, 10, 25));
       submitButton.setBorder(BorderFactory.createLineBorder(new Color(0,0,0), 2));
      // submitButton.setBorder(BorderFactory.createRoundedBorder(15, new Color(0, 102, 204)));
      submitButton.setBorder(new RoundedBorder(15, new Color(127,127,127))); 
      containerPanel.add(submitButton, gbcSubmitButton);

       // Add the container panel to the main panel
       GridBagConstraints gbcMain = new GridBagConstraints();
       gbcMain.gridx = 0;
       gbcMain.gridy = GridBagConstraints.RELATIVE;
       gbcMain.insets = new Insets(20, 20, 20, 20);
       mainPanel.add(containerPanel, gbcMain);

       frame.add(mainPanel, BorderLayout.CENTER);

       // Set the preferred size of the frame to half of the screen
       Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
       int preferredWidth = screenSize.width / 2;
       int preferredHeight = screenSize.height / 2;
       frame.setPreferredSize(new Dimension(preferredWidth, preferredHeight));

       // Display the frame
       frame.pack();
       frame.setLocationRelativeTo(null);
       frame.setVisible(true);
   }



    private static void processExcelData() {
        String sourceFilePath = "SortedExcel.xlsx";
        String destinationFilePath = "sample1.xlsx";

        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             Workbook sourceWorkbook = new XSSFWorkbook(fis);
             Workbook destinationWorkbook = new XSSFWorkbook(new FileInputStream("template.xlsx"));
             FileOutputStream fos = new FileOutputStream(destinationFilePath)) {
            
            Map<String, List<String>> sheetColumnDataMap = new HashMap<>();

            for (int sheetIndex = 0; sheetIndex < sourceWorkbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sourceSheet = sourceWorkbook.getSheetAt(sheetIndex);
                String sheetName = sourceSheet.getSheetName();
                List<String> columnData = getColumnData(sourceSheet, 0);
                sheetColumnDataMap.put(sheetName, columnData);
            }

            int row = 6, count = 0, b1 = 0, b2 = 0, rowEnd = 11;
            int totalStudents = 0, k = 0, k2 =Examsnum-1,flag1 =0,flag2=0;

            for (String key : selectedKeys) {
                List<String> columnData = sheetColumnDataMap.get(key);
                totalStudents += columnData.size();
            }

            int classrooms = totalStudents / 30;
            System.out.println("Total students appearing for exam= "+ totalStudents);

             for (int j = 0; j < classrooms+3; j++) {
                List<String> columnData;

                 while (count < 18) {
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
                                desiredKey = selectedKeys.get(++k);
                                if(flag2==1)
                                    break;
                                if (k2 == k)
                                {
                                    flag1 = 1;
                                    break;
                                }
                                columnData = sheetColumnDataMap.get(desiredKey);
                                b1 = i = 0;
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
                    if(flag1 ==1)
                        break;
                }
                count = 18;
                while (count < 30) {
                    if(flag2 == 1)
                        break;
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
                   
                                desiredKey = selectedKeys.get(--k2);
                               if (k == k2)
                                {
                                    flag2 = 1;
                                    break;
                                }
                                columnData = sheetColumnDataMap.get(desiredKey);
                                b2 = i = 0;
                            }
                            rowIndex++;
                            if (rowIndex > rowEnd) {
                                rowIndex = row; // Reset row index to 6 (B7) when reaching the end of the range
                                columnIndex += 4; // Move to the next column (E or G)
                            }
                            if (count >= 30)
                            {
                                if(flag1==1){
                                    b1=b2;  
                                    flag2=1;  
                                 }                               
                                break;
                            }
                        }
                        break;
                    }
                    if(flag2 == 1)
                        break;
                }
                count = 0;
                row += 21;
                rowEnd += 21;
          
        }
            destinationWorkbook.write(fos);
            System.out.println("Data written to the destination file successfully.");

            Sheet destinationSheet = destinationWorkbook.getSheet("AAA");
            hall = totalHall(destinationSheet);
            System.out.println("Number of halls in AAA sheet: " + hall);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    

private static void openClassroomInputDialog(int hall) {
    String filePath = "sample1.xlsx";
    JFrame inputFrame = new JFrame("Classroom Input");
    inputFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
  //  inputFrame.setSize(400, 200);
    inputFrame.setLayout(new GridBagLayout());
    inputFrame.getContentPane().setBackground(new Color(29, 51, 84)); // Set the background color

    JLabel label = new JLabel("Required number of classrooms: " + hall);
    JLabel label2 = new JLabel("Enter the classrooms (separated by spaces): ");
    JLabel dateLabel = new JLabel("Enter the date (DD-MM-YYYY): ");
    JTextField dateTextField = new JTextField(15);
    JTextField textField = new JTextField(30);

    JButton submitButton = new JButton("Submit");
    List<Integer> hallsList = new ArrayList<>();

    submitButton.addActionListener(new ActionListener() {
        public void actionPerformed(ActionEvent e) {
            String input = textField.getText();
            String[] classroomStrings = input.split(" ");
            
            if (classroomStrings.length != hall) {
                JOptionPane.showMessageDialog(inputFrame,
                        "Please enter exactly " + hall + " classrooms!",
                        "Invalid Input",
                        JOptionPane.ERROR_MESSAGE);
                return;
            }
            String enteredDate = dateTextField.getText();
            // Add date validation logic if needed

            for (String classroomString : classroomStrings) {
                try {
                    int enteredClassroom = Integer.parseInt(classroomString);
                    hallsList.add(enteredClassroom);
                } catch (NumberFormatException ex) {
                    JOptionPane.showMessageDialog(inputFrame,
                            "Invalid input. Please enter valid numbers.",
                            "Invalid Input",
                            JOptionPane.ERROR_MESSAGE);
                    return;
                }
            }

            inputFrame.dispose();

                    try(FileInputStream inputStream = new FileInputStream(filePath);
                    Workbook destinationWorkbook = new XSSFWorkbook(inputStream)){

                    int row=3,columnh = 6,columnd=1;
                    Integer[] halls = hallsList.toArray(new Integer[hallsList.size()]);
                    Sheet destinationSheet = destinationWorkbook.getSheet("AAA");
                    if (destinationSheet == null) {
                        destinationSheet = destinationWorkbook.createSheet("AAA");
                    }
                    for(int i=0;i<hall;i++)
                    {
                         Row destinationRow = destinationSheet.getRow(row);
                            if (destinationRow == null) {
                                destinationRow = destinationSheet.createRow(row);
                            }

                            Cell destinationCellh = destinationRow.getCell(columnh);
                            if (destinationCellh == null) {
                                destinationCellh = destinationRow.createCell(columnh);
                            }
                            Cell destinationCelld = destinationRow.getCell(columnd);
                            if (destinationCelld == null) {
                                destinationCelld = destinationRow.createCell(columnd);
                            }

                            destinationCellh.setCellValue("HALL  "+halls[i]);
                            destinationCelld.setCellValue(enteredDate);
                            row+=21;
                    }
                    try(FileOutputStream outpuStream = new FileOutputStream(filePath)){
                        destinationWorkbook.write(outpuStream);
                    }
                }catch(IOException e1){
                    e1.printStackTrace();
                }
                    // Now hallsList contains the entered classroom numbers
                    // You can continue with your processing using hallsList
                }
                        /*private static boolean isValidDate(String date) {
                        String dateFormat = "\\d{4}-\\d{2}-\\d{2}";
                        return date.matches(dateFormat);
                    }*/
            });
        
        JPanel inputPanel = new JPanel(new GridBagLayout());
    GridBagConstraints constraints = new GridBagConstraints();
    constraints.fill = GridBagConstraints.HORIZONTAL;
    constraints.weightx = 1;
    constraints.weighty = 1;
    constraints.insets = new Insets(5, 5, 5, 5);

    constraints.gridx = 0;
    constraints.gridy = 0;
    inputPanel.add(label, constraints);

    constraints.gridx = 0;
    constraints.gridy = 1;
    inputPanel.add(label2, constraints);

    constraints.gridx = 0;
    constraints.gridy = 2;
    inputPanel.add(textField, constraints);

    constraints.gridx = 0;
    constraints.gridy = 3;
    inputPanel.add(dateLabel, constraints);

    constraints.gridx = 0;
    constraints.gridy = 4;
    inputPanel.add(dateTextField, constraints);

    constraints.gridx = 0;
    constraints.gridy = 5;
    inputPanel.add(submitButton, constraints);
     

    inputFrame.add(inputPanel); // Add the inputPanel to the inputFrame
    // Set the preferred size of the frame to half of the screen
       Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
       int preferredWidth = screenSize.width / 2;
       int preferredHeight = screenSize.height / 2;
       inputFrame.setPreferredSize(new Dimension(preferredWidth, preferredHeight));
       inputFrame.pack();
    inputFrame.setLocationRelativeTo(null); // Center the inputFrame on the screen
    inputFrame.setVisible(true); // Make the inputFrame visible
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