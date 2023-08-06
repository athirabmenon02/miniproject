package com.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;

public class CombinedGUI {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            JFrame frame = new JFrame("Combined GUI");
            frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            frame.setSize(400, 200);
            frame.setLayout(new BorderLayout());

            JLabel label = new JLabel("Click the button to allocate halls");
            JButton button = new JButton("Allocate Halls");
            button.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    frame.dispose();
                    com.example.SemesterGUI.main(new String[0]);
                    com.example.writer.ExcelDataWriterWithGUI.main(new String[0]);
                }
            });

            JPanel panel = new JPanel();
            panel.setLayout(new GridBagLayout());
            GridBagConstraints constraints = new GridBagConstraints();
            constraints.fill = GridBagConstraints.HORIZONTAL;
            constraints.weightx = 1;
            constraints.weighty = 1;
            constraints.insets = new Insets(5, 5, 5, 5);

            constraints.gridx = 0;
            constraints.gridy = 0;
            panel.add(label, constraints);

            constraints.gridx = 0;
            constraints.gridy = 1;
            panel.add(button, constraints);

            frame.add(panel, BorderLayout.CENTER);
            frame.setVisible(true);
        });
    }
}
