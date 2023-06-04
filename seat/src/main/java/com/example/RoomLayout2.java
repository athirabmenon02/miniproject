package com.example;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;

public class RoomLayout2 extends JFrame {
    private Map<String, ArrayList<String>> roomLayouts;

    public RoomLayout2() {
        roomLayouts = createRoomLayouts();

        setTitle("Room Layout");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Create a panel for room selection
        JPanel selectionPanel = createSelectionPanel();
        add(selectionPanel, BorderLayout.NORTH);

        pack();
        setExtendedState(JFrame.MAXIMIZED_BOTH);
        setLocationRelativeTo(null);
        setVisible(true);
    }

    private Map<String, ArrayList<String>> createRoomLayouts() {
        Map<String, ArrayList<String>> layouts = new HashMap<>();

        // Room 011
        ArrayList<String> layout011 = new ArrayList<>();
        layout011.add("▭ ▭"); // Bench 1
        layout011.add("▭");    // Middle seat in Bench 2
        layout011.add("▭ ▭"); // Bench 3
        layout011.add("▭ ▭"); // Bench 4
        layout011.add("▭");    // Middle seat in Bench 5
        layout011.add("▭ ▭"); // Bench 6
        layout011.add("▭ ▭"); // Bench 7
        layout011.add("▭");    // Middle seat in Bench 8
        layout011.add("▭ ▭"); // Bench 9
        layout011.add("▭ ▭"); // Bench 10
        layout011.add("▭"); // Bench 11
        layout011.add("▭ ▭"); // Bench 12
        layout011.add("▭ ▭"); // Bench 13
        layout011.add("▭"); // Bench 14
        layout011.add("▭ ▭"); // Bench 15
        layout011.add("▭ ▭"); // Bench 16
        layout011.add("▭"); // Bench 17
        layout011.add("▭ ▭"); // Bench 18
        layout011.add("▭ ▭"); // Bench 19
        layout011.add(" ");    // Empty space
        layout011.add("▭ ▭"); // Bench 20
        layouts.put("011", layout011);

        // Add more room layouts here...

        return layouts;
    }

    private JPanel createSelectionPanel() {
        JPanel selectionPanel = new JPanel();
        selectionPanel.setLayout(new FlowLayout());

        JLabel roomLabel = new JLabel("Select Room: ");
        selectionPanel.add(roomLabel);

        // Create a combo box for room selection
        String[] roomNumbers = {"011", "108", "109", "111", "112", "208", "209", "211", "212", "308", "309", "311", "312"};
        JComboBox<String> roomComboBox = new JComboBox<>(roomNumbers);
        selectionPanel.add(roomComboBox);

        // Create a button to display the layout
        JButton showLayoutButton = new JButton("Show Layout");
        selectionPanel.add(showLayoutButton);

        // Create a panel to display the layout
        JPanel layoutPanel = new JPanel();
        layoutPanel.setLayout(new GridLayout(7, 3));

        showLayoutButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String selectedRoom = (String) roomComboBox.getSelectedItem();
                ArrayList<String> layout = roomLayouts.get(selectedRoom);

                layoutPanel.removeAll();
                for (int i = 0; i < layout.size(); i++) {
                    String bench = layout.get(i);
                    JPanel benchPanel = new JPanel();
                    benchPanel.setLayout(new GridLayout(1, 3));

                    String[] seats = bench.split(" ");
                    for (String seat : seats) {
                        JLabel seatLabel = new JLabel(seat);
                        seatLabel.setHorizontalAlignment(JLabel.CENTER);
                        seatLabel.setFont(new Font(Font.MONOSPACED, Font.BOLD, 20));
                        benchPanel.add(seatLabel);
                    }

                    layoutPanel.add(benchPanel);
                }

                layoutPanel.revalidate();
                layoutPanel.repaint();
            }
        });

        JPanel mainPanel = new JPanel();
        mainPanel.setLayout(new BorderLayout());
        mainPanel.add(selectionPanel, BorderLayout.NORTH);
        mainPanel.add(layoutPanel, BorderLayout.CENTER);

        return mainPanel;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            @Override
            public void run() {
                new RoomLayout2();
            }
        });
    }
}
