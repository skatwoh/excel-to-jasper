package com.example;

import com.formdev.flatlaf.FlatDarkLaf;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.Vector;

public class ExcelToJasperFullApp extends JFrame {

    private JTextField fileField;
    private JTable sheetTable;
    private DefaultTableModel tableModel;

    public ExcelToJasperFullApp() {
        setTitle("Excel → Jasper PRO");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 800);
        setLocationRelativeTo(null);

        // Set FlatLaf Dark
        try {
            UIManager.setLookAndFeel(new FlatDarkLaf());
        } catch (Exception ex) {
            System.err.println("Failed to initialize LaF");
        }

        JPanel mainPanel = new JPanel(new BorderLayout());
        mainPanel.setBackground(new Color(30, 30, 30));

        // --- Header ---
        JPanel headerPanel = new JPanel(new BorderLayout());
        headerPanel.setOpaque(false);
        headerPanel.setBorder(new EmptyBorder(15, 20, 15, 20));

        JLabel titleLabel = new JLabel("Excel → Jasper PRO");
        titleLabel.setFont(new Font("SansSerif", Font.BOLD, 18));
        titleLabel.setForeground(Color.WHITE);
        headerPanel.add(titleLabel, BorderLayout.WEST);

        JPanel headerRight = new JPanel(new FlowLayout(FlowLayout.RIGHT, 15, 0));
        headerRight.setOpaque(false);
        headerRight.add(new JLabel("Free"));
        JButton licenseBtn = new JButton("License");
        JButton helpBtn = new JButton("?");
        headerRight.add(licenseBtn);
        headerRight.add(helpBtn);
        headerPanel.add(headerRight, BorderLayout.EAST);

        mainPanel.add(headerPanel, BorderLayout.NORTH);

        // --- Tabs ---
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setBorder(new EmptyBorder(0, 10, 0, 10));

        JPanel importTab = createImportTab();
        tabbedPane.addTab("Import", importTab);
        tabbedPane.addTab("Column mapping", new JPanel());
        tabbedPane.addTab("Preview", new JPanel());
        tabbedPane.addTab("Output", new JPanel());
        tabbedPane.addTab("Help", new JPanel());

        mainPanel.add(tabbedPane, BorderLayout.CENTER);

        // --- Footer ---
        JPanel footerPanel = new JPanel(new BorderLayout());
        footerPanel.setOpaque(false);
        footerPanel.setBorder(new EmptyBorder(10, 20, 10, 20));

        JLabel versionLabel = new JLabel("Excel → Jasper PRO v2.1.0");
        versionLabel.setForeground(Color.GRAY);
        footerPanel.add(versionLabel, BorderLayout.WEST);

        JLabel licenseInfoLabel = new JLabel("Free edition — Activate License");
        licenseInfoLabel.setForeground(Color.GRAY);
        footerPanel.add(licenseInfoLabel, BorderLayout.EAST);

        mainPanel.add(footerPanel, BorderLayout.SOUTH);

        add(mainPanel);
    }

    private JPanel createImportTab() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setBorder(new EmptyBorder(20, 20, 20, 20));
        panel.setOpaque(false);

        // Excel file section
        panel.add(createLabel("Excel file"));
        panel.add(Box.createRigidArea(new Dimension(0, 5)));

        JPanel fileSelectionPanel = new JPanel(new BorderLayout(10, 0));
        fileSelectionPanel.setOpaque(false);
        fileSelectionPanel.setMaximumSize(new Dimension(Integer.MAX_VALUE, 40));

        fileField = new JTextField();
        fileField.setPreferredSize(new Dimension(0, 35));
        JButton browseBtn = new JButton("Browse...");
        browseBtn.addActionListener(e -> browseFile());

        fileSelectionPanel.add(fileField, BorderLayout.CENTER);
        fileSelectionPanel.add(browseBtn, BorderLayout.EAST);
        panel.add(fileSelectionPanel);

        panel.add(Box.createRigidArea(new Dimension(0, 20)));

        // Sheets detected section
        panel.add(createLabel("Sheets detected"));
        panel.add(Box.createRigidArea(new Dimension(0, 10)));

        String[] columnNames = {"#", "Sheet name", "Rows", "Cols", "Action"};
        tableModel = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return column == 4; // Only action column editable (for buttons if we use them)
            }
        };
        sheetTable = new JTable(tableModel);
        sheetTable.setRowHeight(35);
        sheetTable.getTableHeader().setPreferredSize(new Dimension(0, 35));

        // Custom renderer for Action column
        sheetTable.getColumnModel().getColumn(4).setCellRenderer(new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                JLabel label = (JLabel) super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                label.setHorizontalAlignment(JLabel.CENTER);
                if ("Selected".equals(value)) {
                    label.setBackground(new Color(60, 100, 150));
                    label.setForeground(Color.WHITE);
                } else if ("Select".equals(value)) {
                    label.setBackground(new Color(50, 50, 50));
                    label.setForeground(Color.LIGHT_GRAY);
                }
                return label;
            }
        });

        sheetTable.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent e) {
                int row = sheetTable.rowAtPoint(e.getPoint());
                int col = sheetTable.columnAtPoint(e.getPoint());
                if (col == 4) {
                    for (int i = 0; i < tableModel.getRowCount(); i++) {
                        tableModel.setValueAt("Select", i, 4);
                    }
                    tableModel.setValueAt("Selected", row, 4);
                }
            }
        });

        JScrollPane scrollPane = new JScrollPane(sheetTable);
        scrollPane.setPreferredSize(new Dimension(0, 150));
        scrollPane.setMaximumSize(new Dimension(Integer.MAX_VALUE, 200));
        panel.add(scrollPane);

        panel.add(Box.createRigidArea(new Dimension(0, 20)));

        // Header color section
        panel.add(createLabel("Header color"));
        panel.add(Box.createRigidArea(new Dimension(0, 10)));

        JPanel colorPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
        colorPanel.setOpaque(false);
        colorPanel.setMaximumSize(new Dimension(Integer.MAX_VALUE, 40));

        JPanel colorBox = new JPanel();
        colorBox.setPreferredSize(new Dimension(30, 30));
        colorBox.setBackground(new Color(211, 209, 199)); // #D3D1C7
        colorBox.setBorder(BorderFactory.createLineBorder(Color.GRAY));

        colorPanel.add(colorBox);
        colorPanel.add(new JLabel("#D3D1C7 — Light Gray"));
        JButton changeColorBtn = new JButton("Change color");
        colorPanel.add(changeColorBtn);
        panel.add(colorPanel);

        panel.add(Box.createRigidArea(new Dimension(0, 30)));

        // Next button
        JButton nextBtn = new JButton("Next: Column mapping →");
        nextBtn.setFont(new Font("SansSerif", Font.BOLD, 14));
        nextBtn.setPreferredSize(new Dimension(200, 40));
        nextBtn.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(nextBtn);

        panel.add(Box.createVerticalGlue());

        return panel;
    }

    private JLabel createLabel(String text) {
        JLabel label = new JLabel(text);
        label.setForeground(Color.GRAY);
        label.setAlignmentX(Component.LEFT_ALIGNMENT);
        return label;
    }

    private void browseFile() {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File selectedFile = chooser.getSelectedFile();
            fileField.setText(selectedFile.getAbsolutePath());
            loadSheets(selectedFile);
        }
    }

    private void loadSheets(File file) {
        tableModel.setRowCount(0);
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                int rows = sheet.getPhysicalNumberOfRows();
                int cols = 0;
                if (rows > 0) {
                    for (Row row : sheet) {
                        cols = Math.max(cols, row.getPhysicalNumberOfCells());
                    }
                }

                Vector<Object> rowData = new Vector<>();
                rowData.add(i + 1);
                rowData.add(sheet.getSheetName());
                rowData.add(rows);
                rowData.add(cols);
                rowData.add(i == 0 ? "Selected" : "Select");
                tableModel.addRow(rowData);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error loading Excel file: " + ex.getMessage());
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            new ExcelToJasperFullApp().setVisible(true);
        });
    }
}
