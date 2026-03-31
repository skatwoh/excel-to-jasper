package com.example;

import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;

public class ExcelToJasperProApp extends JFrame {

    private JTextField fileField;
    private JTable previewTable;

    private int headerStartRow = -1;
    private int headerEndRow = -1;

    public ExcelToJasperProApp() {

        setTitle("Excel → Jasper PRO Tool");
        setSize(1000, 650);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // ===== TOP PANEL =====
        JPanel topPanel = new JPanel(new BorderLayout(10, 10));
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        fileField = new JTextField();
        JButton browseBtn = new JButton("Browse");
        browseBtn.addActionListener(this::browseFile);

        topPanel.add(fileField, BorderLayout.CENTER);
        topPanel.add(browseBtn, BorderLayout.EAST);

        add(topPanel, BorderLayout.NORTH);

        // ===== TABLE =====
        previewTable = new JTable();
        previewTable.setSelectionMode(ListSelectionModel.SINGLE_INTERVAL_SELECTION);
        previewTable.setRowSelectionAllowed(true);

        previewTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                headerStartRow = previewTable.getSelectionModel().getMinSelectionIndex();
                headerEndRow = previewTable.getSelectionModel().getMaxSelectionIndex();
                previewTable.repaint();
            }
        });

        previewTable.setDefaultRenderer(Object.class, new HeaderHighlightRenderer());

        add(new JScrollPane(previewTable), BorderLayout.CENTER);

        // ===== BOTTOM PANEL =====
        JPanel bottomPanel = new JPanel();

        JButton previewBtn = new JButton("Preview Excel");
        previewBtn.addActionListener(this::previewExcel);

        JButton generateBtn = new JButton("Generate JRXML");
        generateBtn.addActionListener(this::generateReport);

        bottomPanel.add(previewBtn);
        bottomPanel.add(generateBtn);

        add(bottomPanel, BorderLayout.SOUTH);
    }

    // =============================
    // BROWSE FILE
    // =============================
    private void browseFile(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            fileField.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    // =============================
    // PREVIEW EXCEL
    // =============================
    private void previewExcel(ActionEvent e) {
        try {
            FileInputStream fis = new FileInputStream(fileField.getText());
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            int maxColumns = 0;
            for (Row row : sheet) {
                maxColumns = Math.max(maxColumns, row.getPhysicalNumberOfCells());
            }

            DefaultTableModel model = new DefaultTableModel();

            for (int i = 0; i < maxColumns; i++) {
                model.addColumn("Col " + i);
            }

            for (Row row : sheet) {
                Vector<String> rowData = new Vector<>();
                for (int i = 0; i < maxColumns; i++) {
                    Cell cell = row.getCell(i);
                    rowData.add(cell == null ? "" : cell.toString());
                }
                model.addRow(rowData);
            }

            previewTable.setModel(model);

            workbook.close();
            fis.close();

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Preview Error: " + ex.getMessage());
        }
    }

    // =============================
    // GENERATE JRXML
    // =============================
    private void generateReport(ActionEvent e) {
        try {

            if (headerStartRow == -1) {
                JOptionPane.showMessageDialog(this, "Select header rows first!");
                return;
            }

            int headerCount = headerEndRow - headerStartRow + 1;

            String excelPath = fileField.getText();
            String outputPath = new File(excelPath).getParent() + File.separator + "output.jrxml";

            convert(excelPath, outputPath, headerStartRow, headerCount);

            JOptionPane.showMessageDialog(this, "Generated:\n" + outputPath);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error: " + ex.getMessage());
        }
    }

    // =============================
    // CONVERT LOGIC
    // =============================
    public static void convert(String excelPath,
                               String jrxmlPath,
                               int headerStartRow,
                               int headerRowCount) throws Exception {

        FileInputStream fis = new FileInputStream(excelPath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        Row lastHeaderRow = sheet.getRow(headerStartRow + headerRowCount - 1);
        int columnCount = lastHeaderRow.getPhysicalNumberOfCells();

        JasperDesign design = new JasperDesign();
        design.setName("AutoReport");

        List<String> fieldNames = new ArrayList<>();
        Set<String> used = new HashSet<>();

        for (int i = 0; i < columnCount; i++) {
            String raw = lastHeaderRow.getCell(i) == null
                    ? ""
                    : lastHeaderRow.getCell(i).toString().trim();

            String fieldName = raw.isEmpty() ? "COLUMN_" + i : raw.replace(" ", "_");

            String original = fieldName;
            int counter = 1;
            while (used.contains(fieldName)) {
                fieldName = original + "_" + counter++;
            }

            used.add(fieldName);
            fieldNames.add(fieldName);

            JRDesignField field = new JRDesignField();
            field.setName(fieldName);
            field.setValueClass(String.class);
            design.addField(field);
        }

        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        workbook.close();
        fis.close();
    }

    // =============================
    // HEADER HIGHLIGHT RENDERER
    // =============================
    class HeaderHighlightRenderer extends DefaultTableCellRenderer {
        @Override
        public Component getTableCellRendererComponent(JTable table,
                                                       Object value,
                                                       boolean isSelected,
                                                       boolean hasFocus,
                                                       int row,
                                                       int column) {

            Component c = super.getTableCellRendererComponent(
                    table, value, isSelected, hasFocus, row, column);

            if (headerStartRow != -1 &&
                    row >= headerStartRow &&
                    row <= headerEndRow) {

                c.setBackground(new Color(255, 230, 150));
            } else {
                c.setBackground(Color.WHITE);
            }

            return c;
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ExcelToJasperProApp().setVisible(true));
    }
}