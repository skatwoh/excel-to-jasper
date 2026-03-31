package com.example;

import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;

public class ExcelToJasperSwingApp extends JFrame {

    private JTextField fileField;
    private JTextField headerStartField;
    private JTextField headerCountField;
    private JTable previewTable;

    public ExcelToJasperSwingApp() {

        setTitle("Excel → Jasper Converter");
        setSize(900, 600);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // ===== TOP PANEL =====
        JPanel topPanel = new JPanel(new GridLayout(3, 3, 10, 10));
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        topPanel.add(new JLabel("Excel File:"));
        fileField = new JTextField();
        topPanel.add(fileField);

        JButton browseBtn = new JButton("Browse");
        browseBtn.addActionListener(this::browseFile);
        topPanel.add(browseBtn);

        topPanel.add(new JLabel("Header Start Row (0-based):"));
        headerStartField = new JTextField("0");
        topPanel.add(headerStartField);

        topPanel.add(new JLabel("Header Row Count:"));
        headerCountField = new JTextField("2");
        topPanel.add(headerCountField);

        add(topPanel, BorderLayout.NORTH);

        // ===== CENTER PREVIEW TABLE =====
        previewTable = new JTable();
        JScrollPane scrollPane = new JScrollPane(previewTable);
        add(scrollPane, BorderLayout.CENTER);

        // ===== BOTTOM BUTTONS =====
        JPanel bottomPanel = new JPanel();

        JButton previewBtn = new JButton("Preview");
        previewBtn.addActionListener(this::previewExcel);

        JButton generateBtn = new JButton("Generate JRXML");
        generateBtn.addActionListener(this::generateReport);

        bottomPanel.add(previewBtn);
        bottomPanel.add(generateBtn);

        add(bottomPanel, BorderLayout.SOUTH);
    }

    private void browseFile(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            fileField.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    // ===============================
    // PREVIEW EXCEL INTO JTable
    // ===============================
    private void previewExcel(ActionEvent e) {
        try {
            String path = fileField.getText();
            FileInputStream fis = new FileInputStream(path);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);

            int maxColumns = 0;
            for (Row row : sheet) {
                maxColumns = Math.max(maxColumns, row.getPhysicalNumberOfCells());
            }

            DefaultTableModel model = new DefaultTableModel();

            // create column names A, B, C...
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
            JOptionPane.showMessageDialog(this,
                    "Preview Error: " + ex.getMessage());
        }
    }

    // ===============================
    // GENERATE JRXML
    // ===============================
    private void generateReport(ActionEvent e) {
        try {

            String excelPath = fileField.getText();
            int headerStart = Integer.parseInt(headerStartField.getText());
            int headerCount = Integer.parseInt(headerCountField.getText());

            File excelFile = new File(excelPath);

            String fileName = excelFile.getName();
            String baseName = fileName.substring(0, fileName.lastIndexOf("."));

            String outputPath = excelFile.getParent()
                    + File.separator
                    + baseName + ".jrxml";

            convert(excelPath, outputPath, headerStart, headerCount);

            JOptionPane.showMessageDialog(this,
                    "Generated:\n" + outputPath);

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this,
                    "Error: " + ex.getMessage());
        }
    }

    // ===============================
    // CONVERT LOGIC (giống bản trước)
    // ===============================
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
        design.setPageWidth(595);
        design.setPageHeight(842);
        design.setColumnWidth(555);
        design.setLeftMargin(20);
        design.setRightMargin(20);
        design.setTopMargin(20);
        design.setBottomMargin(20);

        int columnWidth = 100;

        List<String> fieldNames = new ArrayList<>();
        Set<String> used = new HashSet<>();

        for (int i = 0; i < columnCount; i++) {

            String raw = lastHeaderRow.getCell(i) == null
                    ? ""
                    : lastHeaderRow.getCell(i).toString().trim();

            String fieldName = raw.isEmpty()
                    ? "COLUMN_" + i
                    : raw.replace(" ", "_");

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

        JRDesignBand headerBand = new JRDesignBand();
        headerBand.setHeight(25 * headerRowCount);

        for (int h = 0; h < headerRowCount; h++) {

            Row headerRow = sheet.getRow(headerStartRow + h);
            int x = 0;

            for (int c = 0; c < columnCount; c++) {

                String text = headerRow.getCell(c) == null
                        ? ""
                        : headerRow.getCell(c).toString();

                JRDesignStaticText staticText = new JRDesignStaticText();
                staticText.setX(x);
                staticText.setY(h * 25);
                staticText.setWidth(columnWidth);
                staticText.setHeight(25);
                staticText.setText(text);
                staticText.setBold(true);
                staticText.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

                headerBand.addElement(staticText);
                x += columnWidth;
            }
        }

        design.setColumnHeader(headerBand);

        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        workbook.close();
        fis.close();
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ExcelToJasperSwingApp().setVisible(true));
    }
}