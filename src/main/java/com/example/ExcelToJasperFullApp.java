package com.example;

import net.sf.jasperreports.components.table.*;
import net.sf.jasperreports.engine.component.ComponentKey;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.*;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.formdev.flatlaf.FlatDarkLaf;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.Normalizer;
import java.util.*;
import java.util.List;

public class ExcelToJasperFullApp extends JFrame {

    private JTable mappingTable;
    private JTable previewTable;
    private JList<String> sheetList;
    private JTextField fileField;
    private JTextField licenseField;
    private JLabel statusLabel;
    private File selectedFile;

    private int usageCount = 0;
    private static final int FREE_LIMIT = 5;

    private Color headerColor = Color.LIGHT_GRAY;

    private ExcelToJasperService excelToJasperService = new ExcelToJasperService();

    public ExcelToJasperFullApp() {
        setTitle("Excel → Jasper PRO UI");
        setSize(1400, 850);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLocationRelativeTo(null);
        initUI();
    }

    private void initUI() {
        JPanel main = new JPanel(new BorderLayout());
        main.setBackground(new Color(30, 30, 30));

        // --- Header Bar ---
        JPanel headerBar = new JPanel(new BorderLayout());
        headerBar.setBackground(new Color(45, 45, 48));
        headerBar.setBorder(new EmptyBorder(15, 20, 15, 20));

        JLabel titleLabel = new JLabel("EXCEL TO JASPER PRO");
        titleLabel.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 22));
        titleLabel.setForeground(Color.WHITE);

        JPanel licensePanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        licensePanel.setOpaque(false);

        licenseField = new JTextField(15);
        licenseField.putClientProperty("JTextField.placeholderText", "Nhập License Key...");
        licenseField.putClientProperty("FlatLaf.style", "arc: 10");

        statusLabel = new JLabel("Trial: 0/" + FREE_LIMIT);
        statusLabel.setForeground(new Color(200, 200, 200));

        licensePanel.add(new JLabel("License: "));
        licensePanel.add(licenseField);
        licensePanel.add(statusLabel);

        headerBar.add(titleLabel, BorderLayout.WEST);
        headerBar.add(licensePanel, BorderLayout.EAST);

        // --- Top Dashboard: File & Settings ---
        JPanel settingsCard = new JPanel(new GridBagLayout());
        settingsCard.putClientProperty("FlatLaf.style", "arc: 15");
        settingsCard.setBorder(new EmptyBorder(20, 20, 20, 20));

        JButton chooseBtn = new JButton("Chọn File Excel");
        chooseBtn.putClientProperty("JButton.buttonType", "roundRect");
        chooseBtn.setBackground(new Color(0, 122, 204));
        chooseBtn.setForeground(Color.WHITE);

        JButton colorBtn = new JButton("Màu Header Report");
        colorBtn.putClientProperty("JButton.buttonType", "roundRect");

        fileField = new JTextField();
        fileField.setEditable(false);
        fileField.putClientProperty("JTextField.placeholderText", "Đường dẫn file Excel...");
        fileField.putClientProperty("FlatLaf.style", "arc: 10");

        chooseBtn.addActionListener(e -> chooseFile());
        colorBtn.addActionListener(e -> {
            Color chosen = JColorChooser.showDialog(this, "Chọn màu header", headerColor);
            if (chosen != null) headerColor = chosen;
        });

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 10, 5, 10);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0;
        settingsCard.add(chooseBtn, gbc);

        gbc.gridx = 1; gbc.gridy = 0; gbc.weightx = 1;
        settingsCard.add(fileField, gbc);

        gbc.gridx = 2; gbc.gridy = 0; gbc.weightx = 0;
        settingsCard.add(colorBtn, gbc);

        // --- Sidebar: Sheet Selection ---
        sheetList = new JList<>();
        sheetList.setFixedCellHeight(40);
        sheetList.setFont(new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 14));
        sheetList.setSelectionBackground(new Color(0, 122, 204, 100));
        sheetList.addListSelectionListener(e -> loadData());

        JScrollPane sheetScroll = new JScrollPane(sheetList);
        sheetScroll.setBorder(BorderFactory.createEmptyBorder());

        JPanel sidebarPanel = new JPanel(new BorderLayout());
        sidebarPanel.setBackground(new Color(37, 37, 38));
        sidebarPanel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(0, 0, 0, 1, new Color(63, 63, 70)),
                new EmptyBorder(10, 0, 10, 0)));

        JLabel sidebarTitle = new JLabel("  DANH SÁCH SHEET");
        sidebarTitle.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12));
        sidebarTitle.setForeground(new Color(150, 150, 150));
        sidebarTitle.setPreferredSize(new Dimension(200, 30));

        sidebarPanel.add(sidebarTitle, BorderLayout.NORTH);
        sidebarPanel.add(sheetScroll, BorderLayout.CENTER);

        // --- Main Content: Preview & Mapping ---
        previewTable = new JTable();
        previewTable.setRowHeight(25);
        previewTable.setShowGrid(true);
        previewTable.setGridColor(new Color(63, 63, 70));

        JScrollPane previewScroll = new JScrollPane(previewTable);
        previewScroll.setBorder(BorderFactory.createLineBorder(new Color(63, 63, 70)));

        JPanel previewCard = new JPanel(new BorderLayout());
        previewCard.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEmptyBorder(10, 10, 10, 10), "XEM TRƯỚC DỮ LIỆU",
                TitledBorder.LEFT, TitledBorder.TOP, new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12)));
        previewCard.add(previewScroll, BorderLayout.CENTER);

        mappingTable = new JTable();
        mappingTable.setRowHeight(30);

        JScrollPane mappingScroll = new JScrollPane(mappingTable);
        mappingScroll.setBorder(BorderFactory.createLineBorder(new Color(63, 63, 70)));

        JPanel mappingCard = new JPanel(new BorderLayout());
        mappingCard.setBorder(BorderFactory.createTitledBorder(
                BorderFactory.createEmptyBorder(10, 10, 10, 10), "CẤU HÌNH MAPPING CỘT",
                TitledBorder.LEFT, TitledBorder.TOP, new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 12)));
        mappingCard.add(mappingScroll, BorderLayout.CENTER);

        JSplitPane rightSplit = new JSplitPane(JSplitPane.VERTICAL_SPLIT, previewCard, mappingCard);
        rightSplit.setDividerLocation(300);
        rightSplit.setResizeWeight(0.4);
        rightSplit.setBorder(BorderFactory.createEmptyBorder());

        JPanel centerWrapper = new JPanel(new BorderLayout());
        centerWrapper.add(settingsCard, BorderLayout.NORTH);
        centerWrapper.add(rightSplit, BorderLayout.CENTER);

        JSplitPane mainSplit = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, sidebarPanel, centerWrapper);
        mainSplit.setDividerLocation(250);
        mainSplit.setBorder(BorderFactory.createEmptyBorder());

        // --- Footer Action Bar ---
        JButton convertBtn = new JButton("GENERATE JASPER REPORT");
        convertBtn.putClientProperty("JButton.buttonType", "roundRect");
        convertBtn.setBackground(new Color(34, 139, 34)); // Forest Green
        convertBtn.setForeground(Color.WHITE);
        convertBtn.setFont(new java.awt.Font("Segoe UI", java.awt.Font.BOLD, 15));
        convertBtn.setPreferredSize(new Dimension(280, 50));
        convertBtn.addActionListener(e -> convert());

        JPanel footerBar = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        footerBar.setBackground(new Color(45, 45, 48));
        footerBar.setBorder(new EmptyBorder(10, 20, 10, 20));
        footerBar.add(convertBtn);

        main.add(headerBar, BorderLayout.NORTH);
        main.add(mainSplit, BorderLayout.CENTER);
        main.add(footerBar, BorderLayout.SOUTH);

        add(main);
    }

    private void chooseFile() {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            selectedFile = chooser.getSelectedFile();
            fileField.setText(selectedFile.getAbsolutePath());
            loadSheets();
        }
    }

    private void loadSheets() {
        try (FileInputStream fis = new FileInputStream(selectedFile)) {
            List<String> names = excelToJasperService.getSheetNames(fis);
            sheetList.setListData(new Vector<>(names));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void loadData() {
        String sheetName = sheetList.getSelectedValue();
        if (sheetName == null) return;

        try (FileInputStream fis = new FileInputStream(selectedFile);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheet(sheetName);
            Row headerRow = sheet.getRow(0);
            int colCount = headerRow.getPhysicalNumberOfCells();

            // Preview 20 rows
            Vector<String> columns = new Vector<>();
            for (int i = 0; i < colCount; i++) columns.add(headerRow.getCell(i).toString());

            Vector<Vector<String>> data = new Vector<>();
            for (int r = 1; r <= 20; r++) {
                Row row = sheet.getRow(r);
                if (row == null) break;
                Vector<String> rowData = new Vector<>();
                for (int c = 0; c < colCount; c++) {
                    Cell cell = row.getCell(c);
                    rowData.add(cell == null ? "" : cell.toString());
                }
                data.add(rowData);
            }
            previewTable.setModel(new DefaultTableModel(data, columns));

            // Mapping Data
            List<ExcelToJasperService.ColumnMapping> mappings = excelToJasperService.analyzeSheet(new FileInputStream(selectedFile), sheetName);
            String[] mapCols = {"Use", "Original", "Field", "Label", "Param", "Source", "Expression", "Width"};

            DefaultTableModel model = new DefaultTableModel(mapCols, 0) {
                public Class<?> getColumnClass(int c) { return c == 0 ? Boolean.class : (c == 7 ? Integer.class : String.class); }
                public boolean isCellEditable(int r, int c) { return c != 1; }
            };

            for (ExcelToJasperService.ColumnMapping m : mappings) {
                model.addRow(new Object[]{m.use, m.originalName, m.fieldName, m.label, m.paramName, m.source, m.expression, m.width});
            }
            mappingTable.setModel(model);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void convert() {
        try {
            boolean isPro = LicenseManager.isValid(licenseField.getText());
            if (!isPro && usageCount >= FREE_LIMIT) {
                JOptionPane.showMessageDialog(this, "Trial limit reached.", "License Required", JOptionPane.WARNING_MESSAGE);
                return;
            }

            DefaultTableModel model = (DefaultTableModel) mappingTable.getModel();
            List<ExcelToJasperService.ColumnMapping> cols = new ArrayList<>();
            for (int i = 0; i < model.getRowCount(); i++) {
                ExcelToJasperService.ColumnMapping cm = new ExcelToJasperService.ColumnMapping();
                cm.use = (boolean) model.getValueAt(i, 0);
                cm.fieldName = model.getValueAt(i, 2).toString();
                cm.label = model.getValueAt(i, 3).toString();
                cm.paramName = model.getValueAt(i, 4).toString();
                cm.source = model.getValueAt(i, 5).toString();
                cm.expression = model.getValueAt(i, 6).toString();
                cm.width = Integer.parseInt(model.getValueAt(i, 7).toString());
                cols.add(cm);
            }

            String sheetName = sheetList.getSelectedValue();
            String outPath = selectedFile.getAbsolutePath().replace(".xlsx", "_" + normalize(sheetName) + ".jrxml");

            try (FileOutputStream fos = new FileOutputStream(outPath)) {
                excelToJasperService.generateJRXML(cols, fos, headerColor);
            }

            usageCount++;
            if (isPro) {
                statusLabel.setText("Activated (PRO)");
                statusLabel.setForeground(new Color(50, 205, 50));
            } else {
                statusLabel.setText("Trial: " + usageCount + "/" + FREE_LIMIT);
            }

            JOptionPane.showMessageDialog(this, "Generated: " + outPath);

        } catch (Exception e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Error: " + e.getMessage());
        }
    }

    private String normalize(String input) {
        String temp = Normalizer.normalize(input, Normalizer.Form.NFD);
        return temp.replaceAll("[^\\p{ASCII}]", "").replaceAll("[^a-zA-Z0-9]", "_");
    }

    public static void main(String[] args) {
        FlatDarkLaf.setup();
        SwingUtilities.invokeLater(() -> new ExcelToJasperFullApp().setVisible(true));
    }
}
