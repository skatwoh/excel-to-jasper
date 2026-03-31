package com.example;

import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.type.ModeEnum;
import net.sf.jasperreports.engine.type.LineStyleEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.Color;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;

public class ExcelToJasperApp2 extends JFrame {

    // =====================================================================
    // CONSTANTS — chỉnh 1 chỗ, áp dụng toàn bộ JRXML output
    // =====================================================================
    private static final int PAGE_WIDTH        = 842; // A4 landscape (pt)
    private static final int PAGE_HEIGHT       = 595;
    private static final int MARGIN            = 20;
    private static final int USABLE_WIDTH      = PAGE_WIDTH - MARGIN * 2; // 802
    private static final int HEADER_ROW_HEIGHT = 25;
    private static final int DETAIL_ROW_HEIGHT = 20;
    private static final int FONT_SIZE_HEADER  = 10;
    private static final int FONT_SIZE_DETAIL  = 9;
    private static final int COL_WIDTH_STEP    = 10; // làm tròn lên bội số này

    // =====================================================================
    // UI fields
    // =====================================================================
    private JTextField fileField;
    private JTextField headerStartField;
    private JTextField headerCountField;

    private DefaultListModel<String> sheetListModel = new DefaultListModel<>();
    private JList<String> sheetList;
    private JTabbedPane previewTabs;

    // =====================================================================
    // CONSTRUCTOR
    // =====================================================================
    public ExcelToJasperApp2() {
        setTitle("Excel → Jasper Converter (Multi-Sheet)");
        setSize(1100, 700);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // ----- TOP PANEL -----
        JPanel topPanel = new JPanel(new GridBagLayout());
        topPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 5, 10));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        gbc.gridx = 0; gbc.gridy = 0; gbc.weightx = 0;
        topPanel.add(new JLabel("Excel File:"), gbc);
        gbc.gridx = 1; gbc.weightx = 1;
        fileField = new JTextField();
        topPanel.add(fileField, gbc);
        gbc.gridx = 2; gbc.weightx = 0;
        JButton browseBtn = new JButton("Browse");
        browseBtn.addActionListener(this::browseFile);
        topPanel.add(browseBtn, gbc);
        gbc.gridx = 3;
        JButton loadSheetsBtn = new JButton("Load Sheets");
        loadSheetsBtn.addActionListener(this::loadSheets);
        topPanel.add(loadSheetsBtn, gbc);

        gbc.gridx = 0; gbc.gridy = 1; gbc.weightx = 0;
        topPanel.add(new JLabel("Header Start Row (0-based):"), gbc);
        gbc.gridx = 1; gbc.weightx = 1;
        headerStartField = new JTextField("0");
        topPanel.add(headerStartField, gbc);

        gbc.gridx = 0; gbc.gridy = 2; gbc.weightx = 0;
        topPanel.add(new JLabel("Header Row Count:"), gbc);
        gbc.gridx = 1; gbc.weightx = 1;
        headerCountField = new JTextField("2");
        topPanel.add(headerCountField, gbc);

        add(topPanel, BorderLayout.NORTH);

        // ----- CENTER: Sheet selector + Preview tabs -----
        JSplitPane splitPane = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
        splitPane.setDividerLocation(200);

        JPanel sheetPanel = new JPanel(new BorderLayout());
        sheetPanel.setBorder(BorderFactory.createTitledBorder("Sheets (Ctrl+Click để chọn nhiều)"));
        sheetList = new JList<>(sheetListModel);
        sheetList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        sheetPanel.add(new JScrollPane(sheetList), BorderLayout.CENTER);
        JButton selectAllBtn = new JButton("Chọn tất cả");
        selectAllBtn.addActionListener(e -> {
            if (sheetListModel.size() > 0)
                sheetList.setSelectionInterval(0, sheetListModel.size() - 1);
        });
        sheetPanel.add(selectAllBtn, BorderLayout.SOUTH);
        splitPane.setLeftComponent(sheetPanel);

        previewTabs = new JTabbedPane();
        splitPane.setRightComponent(previewTabs);
        add(splitPane, BorderLayout.CENTER);

        // ----- BOTTOM BUTTONS -----
        JPanel bottomPanel = new JPanel();
        JButton previewBtn = new JButton("Preview (các sheet đã chọn)");
        previewBtn.addActionListener(this::previewSelectedSheets);
        JButton generateBtn = new JButton("Generate JRXML (các sheet đã chọn)");
        generateBtn.addActionListener(this::generateReport);
        bottomPanel.add(previewBtn);
        bottomPanel.add(generateBtn);
        add(bottomPanel, BorderLayout.SOUTH);
    }

    // =====================================================================
    // BROWSE FILE
    // =====================================================================
    private void browseFile(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            fileField.setText(chooser.getSelectedFile().getAbsolutePath());
            loadSheets(e);
        }
    }

    // =====================================================================
    // LOAD SHEETS
    // =====================================================================
    private void loadSheets(ActionEvent e) {
        String path = fileField.getText().trim();
        if (path.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn file Excel trước.");
            return;
        }
        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            sheetListModel.clear();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++)
                sheetListModel.addElement(workbook.getSheetName(i));
            if (sheetListModel.size() > 0)
                sheetList.setSelectionInterval(0, sheetListModel.size() - 1);
            previewTabs.removeAll();

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Lỗi đọc sheet: " + ex.getMessage());
        }
    }

    // =====================================================================
    // PREVIEW
    // =====================================================================
    private void previewSelectedSheets(ActionEvent e) {
        List<String> selected = sheetList.getSelectedValuesList();
        if (selected.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn ít nhất một sheet.");
            return;
        }
        String path = fileField.getText().trim();
        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            previewTabs.removeAll();
            for (String sheetName : selected) {
                Sheet sheet = workbook.getSheet(sheetName);
                if (sheet == null) continue;
                previewTabs.addTab(sheetName, new JScrollPane(buildTableFromSheet(sheet)));
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Preview Error: " + ex.getMessage());
        }
    }

    private JTable buildTableFromSheet(Sheet sheet) {
        int maxColumns = 0;
        for (Row row : sheet)
            maxColumns = Math.max(maxColumns, row.getPhysicalNumberOfCells());

        DefaultTableModel model = new DefaultTableModel();
        for (int i = 0; i < maxColumns; i++) model.addColumn("Col " + i);
        for (Row row : sheet) {
            Vector<String> rowData = new Vector<>();
            for (int i = 0; i < maxColumns; i++) {
                Cell cell = row.getCell(i);
                rowData.add(cell == null ? "" : cell.toString());
            }
            model.addRow(rowData);
        }
        return new JTable(model);
    }

    // =====================================================================
    // GENERATE JRXML
    // =====================================================================
    private void generateReport(ActionEvent e) {
        List<String> selected = sheetList.getSelectedValuesList();
        if (selected.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn ít nhất một sheet.");
            return;
        }
        String excelPath = fileField.getText().trim();
        if (excelPath.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Vui lòng chọn file Excel.");
            return;
        }

        int headerStart, headerCount;
        try {
            headerStart = Integer.parseInt(headerStartField.getText().trim());
            headerCount = Integer.parseInt(headerCountField.getText().trim());
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "Header row phải là số nguyên.");
            return;
        }

        File excelFile  = new File(excelPath);
        String baseName = excelFile.getName().replaceAll("\\.[^.]+$", "");
        String dir      = excelFile.getParent();

        StringBuilder sb = new StringBuilder("Đã tạo:\n");
        List<String> errors = new ArrayList<>();

        for (String sheetName : selected) {
            String safeSheet  = sheetName.replaceAll("[\\\\/:*?\"<>|]", "_");
            String outputPath = dir + File.separator + baseName + "_" + safeSheet + ".jrxml";
            try {
                convert(excelPath, outputPath, sheetName, headerStart, headerCount);
                sb.append("  • ").append(outputPath).append("\n");
            } catch (Exception ex) {
                errors.add("Sheet [" + sheetName + "]: " + ex.getMessage());
            }
        }
        if (!errors.isEmpty()) {
            sb.append("\nLỗi:\n");
            for (String err : errors) sb.append("  ✗ ").append(err).append("\n");
        }
        JOptionPane.showMessageDialog(this, sb.toString());
    }

    // =====================================================================
    // CONVERT LOGIC
    // =====================================================================
    public static void convert(String excelPath,
                               String jrxmlPath,
                               String sheetName,
                               int headerStartRow,
                               int headerRowCount) throws Exception {

        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null)
                throw new IllegalArgumentException("Sheet không tồn tại: " + sheetName);

            Row lastHeaderRow = sheet.getRow(headerStartRow + headerRowCount - 1);
            if (lastHeaderRow == null)
                throw new IllegalArgumentException(
                        "Sheet [" + sheetName + "]: không tìm thấy header row tại index "
                                + (headerStartRow + headerRowCount - 1));

            int columnCount = lastHeaderRow.getPhysicalNumberOfCells();

            // columnWidth tự động, làm tròn lên bội số COL_WIDTH_STEP
            int rawWidth    = columnCount > 0 ? USABLE_WIDTH / columnCount : 100;
            int columnWidth = ((rawWidth + COL_WIDTH_STEP - 1) / COL_WIDTH_STEP) * COL_WIDTH_STEP;

            // ------------------------------------------------------------------
            // JASPER DESIGN — A4 landscape
            // ------------------------------------------------------------------
            JasperDesign design = new JasperDesign();
            design.setName(sheetName.replaceAll("[^a-zA-Z0-9_]", "_"));
            design.setPageWidth(PAGE_WIDTH);
            design.setPageHeight(PAGE_HEIGHT);
            design.setColumnWidth(USABLE_WIDTH);
            design.setLeftMargin(MARGIN);
            design.setRightMargin(MARGIN);
            design.setTopMargin(MARGIN);
            design.setBottomMargin(MARGIN);

            // ------------------------------------------------------------------
            // FIELDS (dùng dòng header cuối cùng)
            // ------------------------------------------------------------------
            List<String> fieldNames = new ArrayList<>();
            Set<String> used = new HashSet<>();

            for (int i = 0; i < columnCount; i++) {
                String raw = lastHeaderRow.getCell(i) == null
                        ? ""
                        : lastHeaderRow.getCell(i).toString().trim();

                String fieldName = raw.isEmpty()
                        ? "COLUMN_" + i
                        : raw.replaceAll("[^a-zA-Z0-9_]", "_");

                if (fieldName.matches("^[0-9].*")) fieldName = "F_" + fieldName;

                String original = fieldName;
                int counter = 1;
                while (used.contains(fieldName)) fieldName = original + "_" + counter++;
                used.add(fieldName);
                fieldNames.add(fieldName);

                JRDesignField field = new JRDesignField();
                field.setName(fieldName);
                field.setValueClass(String.class);
                design.addField(field);
            }

            // ------------------------------------------------------------------
            // COLUMN HEADER BAND — Bold + font size + CENTER + border + background
            // ------------------------------------------------------------------
            JRDesignBand headerBand = new JRDesignBand();
            headerBand.setHeight(HEADER_ROW_HEIGHT * headerRowCount);

            for (int h = 0; h < headerRowCount; h++) {
                Row headerRow = sheet.getRow(headerStartRow + h);
                int x = 0;
                for (int c = 0; c < columnCount; c++) {
                    String text = (headerRow == null || headerRow.getCell(c) == null)
                            ? ""
                            : headerRow.getCell(c).toString();

                    JRDesignStaticText st = new JRDesignStaticText();
                    st.setX(x);
                    st.setY(h * HEADER_ROW_HEIGHT);
                    st.setWidth(columnWidth);
                    st.setHeight(HEADER_ROW_HEIGHT);
                    st.setText(text);

                    st.setBold(true);
                    st.setFontSize((float) FONT_SIZE_HEADER);
                    st.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

                    // Background xám nhạt cho header
                    st.setMode(ModeEnum.OPAQUE);
                    st.setBackcolor(new Color(220, 220, 220));

                    // Border đầy đủ 4 cạnh
                    applyBorder(st.getLineBox(), 0.5f, LineStyleEnum.SOLID);

                    headerBand.addElement(st);
                    x += columnWidth;
                }
            }
            design.setColumnHeader(headerBand);

            // ------------------------------------------------------------------
            // DETAIL BAND — CENTER align + font size + border
            // ------------------------------------------------------------------
            JRDesignBand detailBand = new JRDesignBand();
            detailBand.setHeight(DETAIL_ROW_HEIGHT);
            int x = 0;

            for (String fieldName : fieldNames) {
                JRDesignTextField tf = new JRDesignTextField();
                tf.setX(x);
                tf.setY(0);
                tf.setWidth(columnWidth);
                tf.setHeight(DETAIL_ROW_HEIGHT);

                tf.setFontSize((float) FONT_SIZE_DETAIL);
                tf.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

                // Border đầy đủ 4 cạnh
                applyBorder(tf.getLineBox(), 0.5f, LineStyleEnum.SOLID);

                JRDesignExpression expr = new JRDesignExpression();
                expr.setText("$F{" + fieldName + "}");
                tf.setExpression(expr);

                detailBand.addElement(tf);
                x += columnWidth;
            }

            ((JRDesignSection) design.getDetailSection()).addBand(detailBand);

            JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");
        }
    }

    // =====================================================================
    // HELPER: áp dụng border cho JRLineBox
    // =====================================================================
    private static void applyBorder(net.sf.jasperreports.engine.JRLineBox box,
                                    float width,
                                    LineStyleEnum style) {
        box.getTopPen().setLineWidth(width);
        box.getTopPen().setLineStyle(style);
        box.getBottomPen().setLineWidth(width);
        box.getBottomPen().setLineStyle(style);
        box.getLeftPen().setLineWidth(width);
        box.getLeftPen().setLineStyle(style);
        box.getRightPen().setLineWidth(width);
        box.getRightPen().setLineStyle(style);

        // padding nhỏ để chữ không dính sát border
        box.setTopPadding(2);
        box.setBottomPadding(2);
        box.setLeftPadding(3);
        box.setRightPadding(3);
    }

    // =====================================================================
    // Backward-compatible overload (giữ lại API cũ của bản gốc)
    // =====================================================================
    public static void convert(String excelPath,
                               String jrxmlPath,
                               int headerStartRow,
                               int headerRowCount) throws Exception {
        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            convert(excelPath, jrxmlPath, workbook.getSheetName(0), headerStartRow, headerRowCount);
        }
    }

    // =====================================================================
    // MAIN
    // =====================================================================
    public static void main(String[] args) {
        // Giữ lại cách dùng cũ nếu chạy từ command line với args
        if (args.length >= 2) {
            try {
                String excelPath  = args[0];
                String jrxmlPath  = args[1];
                int headerStart   = args.length > 2 ? Integer.parseInt(args[2]) : 0;
                int headerCount   = args.length > 3 ? Integer.parseInt(args[3]) : 2;
                convert(excelPath, jrxmlPath, headerStart, headerCount);
                System.out.println("DONE! Generated: " + jrxmlPath);
            } catch (Exception e) {
                System.err.println("Error: " + e.getMessage());
            }
        } else {
            // Khởi động Swing UI
            SwingUtilities.invokeLater(() -> new ExcelToJasperApp2().setVisible(true));
        }
    }
}