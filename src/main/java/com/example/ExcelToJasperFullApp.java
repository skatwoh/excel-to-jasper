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
    private static final String VALID_LICENSE = "PRO-2024-EXCEL-JASPER";

    private Color headerColor = Color.LIGHT_GRAY; // 🎨 màu header

    public ExcelToJasperFullApp() {
        setTitle("Excel → Jasper PRO UI");
        setSize(1300, 800);
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
        try (Workbook wb = new XSSFWorkbook(new FileInputStream(selectedFile))) {
            Vector<String> names = new Vector<>();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                names.add(wb.getSheetName(i));
            }
            sheetList.setListData(names);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void loadData() {
        String sheetName = sheetList.getSelectedValue();
        if (sheetName == null) return;

        try (Workbook wb = new XSSFWorkbook(new FileInputStream(selectedFile))) {
            Sheet sheet = wb.getSheet(sheetName);
            Row header = sheet.getRow(0);

            int colCount = header.getPhysicalNumberOfCells();

            Vector<String> columns = new Vector<>();
            Vector<Vector<String>> data = new Vector<>();

            for (int i = 0; i < colCount; i++) {
                columns.add(header.getCell(i).toString());
            }

            for (int r = 1; r <= 20; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                Vector<String> rowData = new Vector<>();
                for (int c = 0; c < colCount; c++) {
                    Cell cell = row.getCell(c);
                    rowData.add(cell == null ? "" : cell.toString());
                }
                data.add(rowData);
            }

            previewTable.setModel(new DefaultTableModel(data, columns));

            Vector<String> mapCols = new Vector<>(Arrays.asList(
                    "Use", "Original", "Field", "Label", "Param", "Source", "Expression"
            ));

            Vector<Vector<Object>> mapData = new Vector<>();

            for (String col : columns) {
                String clean = col.replace(" ", "_");

                Vector<Object> row = new Vector<>();
                row.add(true);
                row.add(col);
                row.add(clean);
                row.add(col);
                row.add("P_" + clean);
                row.add("FIELD");
                row.add("$F{" + clean + "}");

                mapData.add(row);
            }

            DefaultTableModel model = new DefaultTableModel(mapData, mapCols) {
                public Class<?> getColumnClass(int c) {
                    return c == 0 ? Boolean.class : String.class;
                }

                public boolean isCellEditable(int r, int c) {
                    return c != 1;
                }
            };

            mappingTable.setModel(model);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void convert() {
        try {
            boolean isPro = licenseField.getText().trim().equals(VALID_LICENSE);

            if (!isPro && usageCount >= FREE_LIMIT) {
                JOptionPane.showMessageDialog(this,
                    "Bạn đã hết lượt dùng thử (Giới hạn: " + FREE_LIMIT + ").\nVui lòng nhập License để tiếp tục!",
                    "License Required", JOptionPane.WARNING_MESSAGE);
                return;
            }

            DefaultTableModel model = (DefaultTableModel) mappingTable.getModel();

            List<Map<String, String>> cols = new ArrayList<>();

            for (int i = 0; i < model.getRowCount(); i++) {
                boolean use = (boolean) model.getValueAt(i, 0);
                if (!use) continue;

                String field = model.getValueAt(i, 2).toString();
                String label = model.getValueAt(i, 3).toString();
                String param = cleanParam(model.getValueAt(i, 4).toString());
                String source = model.getValueAt(i, 5).toString();
                String exp = model.getValueAt(i, 6).toString();

                if (exp == null || exp.trim().isEmpty()) {
                    exp = source.equals("PARAM")
                            ? "$P{" + param + "}"
                            : "$F{" + field + "}";
                }

                Map<String, String> m = new HashMap<>();
                m.put("field", field);
                m.put("label", label);
                m.put("param", param);
                m.put("exp", exp);

                cols.add(m);
            }

            String sheetName = sheetList.getSelectedValue();
            String base = selectedFile.getAbsolutePath().replace(".xlsx", "");
            String safeSheetName = normalize(sheetName);

            String out = base + "_" + safeSheetName + ".jrxml";

            buildJRXML(out, cols);

            usageCount++;
            if (isPro) {
                statusLabel.setText("Activated (PRO)");
                statusLabel.setForeground(new Color(50, 205, 50));
            } else {
                statusLabel.setText("Trial: " + usageCount + "/" + FREE_LIMIT);
            }

            JOptionPane.showMessageDialog(this, "DONE: " + out);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String cleanParam(String p) {
        return p.replace("$P{", "").replace("}", "").trim();
    }

    private String normalize(String input) {
        String temp = Normalizer.normalize(input, Normalizer.Form.NFD);
        return temp.replaceAll("[^\\p{ASCII}]", "")
                .replaceAll("[^a-zA-Z0-9]", "_");
    }

    private void buildJRXML(String path, List<Map<String, String>> cols) throws Exception {

        JasperDesign design = new JasperDesign();
        design.setName("FINAL_REPORT");

        int width = cols.size() * 140;
        design.setColumnWidth(width);
        design.setPageWidth(width + 40);
        design.setPageHeight(842);

        // SUB DATASET
        JRDesignDataset dataset = new JRDesignDataset(false);
        dataset.setName("ItemDataSource");

        for (Map<String, String> c : cols) {
            JRDesignField f = new JRDesignField();
            f.setName(c.get("field"));
            f.setValueClass(String.class);
            dataset.addField(f);
        }

        design.addDataset(dataset);

        // PARAM
        JRDesignParameter dsParam = new JRDesignParameter();
        dsParam.setName("ItemDataSource");
        dsParam.setValueClass(JRBeanCollectionDataSource.class);
        design.addParameter(dsParam);

        for (Map<String, String> c : cols) {
            JRDesignParameter p = new JRDesignParameter();
            p.setName(c.get("param"));
            p.setValueClass(String.class);
            design.addParameter(p);
        }

        // TABLE
        StandardTable table = new StandardTable();

        JRDesignDatasetRun run = new JRDesignDatasetRun();
        run.setDatasetName("ItemDataSource");
        run.setDataSourceExpression(new JRDesignExpression("$P{ItemDataSource}"));

        table.setDatasetRun(run);

        for (Map<String, String> c : cols) {

            StandardColumn col = new StandardColumn();
            col.setWidth(140);

            // ===== HEADER =====
            DesignCell header = new DesignCell();
            header.setHeight(30);
            header.getLineBox().getPen().setLineWidth(1f);
            header.getLineBox().getPen().setLineStyle(LineStyleEnum.SOLID);

            JRDesignStaticText txt = new JRDesignStaticText();
            txt.setText(c.get("label"));
            txt.setWidth(140);
            txt.setHeight(30);
            txt.setBold(true);
            txt.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
            txt.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);
            txt.setBackcolor(headerColor);
            txt.setMode(ModeEnum.OPAQUE);

            header.addElement(txt);
            col.setColumnHeader(header);

            // ===== DETAIL =====
            DesignCell detail = new DesignCell();
            detail.setHeight(25);
            detail.getLineBox().getPen().setLineWidth(1f);
            detail.getLineBox().getPen().setLineStyle(LineStyleEnum.SOLID);

            JRDesignTextField tf = new JRDesignTextField();
            tf.setExpression(new JRDesignExpression(c.get("exp")));
            tf.setWidth(140);
            tf.setHeight(25);

            tf.setHorizontalTextAlign(HorizontalTextAlignEnum.LEFT);
            tf.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);
            tf.setStretchWithOverflow(true);

            // zebra row
            JRDesignExpression bgExp = new JRDesignExpression();
            bgExp.setText("$V{REPORT_COUNT} % 2 == 0 ? java.awt.Color.WHITE : new java.awt.Color(240,240,240)");
            tf.setMode(ModeEnum.OPAQUE);
            tf.setBackcolor(Color.WHITE);

            detail.addElement(tf);
            col.setDetailCell(detail);

            table.addColumn(col);
        }

        JRDesignComponentElement comp = new JRDesignComponentElement();
        comp.setComponentKey(new ComponentKey(
                "http://jasperreports.sourceforge.net/jasperreports/components",
                "jr",
                "table"
        ));

        comp.setComponent(table);
        comp.setWidth(width);

        JRDesignBand band = new JRDesignBand();
        band.setHeight(60);
        band.addElement(comp);

        ((JRDesignSection) design.getDetailSection()).addBand(band);

        JRXmlWriter.writeReport(design, path, "UTF-8");
    }

    public static void main(String[] args) {
        FlatDarkLaf.setup();
        SwingUtilities.invokeLater(() -> new ExcelToJasperFullApp().setVisible(true));
    }
}