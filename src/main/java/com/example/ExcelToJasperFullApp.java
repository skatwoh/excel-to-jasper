package com.example;

import net.sf.jasperreports.components.table.*;
import net.sf.jasperreports.engine.component.ComponentKey;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.*;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
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
    private File selectedFile;

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

        JButton chooseBtn = new JButton("Chọn Excel");
        JButton colorBtn = new JButton("Màu Header");

        fileField = new JTextField();
        fileField.setEditable(false);

        chooseBtn.addActionListener(e -> chooseFile());

        colorBtn.addActionListener(e -> {
            Color chosen = JColorChooser.showDialog(this, "Chọn màu header", headerColor);
            if (chosen != null) headerColor = chosen;
        });

        JPanel top = new JPanel(new BorderLayout());
        top.add(fileField, BorderLayout.CENTER);

        JPanel leftTop = new JPanel();
        leftTop.add(chooseBtn);
        leftTop.add(colorBtn);

        top.add(leftTop, BorderLayout.WEST);

        sheetList = new JList<>();
        sheetList.addListSelectionListener(e -> loadData());

        previewTable = new JTable();
        mappingTable = new JTable();

        JSplitPane right = new JSplitPane(JSplitPane.VERTICAL_SPLIT,
                new JScrollPane(previewTable),
                new JScrollPane(mappingTable));
        right.setDividerLocation(400);

        JSplitPane mainSplit = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(sheetList), right);

        JButton convertBtn = new JButton("Convert JRXML");
        convertBtn.addActionListener(e -> convert());

        main.add(top, BorderLayout.NORTH);
        main.add(mainSplit, BorderLayout.CENTER);
        main.add(convertBtn, BorderLayout.SOUTH);

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
        SwingUtilities.invokeLater(() -> new ExcelToJasperFullApp().setVisible(true));
    }
}