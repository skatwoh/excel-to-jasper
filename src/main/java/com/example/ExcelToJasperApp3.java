package com.example;

import net.sf.jasperreports.engine.JRLineBox;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.*;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;

public class ExcelToJasperApp3 {

    private JFrame frame;
    private JList<String> sheetList;
    private JTable previewTable;
    private Workbook workbook;
    private String excelPath;

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ExcelToJasperApp3().initUI());
    }

    private void initUI() {
        frame = new JFrame("Excel → Jasper Tool");
        frame.setSize(1000, 600);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new BorderLayout());

        // ===== TOP PANEL =====
        JPanel topPanel = new JPanel();

        JButton btnLoad = new JButton("Chọn Excel");
        JButton btnPreview = new JButton("Preview");
        JButton btnGenerate = new JButton("Generate JRXML");

        topPanel.add(btnLoad);
        topPanel.add(btnPreview);
        topPanel.add(btnGenerate);

        frame.add(topPanel, BorderLayout.NORTH);

        // ===== LEFT: SHEET LIST =====
        sheetList = new JList<>();
        sheetList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        frame.add(new JScrollPane(sheetList), BorderLayout.WEST);

        // ===== CENTER: PREVIEW =====
        previewTable = new JTable();
        frame.add(new JScrollPane(previewTable), BorderLayout.CENTER);

        // ===== ACTIONS =====

        // LOAD FILE
        btnLoad.addActionListener(e -> loadExcel());

        // PREVIEW
        btnPreview.addActionListener(e -> previewSheet());

        // GENERATE
        btnGenerate.addActionListener(e -> generateSelected());

        frame.setVisible(true);
    }

    private void loadExcel() {
        try {
            JFileChooser chooser = new JFileChooser();
            if (chooser.showOpenDialog(frame) != JFileChooser.APPROVE_OPTION) return;

            File file = chooser.getSelectedFile();
            excelPath = file.getAbsolutePath();

            workbook = new XSSFWorkbook(new FileInputStream(file));

            List<String> sheets = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                sheets.add(workbook.getSheetName(i));
            }

            sheetList.setListData(sheets.toArray(new String[0]));

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void previewSheet() {
        try {
            String sheetName = sheetList.getSelectedValue();
            if (sheetName == null) return;

            Sheet sheet = workbook.getSheet(sheetName);

            DefaultTableModel model = new DefaultTableModel();

            int maxCols = 0;

            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    maxCols = Math.max(maxCols, row.getLastCellNum());
                }
            }

            // header
            for (int c = 0; c < maxCols; c++) {
                model.addColumn("C" + c);
            }

            // data
            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                Object[] data = new Object[maxCols];

                for (int c = 0; c < maxCols; c++) {
                    if (row != null && row.getCell(c) != null) {
                        data[c] = row.getCell(c).toString();
                    } else {
                        data[c] = "";
                    }
                }
                model.addRow(data);
            }

            previewTable.setModel(model);

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    private void generateSelected() {
        try {
            List<String> selectedSheets = sheetList.getSelectedValuesList();
            if (selectedSheets.isEmpty()) {
                JOptionPane.showMessageDialog(frame, "Chọn ít nhất 1 sheet");
                return;
            }

            for (String sheet : selectedSheets) {
                String jrxmlPath = excelPath.replace(".xlsx", "_" + sheet + ".jrxml");
                convert(excelPath, sheet, jrxmlPath, 0, 1);
                System.out.println("DONE: " + jrxmlPath);
            }

            JOptionPane.showMessageDialog(frame, "Generate xong!");

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    // ================= CONVERT =================

    public static void convert(String excelPath,
                               String sheetName,
                               String jrxmlPath,
                               int headerStartRow,
                               int headerRowCount) throws Exception {

        Workbook wb = new XSSFWorkbook(new FileInputStream(excelPath));
        Sheet sheet = wb.getSheet(sheetName);

        Row lastHeader = sheet.getRow(headerStartRow + headerRowCount - 1);
        int colCount = lastHeader.getPhysicalNumberOfCells();

        List<Integer> colWidths = new ArrayList<>();
        int totalWidth = 0;

        for (int c = 0; c < colCount; c++) {
            int w = sheet.getColumnWidth(c);
            int px = (int) (w / 256.0 * 7);
            if (px < 30) px = 30;

            colWidths.add(px);
            totalWidth += px;
        }

        JasperDesign design = new JasperDesign();
        design.setName(sheetName);

        int margin = 20;

        design.setLeftMargin(margin);
        design.setRightMargin(margin);
        design.setTopMargin(20);
        design.setBottomMargin(20);

        design.setColumnWidth(totalWidth);
        design.setPageWidth(totalWidth + margin * 2);
        design.setPageHeight(842);

        List<String> fields = new ArrayList<>();
        Set<String> used = new HashSet<>();

        for (int i = 0; i < colCount; i++) {
            String name = lastHeader.getCell(i) == null
                    ? "COL_" + i
                    : lastHeader.getCell(i).toString().replace(" ", "_");

            String base = name;
            int count = 1;

            while (used.contains(name)) {
                name = base + "_" + count++;
            }

            used.add(name);
            fields.add(name);

            JRDesignField f = new JRDesignField();
            f.setName(name);
            f.setValueClass(String.class);
            design.addField(f);
        }

        Map<String, CellRangeAddress> mergeMap = new HashMap<>();

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            mergeMap.put(r.getFirstRow() + "_" + r.getFirstColumn(), r);
        }

        JRDesignBand header = new JRDesignBand();
        header.setHeight(30 * headerRowCount);

        for (int r = 0; r < headerRowCount; r++) {
            Row row = sheet.getRow(headerStartRow + r);
            int x = 0;

            for (int c = 0; c < colCount; c++) {

                if (isMergedButNotFirst(sheet, headerStartRow + r, c)) {
                    x += colWidths.get(c);
                    continue;
                }

                String key = (headerStartRow + r) + "_" + c;
                CellRangeAddress region = mergeMap.get(key);

                int width = colWidths.get(c);
                int height = 30;

                if (region != null) {
                    width = 0;
                    for (int i = region.getFirstColumn(); i <= region.getLastColumn(); i++) {
                        width += colWidths.get(i);
                    }
                    height = (region.getLastRow() - region.getFirstRow() + 1) * 30;
                }

                String text = getCellValue(row.getCell(c));

                JRDesignStaticText st = new JRDesignStaticText();
                st.setX(x);
                st.setY(r * 30);
                st.setWidth(width);
                st.setHeight(height);
                st.setText(text);
                st.setBold(true);
                st.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
                st.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);

                applyStyle(st, row.getCell(c));

                header.addElement(st);

                x += colWidths.get(c);
            }
        }

        design.setColumnHeader(header);

        JRDesignBand detail = new JRDesignBand();
        detail.setHeight(25);

        int x = 0;

        for (int i = 0; i < fields.size(); i++) {

            JRDesignTextField tf = new JRDesignTextField();
            tf.setX(x);
            tf.setY(0);
            tf.setWidth(colWidths.get(i));
            tf.setHeight(25);

            JRDesignExpression ex = new JRDesignExpression();
            ex.setText("$F{" + fields.get(i) + "}");
            tf.setExpression(ex);

            applyStyle(tf, null);

            detail.addElement(tf);

            x += colWidths.get(i);
        }

        ((JRDesignSection) design.getDetailSection()).addBand(detail);

        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        wb.close();
    }

    private static boolean isMergedButNotFirst(Sheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            if (r.isInRange(row, col)) {
                return !(r.getFirstRow() == row && r.getFirstColumn() == col);
            }
        }
        return false;
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return cell.toString();
    }

    private static void applyStyle(JRDesignTextElement element, Cell cell) {
        JRLineBox box = element.getLineBox();

        box.getTopPen().setLineWidth(0.5f);
        box.getBottomPen().setLineWidth(0.5f);
        box.getLeftPen().setLineWidth(0.5f);
        box.getRightPen().setLineWidth(0.5f);

        if (cell != null) {
            CellStyle style = cell.getCellStyle();
            if (style.getFillForegroundColor() != 0) {
                element.setMode(ModeEnum.OPAQUE);
            }
        }
    }
}