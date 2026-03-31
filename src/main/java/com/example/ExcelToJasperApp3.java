package com.example;

import net.sf.jasperreports.engine.JRLineBox;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.*;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ExcelToJasperApp3 {

    public static void main(String[] args) throws Exception {

        // ===== CHỌN FILE =====
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Chọn file Excel");

        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File file = chooser.getSelectedFile();
        String excelPath = file.getAbsolutePath();

        // ===== LOAD WORKBOOK =====
        Workbook wb = new XSSFWorkbook(new FileInputStream(excelPath));

        List<String> sheetNames = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            sheetNames.add(wb.getSheetName(i));
        }

        // ===== CHỌN SHEET =====
        String selectedSheet = (String) JOptionPane.showInputDialog(
                null,
                "Chọn sheet:",
                "Select Sheet",
                JOptionPane.PLAIN_MESSAGE,
                null,
                sheetNames.toArray(),
                sheetNames.get(0)
        );

        if (selectedSheet == null) {
            wb.close();
            return;
        }

        wb.close();

        String jrxmlPath = excelPath.replace(".xlsx", "_" + selectedSheet + ".jrxml");

        int headerStartRow = 0;
        int headerRowCount = 2;

        convert(excelPath, selectedSheet, jrxmlPath, headerStartRow, headerRowCount);

        System.out.println("DONE: " + jrxmlPath);
    }

    public static void convert(String excelPath,
                               String sheetName,
                               String jrxmlPath,
                               int headerStartRow,
                               int headerRowCount) throws Exception {

        Workbook wb = new XSSFWorkbook(new FileInputStream(excelPath));
        Sheet sheet = wb.getSheet(sheetName);

        Row lastHeader = sheet.getRow(headerStartRow + headerRowCount - 1);
        int colCount = lastHeader.getPhysicalNumberOfCells();

        // ======================
        // WIDTH FROM EXCEL
        // ======================
        List<Integer> colWidths = new ArrayList<>();
        int totalWidth = 0;

        for (int c = 0; c < colCount; c++) {
            int w = sheet.getColumnWidth(c);
            int px = (int) (w / 256.0 * 7);
            if (px < 30) px = 30;

            colWidths.add(px);
            totalWidth += px;
        }

        // ======================
        // DESIGN
        // ======================
        JasperDesign design = new JasperDesign();
        design.setName("PRO_REPORT");

        int margin = 20;

        design.setLeftMargin(margin);
        design.setRightMargin(margin);
        design.setTopMargin(20);
        design.setBottomMargin(20);

        // 🔥 AUTO EXPAND PAGE
        design.setColumnWidth(totalWidth);
        design.setPageWidth(totalWidth + margin * 2);
        design.setPageHeight(842);

        // ======================
        // FIELDS
        // ======================
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

        // ======================
        // MERGE MAP
        // ======================
        Map<String, CellRangeAddress> mergeMap = new HashMap<>();

        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            mergeMap.put(r.getFirstRow() + "_" + r.getFirstColumn(), r);
        }

        // ======================
        // HEADER
        // ======================
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

        // ======================
        // DETAIL
        // ======================
        JRDesignBand detail = new JRDesignBand();
        detail.setHeight(25);

        int x = 0;

        for (int i = 0; i < fields.size(); i++) {

            JRDesignTextField tf = new JRDesignTextField();
            tf.setX(x);
            tf.setY(0);
            tf.setWidth(colWidths.get(i));
            tf.setHeight(25);

            tf.setHorizontalTextAlign(HorizontalTextAlignEnum.LEFT);

            JRDesignExpression ex = new JRDesignExpression();
            ex.setText("$F{" + fields.get(i) + "}");
            tf.setExpression(ex);

            applyStyle(tf, null);

            detail.addElement(tf);

            x += colWidths.get(i);
        }

        ((JRDesignSection) design.getDetailSection()).addBand(detail);

        // ======================
        // EXPORT
        // ======================
        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        wb.close();
    }

    // ======================
    // UTIL
    // ======================

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

    // 🔥 FIX getLineBox lỗi ở đây
    private static void applyStyle(JRDesignTextElement element, Cell cell) {

        // border chuẩn
        JRLineBox box = element.getLineBox();

        box.getTopPen().setLineWidth(0.5f);
        box.getBottomPen().setLineWidth(0.5f);
        box.getLeftPen().setLineWidth(0.5f);
        box.getRightPen().setLineWidth(0.5f);

        // background đơn giản
        if (cell != null) {
            CellStyle style = cell.getCellStyle();
            if (style.getFillForegroundColor() != 0) {
                element.setMode(ModeEnum.OPAQUE);
            }
        }
    }
}