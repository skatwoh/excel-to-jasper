package com.example;

import net.sf.jasperreports.components.table.DesignCell;
import net.sf.jasperreports.components.table.StandardColumn;
import net.sf.jasperreports.components.table.StandardTable;
import net.sf.jasperreports.engine.component.ComponentKey;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class ExcelToJasperTableApp {

    public static void main(String[] args) throws Exception {

        // ===== CHỌN FILE =====
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Chọn file Excel");

        if (chooser.showOpenDialog(null) != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File file = chooser.getSelectedFile();
        String excelPath = file.getAbsolutePath();

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

        String jrxmlPath = excelPath.replace(".xlsx", "_" + selectedSheet + "_table.jrxml");

        convert(excelPath, selectedSheet, jrxmlPath);

        System.out.println("DONE: " + jrxmlPath);
    }

    public static void convert(String excelPath,
                               String sheetName,
                               String jrxmlPath) throws Exception {

        Workbook wb = new XSSFWorkbook(new FileInputStream(excelPath));
        Sheet sheet = wb.getSheet(sheetName);

        Row headerRow = sheet.getRow(0);
        int colCount = headerRow.getPhysicalNumberOfCells();

        // ======================
        // WIDTH
        // ======================
        List<Integer> colWidths = new ArrayList<>();
        int totalWidth = 0;

        for (int c = 0; c < colCount; c++) {
            int w = sheet.getColumnWidth(c);
            int px = (int) (w / 256.0 * 7);
            if (px < 50) px = 50;

            colWidths.add(px);
            totalWidth += px;
        }

        // ======================
        // DESIGN
        // ======================
        JasperDesign design = new JasperDesign();
        design.setName("TABLE_REPORT");

        int margin = 20;

        design.setLeftMargin(margin);
        design.setRightMargin(margin);
        design.setTopMargin(20);
        design.setBottomMargin(20);

        design.setColumnWidth(totalWidth);
        design.setPageWidth(totalWidth + margin * 2);
        design.setPageHeight(842);

        // ======================
        // FIELDS
        // ======================
        List<String> fields = new ArrayList<>();
        Set<String> used = new HashSet<>();

        for (int i = 0; i < colCount; i++) {

            String name = headerRow.getCell(i) == null
                    ? "COL_" + i
                    : headerRow.getCell(i).toString().replace(" ", "_");

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
        // DATASET CHO TABLE
        // ======================
        JRDesignDataset dataset = new JRDesignDataset(false);
        dataset.setName("tableDataset");

        for (String f : fields) {
            JRDesignField field = new JRDesignField();
            field.setName(f);
            field.setValueClass(String.class);
            dataset.addField(field);
        }

        design.addDataset(dataset);

        // ======================
        // TABLE
        // ======================
        StandardTable table = new StandardTable();

        JRDesignDatasetRun datasetRun = new JRDesignDatasetRun();
        datasetRun.setDatasetName("tableDataset");
        datasetRun.setDataSourceExpression(
                new JRDesignExpression("$P{REPORT_DATA_SOURCE}")
        );

        table.setDatasetRun(datasetRun);

        // ======================
        // COLUMNS
        // ======================
        for (int i = 0; i < fields.size(); i++) {

            StandardColumn column = new StandardColumn();
            column.setWidth(colWidths.get(i));

            // HEADER
            DesignCell headerCell = new DesignCell();
            headerCell.setHeight(30);

            JRDesignStaticText headerText = new JRDesignStaticText();
            headerText.setWidth(colWidths.get(i));
            headerText.setHeight(30);
            headerText.setText(fields.get(i));
            headerText.setBold(true);
            headerText.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

            headerCell.addElement(headerText);
            column.setColumnHeader(headerCell);

            // DETAIL
            DesignCell detailCell = new DesignCell();
            detailCell.setHeight(25);

            JRDesignTextField tf = new JRDesignTextField();
            tf.setWidth(colWidths.get(i));
            tf.setHeight(25);

            JRDesignExpression exp = new JRDesignExpression();
            exp.setText("$F{" + fields.get(i) + "}");
            tf.setExpression(exp);

            detailCell.addElement(tf);
            column.setDetailCell(detailCell);

            table.addColumn(column);
        }

        // ======================
        // ADD TABLE TO REPORT
        // ======================
        JRDesignComponentElement componentElement = new JRDesignComponentElement();
        componentElement.setX(0);
        componentElement.setY(0);
        componentElement.setWidth(totalWidth);
        componentElement.setHeight(40);

        componentElement.setComponentKey(
                new ComponentKey("http://jasperreports.sourceforge.net/jasperreports/components",
                        "jr", "table")
        );

        componentElement.setComponent(table);

        JRDesignBand detailBand = new JRDesignBand();
        detailBand.setHeight(50);
        detailBand.addElement(componentElement);

        ((JRDesignSection) design.getDetailSection()).addBand(detailBand);

        // ======================
        // EXPORT
        // ======================
        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        wb.close();
    }
}