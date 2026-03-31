package com.example;

import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ExcelToJasperApp2 {

    public static void main(String[] args) throws Exception {

        // ===== CHỌN FILE EXCEL =====
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Chọn file Excel");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Excel Files", "xlsx"));

        int result = fileChooser.showOpenDialog(null);

        if (result != JFileChooser.APPROVE_OPTION) {
            System.out.println("Không chọn file!");
            return;
        }

        File selectedFile = fileChooser.getSelectedFile();
        String excelPath = selectedFile.getAbsolutePath();

        // ===== OUTPUT =====
        String jrxmlPath = excelPath.replace(".xlsx", ".jrxml");

        int headerStartRow = 0;
        int headerRowCount = 2;

        convert(excelPath, jrxmlPath, headerStartRow, headerRowCount);

        System.out.println("DONE! Generated: " + jrxmlPath);
    }

    public static void convert(String excelPath,
                               String jrxmlPath,
                               int headerStartRow,
                               int headerRowCount) throws Exception {

        FileInputStream fis = new FileInputStream(excelPath);
        Workbook workbook = new XSSFWorkbook(fis);
        Sheet sheet = workbook.getSheetAt(0);

        Row lastHeaderRow = sheet.getRow(headerStartRow + headerRowCount - 1);
        int columnCount = lastHeaderRow.getPhysicalNumberOfCells();

        // ======================
        // TÍNH WIDTH TỪ EXCEL
        // ======================
        List<Integer> columnWidths = new ArrayList<>();
        int totalWidth = 0;

        for (int c = 0; c < columnCount; c++) {
            int excelWidth = sheet.getColumnWidth(c); // đơn vị 1/256 char

            // convert sang pixel (gần đúng)
            int px = (int) (excelWidth / 256.0 * 7);

            if (px < 30) px = 30; // tránh quá nhỏ

            columnWidths.add(px);
            totalWidth += px;
        }

        // ======================
        // CREATE DESIGN (AUTO WIDTH)
        // ======================
        JasperDesign design = new JasperDesign();
        design.setName("AutoReport");

        int leftMargin = 20;
        int rightMargin = 20;

        design.setLeftMargin(leftMargin);
        design.setRightMargin(rightMargin);
        design.setTopMargin(20);
        design.setBottomMargin(20);

        // 🔥 FIX QUAN TRỌNG
        design.setColumnWidth(totalWidth);
        design.setPageWidth(totalWidth + leftMargin + rightMargin);

        design.setPageHeight(842); // giữ A4 dọc

        // ======================
        // CREATE FIELDS
        // ======================
        List<String> fieldNames = new ArrayList<>();
        Set<String> usedNames = new HashSet<>();

        for (int i = 0; i < columnCount; i++) {

            String rawName = lastHeaderRow.getCell(i) == null
                    ? ""
                    : lastHeaderRow.getCell(i).toString().trim();

            String fieldName = rawName.isEmpty()
                    ? "COLUMN_" + i
                    : rawName.replace(" ", "_");

            String original = fieldName;
            int counter = 1;

            while (usedNames.contains(fieldName)) {
                fieldName = original + "_" + counter;
                counter++;
            }

            usedNames.add(fieldName);
            fieldNames.add(fieldName);

            JRDesignField field = new JRDesignField();
            field.setName(fieldName);
            field.setValueClass(String.class);

            design.addField(field);
        }

        // ======================
        // HEADER BAND
        // ======================
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
                staticText.setWidth(columnWidths.get(c)); // 🔥 dùng width thật
                staticText.setHeight(25);
                staticText.setText(text);
                staticText.setBold(true);
                staticText.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

                headerBand.addElement(staticText);
                x += columnWidths.get(c);
            }
        }

        design.setColumnHeader(headerBand);

        // ======================
        // DETAIL BAND
        // ======================
        JRDesignBand detailBand = new JRDesignBand();
        detailBand.setHeight(20);

        int xDetail = 0;

        for (int i = 0; i < fieldNames.size(); i++) {

            String fieldName = fieldNames.get(i);

            JRDesignTextField textField = new JRDesignTextField();
            textField.setX(xDetail);
            textField.setY(0);
            textField.setWidth(columnWidths.get(i)); // 🔥 chuẩn Excel
            textField.setHeight(20);
            textField.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

            JRDesignExpression expression = new JRDesignExpression();
            expression.setText("$F{" + fieldName + "}");
            textField.setExpression(expression);

            detailBand.addElement(textField);
            xDetail += columnWidths.get(i);
        }

        ((JRDesignSection) design.getDetailSection()).addBand(detailBand);

        // ======================
        // EXPORT
        // ======================
        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        workbook.close();
        fis.close();
    }
}