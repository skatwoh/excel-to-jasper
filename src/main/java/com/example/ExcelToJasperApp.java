package com.example;

import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.util.*;

public class ExcelToJasperApp {

    public static void main(String[] args) throws Exception {

        String excelPath = "sample.xlsx";
        String jrxmlPath = "output.jrxml";

        // ===== CHỌN HEADER =====
        int headerStartRow = 0;   // dòng bắt đầu header (0-based)
        int headerRowCount = 2;   // số dòng header

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
        // CREATE JASPER DESIGN
        // ======================
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

        // ======================
        // CREATE FIELDS (DÙNG DÒNG HEADER CUỐI)
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

            // FIX TRÙNG FIELD
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
        // CREATE MULTI HEADER BAND
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

        // ======================
        // CREATE DETAIL BAND
        // ======================
        JRDesignBand detailBand = new JRDesignBand();
        detailBand.setHeight(20);

        int xDetail = 0;

        for (String fieldName : fieldNames) {

            JRDesignTextField textField = new JRDesignTextField();
            textField.setX(xDetail);
            textField.setY(0);
            textField.setWidth(columnWidth);
            textField.setHeight(20);
            textField.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);

            JRDesignExpression expression = new JRDesignExpression();
            expression.setText("$F{" + fieldName + "}");
            textField.setExpression(expression);

            detailBand.addElement(textField);
            xDetail += columnWidth;
        }

        ((JRDesignSection) design.getDetailSection()).addBand(detailBand);

        // ======================
        // EXPORT JRXML
        // ======================
        JRXmlWriter.writeReport(design, jrxmlPath, "UTF-8");

        workbook.close();
        fis.close();
    }
}