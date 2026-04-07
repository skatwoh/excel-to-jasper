package com.example;

import net.sf.jasperreports.components.table.DesignCell;
import net.sf.jasperreports.components.table.StandardColumn;
import net.sf.jasperreports.components.table.StandardTable;
import net.sf.jasperreports.engine.JRLineBox;
import net.sf.jasperreports.engine.component.ComponentKey;
import net.sf.jasperreports.engine.data.JRBeanCollectionDataSource;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.*;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

@Service
public class ExcelToJasperService {

    /**
     * Thông tin chi tiết của một cột để mapping.
     */
    public static class ColumnMapping {
        public boolean use = true;
        public String originalName;
        public String fieldName;
        public String label;
        public String paramName;
        public String source = "FIELD";
        public String expression;
        public int width;

        public ColumnMapping() {}
        public ColumnMapping(String originalName, int width) {
            this.originalName = originalName;
            this.fieldName = originalName.replace(" ", "_");
            this.label = originalName;
            this.paramName = "P_" + fieldName;
            this.expression = "$F{" + fieldName + "}";
            this.width = width;
        }
    }

    public List<String> getSheetNames(InputStream inputStream) throws Exception {
        try (Workbook wb = new XSSFWorkbook(inputStream)) {
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                sheetNames.add(wb.getSheetName(i));
            }
            return sheetNames;
        }
    }

    public List<ColumnMapping> analyzeSheet(InputStream inputStream, String sheetName) throws Exception {
        try (Workbook wb = new XSSFWorkbook(inputStream)) {
            Sheet sheet = wb.getSheet(sheetName);
            Row header = sheet.getRow(0);
            if (header == null) return Collections.emptyList();

            int colCount = header.getPhysicalNumberOfCells();
            List<ColumnMapping> mappings = new ArrayList<>();

            for (int i = 0; i < colCount; i++) {
                String name = header.getCell(i).toString();
                int w = sheet.getColumnWidth(i);
                int px = (int) (w / 256.0 * 7);
                if (px < 50) px = 50;

                mappings.add(new ColumnMapping(name, px));
            }
            return mappings;
        }
    }

    public void generateJRXML(List<ColumnMapping> cols, OutputStream outputStream, java.awt.Color headerColor) throws Exception {
        JasperDesign design = new JasperDesign();
        design.setName("FINAL_REPORT");

        int totalContentWidth = 0;
        for (ColumnMapping c : cols) {
            if (c.use) totalContentWidth += c.width;
        }

        design.setColumnWidth(totalContentWidth);
        design.setPageWidth(totalContentWidth + 40);
        design.setPageHeight(842);
        design.setLeftMargin(20);
        design.setRightMargin(20);
        design.setTopMargin(20);
        design.setBottomMargin(20);

        // DATASET
        JRDesignDataset dataset = new JRDesignDataset(false);
        dataset.setName("ItemDataSource");

        JRDesignParameter dsParam = new JRDesignParameter();
        dsParam.setName("ItemDataSource");
        dsParam.setValueClass(JRBeanCollectionDataSource.class);
        design.addParameter(dsParam);

        for (ColumnMapping c : cols) {
            if (!c.use) continue;
            JRDesignField f = new JRDesignField();
            f.setName(c.fieldName);
            f.setValueClass(String.class);
            dataset.addField(f);

            JRDesignParameter p = new JRDesignParameter();
            p.setName(c.paramName);
            p.setValueClass(String.class);
            design.addParameter(p);
        }
        design.addDataset(dataset);

        // TABLE
        StandardTable table = new StandardTable();
        JRDesignDatasetRun run = new JRDesignDatasetRun();
        run.setDatasetName("ItemDataSource");
        run.setDataSourceExpression(new JRDesignExpression("$P{ItemDataSource}"));
        table.setDatasetRun(run);

        for (ColumnMapping c : cols) {
            if (!c.use) continue;

            StandardColumn col = new StandardColumn();
            col.setWidth(c.width);

            // HEADER
            DesignCell header = new DesignCell();
            header.setHeight(30);
            header.getLineBox().getPen().setLineWidth(0.5f);
            header.getLineBox().getPen().setLineStyle(LineStyleEnum.SOLID);

            JRDesignStaticText txt = new JRDesignStaticText();
            txt.setText(c.label);
            txt.setWidth(c.width);
            txt.setHeight(30);
            txt.setBold(true);
            txt.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
            txt.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);
            txt.setBackcolor(headerColor);
            txt.setMode(ModeEnum.OPAQUE);
            header.addElement(txt);
            col.setColumnHeader(header);

            // DETAIL
            DesignCell detail = new DesignCell();
            detail.setHeight(25);
            detail.getLineBox().getPen().setLineWidth(0.5f);
            detail.getLineBox().getPen().setLineStyle(LineStyleEnum.SOLID);

            JRDesignTextField tf = new JRDesignTextField();
            String exp = (c.expression == null || c.expression.isEmpty())
                ? (c.source.equals("PARAM") ? "$P{" + c.paramName + "}" : "$F{" + c.fieldName + "}")
                : c.expression;
            tf.setExpression(new JRDesignExpression(exp));
            tf.setWidth(c.width);
            tf.setHeight(25);
            tf.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);
            detail.addElement(tf);
            col.setDetailCell(detail);

            table.addColumn(col);
        }

        JRDesignComponentElement comp = new JRDesignComponentElement();
        comp.setComponentKey(new ComponentKey("http://jasperreports.sourceforge.net/jasperreports/components", "jr", "table"));
        comp.setComponent(table);
        comp.setX(0);
        comp.setY(0);
        comp.setWidth(totalContentWidth);
        comp.setHeight(60);

        JRDesignBand band = new JRDesignBand();
        band.setHeight(60);
        band.addElement(comp);
        ((JRDesignSection) design.getDetailSection()).addBand(band);

        JRXmlWriter.writeReport(design, outputStream, "UTF-8");
    }
}
