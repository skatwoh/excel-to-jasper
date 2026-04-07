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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;

@Service
public class ExcelToJasperService {

    public List<String> getSheetNames(InputStream inputStream) throws Exception {
        try (Workbook wb = new XSSFWorkbook(inputStream)) {
            List<String> sheetNames = new ArrayList<>();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                sheetNames.add(wb.getSheetName(i));
            }
            return sheetNames;
        }
    }

    public void convert(InputStream inputStream,
                        String sheetName,
                        OutputStream outputStream,
                        int headerStartRow,
                        int headerRowCount) throws Exception {

        try (Workbook wb = new XSSFWorkbook(inputStream)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

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
            // FIELDS & DATASET
            // ======================
            JRDesignParameter dsParam = new JRDesignParameter();
            dsParam.setName("ItemDataSource");
            dsParam.setValueClass(JRBeanCollectionDataSource.class);
            design.addParameter(dsParam);

            JRDesignDataset dataset = new JRDesignDataset(false);
            dataset.setName("ItemDataSource");

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
                dataset.addField(f);
            }
            design.addDataset(dataset);

            // ======================
            // MERGE MAP
            // ======================
            Map<String, CellRangeAddress> mergeMap = new HashMap<>();

            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                CellRangeAddress r = sheet.getMergedRegion(i);
                mergeMap.put(r.getFirstRow() + "_" + r.getFirstColumn(), r);
            }

            // ======================
            // TABLE
            // ======================
            StandardTable table = new StandardTable();

            JRDesignDatasetRun datasetRun = new JRDesignDatasetRun();
            datasetRun.setDatasetName("ItemDataSource");
            datasetRun.setDataSourceExpression(new JRDesignExpression("$P{ItemDataSource}"));
            table.setDatasetRun(datasetRun);

            for (int i = 0; i < fields.size(); i++) {
                StandardColumn column = new StandardColumn();
                column.setWidth(colWidths.get(i));

                // HEADER
                DesignCell headerCell = new DesignCell();
                headerCell.setHeight(30);
                headerCell.getLineBox().getPen().setLineWidth(0.5f);

                JRDesignStaticText headerText = new JRDesignStaticText();
                headerText.setWidth(colWidths.get(i));
                headerText.setHeight(30);
                headerText.setText(fields.get(i));
                headerText.setBold(true);
                headerText.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
                headerText.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);

                headerCell.addElement(headerText);
                column.setColumnHeader(headerCell);

                // DETAIL
                DesignCell detailCell = new DesignCell();
                detailCell.setHeight(25);
                detailCell.getLineBox().getPen().setLineWidth(0.5f);

                JRDesignTextField tf = new JRDesignTextField();
                tf.setWidth(colWidths.get(i));
                tf.setHeight(25);
                tf.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);
                tf.setExpression(new JRDesignExpression("$F{" + fields.get(i) + "}"));

                detailCell.addElement(tf);
                column.setDetailCell(detailCell);

                table.addColumn(column);
            }

            JRDesignComponentElement componentElement = new JRDesignComponentElement();
            componentElement.setX(0);
            componentElement.setY(0);
            componentElement.setWidth(totalWidth);
            componentElement.setHeight(60);
            componentElement.setComponentKey(new ComponentKey("http://jasperreports.sourceforge.net/jasperreports/components", "jr", "table"));
            componentElement.setComponent(table);

            JRDesignBand detailBand = new JRDesignBand();
            detailBand.setHeight(60);
            detailBand.addElement(componentElement);

            ((JRDesignSection) design.getDetailSection()).addBand(detailBand);

            // ======================
            // EXPORT
            // ======================
            JRXmlWriter.writeReport(design, outputStream, "UTF-8");
        }
    }

    // ======================
    // UTIL
    // ======================

    private boolean isMergedButNotFirst(Sheet sheet, int row, int col) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress r = sheet.getMergedRegion(i);
            if (r.isInRange(row, col)) {
                return !(r.getFirstRow() == row && r.getFirstColumn() == col);
            }
        }
        return false;
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        return cell.toString();
    }

    private void applyStyle(JRDesignTextElement element, Cell cell) {
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
