package com.example;

import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.*;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import java.util.List;

public class ExcelToJasperSwingApp1 extends JFrame {

    private JTextField fileField;
    private JTextField headerStartField;
    private JTextField headerCountField;

    private DefaultListModel<String> sheetListModel = new DefaultListModel<>();
    private JList<String> sheetList;
    private JTabbedPane previewTabs;

    public ExcelToJasperSwingApp1() {
        setTitle("Excel → Jasper Converter");
        setSize(1100, 700);
        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JPanel top = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4,4,4,4);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        gbc.gridx=0; top.add(new JLabel("File:"), gbc);
        gbc.gridx=1; gbc.weightx=1;
        fileField = new JTextField();
        top.add(fileField, gbc);

        gbc.gridx=2; gbc.weightx=0;
        JButton browse = new JButton("Browse");
        browse.addActionListener(this::browseFile);
        top.add(browse, gbc);

        gbc.gridx=0; gbc.gridy=1;
        top.add(new JLabel("Header start:"), gbc);
        gbc.gridx=1;
        headerStartField = new JTextField("0");
        top.add(headerStartField, gbc);

        gbc.gridx=0; gbc.gridy=2;
        top.add(new JLabel("Header rows:"), gbc);
        gbc.gridx=1;
        headerCountField = new JTextField("2");
        top.add(headerCountField, gbc);

        add(top, BorderLayout.NORTH);

        sheetList = new JList<>(sheetListModel);
        sheetList.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);

        previewTabs = new JTabbedPane();

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(sheetList), previewTabs);
        split.setDividerLocation(200);
        add(split, BorderLayout.CENTER);

        JPanel bottom = new JPanel();
        JButton load = new JButton("Load Sheets");
        load.addActionListener(this::loadSheets);
        JButton preview = new JButton("Preview");
        preview.addActionListener(this::preview);
        JButton gen = new JButton("Generate JRXML");
        gen.addActionListener(this::generate);

        bottom.add(load);
        bottom.add(preview);
        bottom.add(gen);

        add(bottom, BorderLayout.SOUTH);
    }

    private void browseFile(ActionEvent e){
        JFileChooser ch = new JFileChooser();
        if(ch.showOpenDialog(this)==JFileChooser.APPROVE_OPTION){
            fileField.setText(ch.getSelectedFile().getAbsolutePath());
        }
    }

    private void loadSheets(ActionEvent e){
        try(FileInputStream fis=new FileInputStream(fileField.getText());
            Workbook wb=new XSSFWorkbook(fis)){
            sheetListModel.clear();
            for(int i=0;i<wb.getNumberOfSheets();i++){
                sheetListModel.addElement(wb.getSheetName(i));
            }
        }catch(Exception ex){ex.printStackTrace();}
    }

    private void preview(ActionEvent e){
        try(FileInputStream fis=new FileInputStream(fileField.getText());
            Workbook wb=new XSSFWorkbook(fis)){
            previewTabs.removeAll();
            for(String s: sheetList.getSelectedValuesList()){
                Sheet sheet=wb.getSheet(s);
                previewTabs.add(s,new JScrollPane(buildTable(sheet)));
            }
        }catch(Exception ex){ex.printStackTrace();}
    }

    private JTable buildTable(Sheet sheet){
        int cols=sheet.getRow(0).getPhysicalNumberOfCells();
        DefaultTableModel m=new DefaultTableModel();
        for(int i=0;i<cols;i++) m.addColumn("C"+i);

        for(Row r:sheet){
            Vector<String> v=new Vector<>();
            for(int i=0;i<cols;i++){
                Cell c=r.getCell(i);
                v.add(c==null?"":c.toString());
            }
            m.addRow(v);
        }
        return new JTable(m);
    }

    private void generate(ActionEvent e){
        try{
            int start=Integer.parseInt(headerStartField.getText());
            int count=Integer.parseInt(headerCountField.getText());

            File f=new File(fileField.getText());
            String base=f.getName().replace(".xlsx","");

            for(String s:sheetList.getSelectedValuesList()){
                convert(fileField.getText(),
                        f.getParent()+"/"+base+"_"+s+".jrxml",
                        s,start,count);
            }

            JOptionPane.showMessageDialog(this,"Done!");
        }catch(Exception ex){ex.printStackTrace();}
    }

    // ========================= CORE CONVERT =========================

    private static float excelWidthToPixel(int width){
        return (float)Math.floor((width/256.0)*7+5);
    }

    public static void convert(String excelPath,String jrxmlPath,
                               String sheetName,int headerStart,int headerCount) throws Exception{

        try(FileInputStream fis=new FileInputStream(excelPath);
            Workbook wb=new XSSFWorkbook(fis)){

            Sheet sheet=wb.getSheet(sheetName);

            // AUTO SIZE (giống Excel)
            Row lastHeader = sheet.getRow(headerStart+headerCount-1);
            int colCount = lastHeader.getPhysicalNumberOfCells();

            for(int i=0;i<colCount;i++) sheet.autoSizeColumn(i);

            float[] px=new float[colCount];
            float total=0;

            for(int i=0;i<colCount;i++){
                float p=excelWidthToPixel(sheet.getColumnWidth(i));
                px[i]=p; total+=p;
            }

            int PAGE=555;
            int MIN=30, MAX=120;

            int[] w=new int[colCount];
            int used=0;
            for(int i=0;i<colCount-1;i++){
                int val=Math.round(px[i]/total*PAGE);
                val=Math.max(MIN,Math.min(MAX,val));
                w[i]=val; used+=val;
            }
            w[colCount-1]=PAGE-used;

            int[] x=new int[colCount];
            for(int i=1;i<colCount;i++) x[i]=x[i-1]+w[i-1];

            JasperDesign d=new JasperDesign();
            d.setName(sheetName);
            d.setPageWidth(595);
            d.setPageHeight(842);
            d.setColumnWidth(PAGE);

            List<String> fields=new ArrayList<>();

            for(int i=0;i<colCount;i++){
                String fn="COL_"+i;
                fields.add(fn);
                JRDesignField f=new JRDesignField();
                f.setName(fn);
                f.setValueClass(String.class);
                d.addField(f);
            }

            int ROW_H=22;
            JRDesignBand header=new JRDesignBand();
            header.setHeight(ROW_H*headerCount);

            // MERGE
            for(int m=0;m<sheet.getNumMergedRegions();m++){
                CellRangeAddress r=sheet.getMergedRegion(m);

                if(r.getFirstRow()<headerStart ||
                        r.getLastRow()>=headerStart+headerCount) continue;

                int c1=r.getFirstColumn();
                int c2=r.getLastColumn();

                int xx=x[c1];
                int ww=0;
                for(int c=c1;c<=c2;c++) ww+=w[c];

                int yy=(r.getFirstRow()-headerStart)*ROW_H;
                int hh=(r.getLastRow()-r.getFirstRow()+1)*ROW_H;

                String txt=sheet.getRow(r.getFirstRow()).getCell(c1).toString();

                header.addElement(makeHeader(txt,xx,yy,ww,hh));
            }

            d.setColumnHeader(header);

            // DETAIL
            JRDesignBand detail=new JRDesignBand();
            detail.setHeight(22);

            for(int i=0;i<colCount;i++){
                JRDesignTextField tf=new JRDesignTextField();
                tf.setX(x[i]);
                tf.setWidth(w[i]);
                tf.setHeight(22);
                tf.setPositionType(PositionTypeEnum.FLOAT);
                tf.setStretchWithOverflow(true);
                tf.setFontName("Calibri");
                tf.setFontSize(11f);

                JRDesignExpression ex=new JRDesignExpression();
                ex.setText("$F{"+fields.get(i)+"}");
                tf.setExpression(ex);

                detail.addElement(tf);
            }

            ((JRDesignSection)d.getDetailSection()).addBand(detail);

            JRXmlWriter.writeReport(d,jrxmlPath,"UTF-8");
        }
    }

    private static JRDesignStaticText makeHeader(String t,int x,int y,int w,int h){
        JRDesignStaticText st=new JRDesignStaticText();
        st.setX(x); st.setY(y);
        st.setWidth(w); st.setHeight(h);
        st.setText(t);

        st.setFontName("Calibri");
        st.setFontSize(11f);
        st.setBold(true);

        st.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
        st.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);

        st.setStretchType(StretchTypeEnum.CONTAINER_HEIGHT);
        st.setMode(ModeEnum.OPAQUE);

        st.getLineBox().getPen().setLineWidth(0.25f);

        return st;
    }

    public static void main(String[] args){
        SwingUtilities.invokeLater(()-> new ExcelToJasperSwingApp1().setVisible(true));
    }
}