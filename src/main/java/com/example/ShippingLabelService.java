package com.example;

import net.sf.jasperreports.components.barcode4j.Code128Component;
import net.sf.jasperreports.components.barcode4j.QRCodeComponent;
import net.sf.jasperreports.engine.component.ComponentKey;
import net.sf.jasperreports.engine.design.*;
import net.sf.jasperreports.engine.type.HorizontalTextAlignEnum;
import net.sf.jasperreports.engine.type.VerticalTextAlignEnum;
import net.sf.jasperreports.engine.xml.JRXmlWriter;
import org.springframework.stereotype.Service;

import java.io.OutputStream;

@Service
public class ShippingLabelService {

    public void generateShippingLabelJRXML(OutputStream outputStream) throws Exception {
        JasperDesign design = new JasperDesign();
        design.setName("ShippingLabel");
        design.setPageWidth(400);
        design.setPageHeight(400);
        design.setColumnWidth(380);
        design.setLeftMargin(10);
        design.setRightMargin(10);
        design.setTopMargin(10);
        design.setBottomMargin(10);

        // Fields
        addField(design, "senderName", String.class);
        addField(design, "senderPhone", String.class);
        addField(design, "senderAddress", String.class);
        addField(design, "receiverName", String.class);
        addField(design, "receiverPhone", String.class);
        addField(design, "receiverAddress", String.class);
        addField(design, "trackingNumber", String.class);
        addField(design, "orderId", String.class);
        addField(design, "weight", String.class);
        addField(design, "codAmount", String.class);
        addField(design, "totalAmount", String.class);
        addField(design, "note", String.class);
        addField(design, "routingCode", String.class);

        JRDesignBand detail = new JRDesignBand();
        detail.setHeight(380);

        // Border rectangle
        JRDesignRectangle mainRect = new JRDesignRectangle();
        mainRect.setX(0);
        mainRect.setY(0);
        mainRect.setWidth(380);
        mainRect.setHeight(380);
        detail.addElement(mainRect);

        // --- TOP SECTION ---
        // Logo and Text (Vietnam Post)
        addStaticText(detail, 5, 5, 100, 30, "VIETNAM POST", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 10, true);

        // Barcode
        JRDesignComponentElement barcode = new JRDesignComponentElement();
        barcode.setX(110); barcode.setY(5); barcode.setWidth(180); barcode.setHeight(40);
        barcode.setComponentKey(new ComponentKey("http://jasperreports.sourceforge.net/jasperreports/components", "jr", "Code128"));
        Code128Component code128 = new Code128Component();
        code128.setModuleWidth(1.0);
        JRDesignExpression barExp = new JRDesignExpression();
        barExp.setText("$F{trackingNumber}");
        code128.setCodeExpression(barExp);
        barcode.setComponent(code128);
        detail.addElement(barcode);

        JRDesignTextField barTextField = new JRDesignTextField();
        barTextField.setX(110); barTextField.setY(45); barTextField.setWidth(180); barTextField.setHeight(15);
        barTextField.setExpression(new JRDesignExpression("$F{trackingNumber}"));
        barTextField.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
        barTextField.setFontSize(8f);
        detail.addElement(barTextField);

        // Top Right Info
        addStaticText(detail, 300, 5, 75, 15, "Lô:", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 8, false);
        addStaticText(detail, 300, 20, 75, 15, "Thứ: 1", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 8, false);
        addStaticText(detail, 300, 35, 75, 15, "Số ĐH:", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 8, false);

        // --- MIDDLE SECTION 1 (Routing) ---
        JRDesignLine line1 = new JRDesignLine();
        line1.setX(0); line1.setY(60); line1.setWidth(380); line1.setHeight(0);
        detail.addElement(line1);

        JRDesignTextField routingField = new JRDesignTextField();
        routingField.setX(5); routingField.setY(65); routingField.setWidth(330); routingField.setHeight(35);
        routingField.setExpression(new JRDesignExpression("$F{routingCode}"));
        routingField.setHorizontalTextAlign(HorizontalTextAlignEnum.CENTER);
        routingField.setVerticalTextAlign(VerticalTextAlignEnum.MIDDLE);
        routingField.setFontSize(12f);
        routingField.setBold(true);
        detail.addElement(routingField);

        // Box 'C'
        JRDesignRectangle cBox = new JRDesignRectangle();
        cBox.setX(340); cBox.setY(65); cBox.setWidth(35); cBox.setHeight(35);
        detail.addElement(cBox);
        addStaticText(detail, 340, 65, 35, 35, "C", HorizontalTextAlignEnum.CENTER, VerticalTextAlignEnum.MIDDLE, 16, true);

        // --- MIDDLE SECTION 2 (Sender/Receiver) ---
        JRDesignLine line2 = new JRDesignLine();
        line2.setX(0); line2.setY(105); line2.setWidth(380); line2.setHeight(0);
        detail.addElement(line2);

        // Vertical line separating sender/receiver
        JRDesignLine vLine1 = new JRDesignLine();
        vLine1.setX(190); vLine1.setY(105); vLine1.setWidth(0); vLine1.setHeight(155);
        detail.addElement(vLine1);

        // Sender
        JRDesignTextField senderField = new JRDesignTextField();
        senderField.setX(5); senderField.setY(110); senderField.setWidth(180); senderField.setHeight(80);
        senderField.setExpression(new JRDesignExpression("\"Từ: \" + $F{senderName} + \" - \" + $F{senderPhone} + \"\\n\" + $F{senderAddress}"));
        senderField.setFontSize(8f);
        detail.addElement(senderField);

        // Receiver
        JRDesignTextField receiverField = new JRDesignTextField();
        receiverField.setX(195); receiverField.setY(110); receiverField.setWidth(180); receiverField.setHeight(90);
        receiverField.setExpression(new JRDesignExpression("\"Đến: \" + $F{receiverName} + \" - \" + $F{receiverPhone} + \"\\n\" + $F{receiverAddress}"));
        receiverField.setFontSize(10f);
        receiverField.setBold(true);
        detail.addElement(receiverField);

        // Instructions
        JRDesignLine line3 = new JRDesignLine();
        line3.setX(0); line3.setY(190); line3.setWidth(190); line3.setHeight(0);
        detail.addElement(line3);
        addStaticText(detail, 5, 195, 180, 15, "Chỉ dẫn giao:", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 9, true);
        addStaticText(detail, 5, 210, 180, 40, "-", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 8, false);

        // Payment info
        JRDesignLine line4 = new JRDesignLine();
        line4.setX(190); line4.setY(200); line4.setWidth(190); line4.setHeight(0);
        detail.addElement(line4);
        JRDesignTextField paymentField = new JRDesignTextField();
        paymentField.setX(195); paymentField.setY(205); paymentField.setWidth(180); paymentField.setHeight(55);
        paymentField.setExpression(new JRDesignExpression("\"CP TIÊU CHUẨN - HÀNG HÓA\\nKL(gr): \" + $F{weight} + \"\\n- COD: \" + $F{codAmount} + \" đ\\n- Tổng thu: \" + $F{totalAmount} + \" đ\""));
        paymentField.setFontSize(8f);
        detail.addElement(paymentField);

        // --- BOTTOM SECTION ---
        JRDesignLine line5 = new JRDesignLine();
        line5.setX(0); line5.setY(260); line5.setWidth(380); line5.setHeight(0);
        detail.addElement(line5);

        // Note
        JRDesignTextField noteField = new JRDesignTextField();
        noteField.setX(5); noteField.setY(265); noteField.setWidth(250); noteField.setHeight(20);
        noteField.setExpression(new JRDesignExpression("\"ND: \" + $F{note}"));
        noteField.setFontSize(8f);
        detail.addElement(noteField);

        // Date and BC
        JRDesignLine line6 = new JRDesignLine();
        line6.setX(0); line6.setY(290); line6.setWidth(260); line6.setHeight(0);
        detail.addElement(line6);
        addStaticText(detail, 5, 295, 250, 15, "Ngày in: 13h54 ngày 13/04/2026      BC: 155700", HorizontalTextAlignEnum.LEFT, VerticalTextAlignEnum.TOP, 7, false);

        // QR Code
        JRDesignComponentElement qrCode = new JRDesignComponentElement();
        qrCode.setX(270); qrCode.setY(265); qrCode.setWidth(60); qrCode.setHeight(60);
        qrCode.setComponentKey(new ComponentKey("http://jasperreports.sourceforge.net/jasperreports/components", "jr", "QRCode"));
        QRCodeComponent qrCodeComp = new QRCodeComponent();
        JRDesignExpression qrExp = new JRDesignExpression();
        qrExp.setText("$F{trackingNumber}");
        qrCodeComp.setCodeExpression(qrExp);
        qrCode.setComponent(qrCodeComp);
        detail.addElement(qrCode);

        // Signature box
        addStaticText(detail, 270, 325, 100, 25, "Chữ ký người nhận\nNgày...Tháng...Năm 20...", HorizontalTextAlignEnum.CENTER, VerticalTextAlignEnum.TOP, 6, false);
        JRDesignRectangle sigRect = new JRDesignRectangle();
        sigRect.setX(270); sigRect.setY(325); sigRect.setWidth(100); sigRect.setHeight(30);
        sigRect.setMode(net.sf.jasperreports.engine.type.ModeEnum.TRANSPARENT);
        detail.addElement(sigRect);

        // Footer
        JRDesignLine line7 = new JRDesignLine();
        line7.setX(0); line7.setY(360); line7.setWidth(380); line7.setHeight(0);
        detail.addElement(line7);
        addStaticText(detail, 5, 365, 370, 15, "Gọi 1900545481: Tuyển nhân viên giao nhận toàn quốc, công việc gần nhà, thu nhập ổn định.", HorizontalTextAlignEnum.CENTER, VerticalTextAlignEnum.MIDDLE, 7, true);

        ((JRDesignSection) design.getDetailSection()).addBand(detail);

        JRXmlWriter.writeReport(design, outputStream, "UTF-8");
    }

    private void addField(JasperDesign design, String name, Class<?> type) throws Exception {
        JRDesignField field = new JRDesignField();
        field.setName(name);
        field.setValueClass(type);
        design.addField(field);
    }

    private void addStaticText(JRDesignBand band, int x, int y, int w, int h, String text, HorizontalTextAlignEnum hAlign, VerticalTextAlignEnum vAlign, int fontSize, boolean bold) {
        JRDesignStaticText st = new JRDesignStaticText();
        st.setX(x);
        st.setY(y);
        st.setWidth(w);
        st.setHeight(h);
        st.setText(text);
        st.setHorizontalTextAlign(hAlign);
        st.setVerticalTextAlign(vAlign);
        st.setFontSize((float)fontSize);
        st.setBold(bold);
        band.addElement(st);
    }
}
