package com.example;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.servlet.http.HttpSession;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Controller
public class ExcelToJasperController {

    @Autowired
    private ExcelToJasperService excelToJasperService;

    @Autowired
    private ShippingLabelService shippingLabelService;

    @GetMapping("/")
    public String index(HttpSession session, Model model) {
        // Carry over attributes from session if they exist
        if (session.getAttribute("sheetNames") != null) {
            model.addAttribute("sheetNames", session.getAttribute("sheetNames"));
            model.addAttribute("fileName", session.getAttribute("fileName"));
        }
        if (session.getAttribute("jrxmlBytes") != null) {
            model.addAttribute("downloadAvailable", true);
            model.addAttribute("jrxmlFileName", session.getAttribute("jrxmlFileName"));
        }
        return "index";
    }

    @PostMapping("/list-sheets")
    public String listSheets(@RequestParam("file") MultipartFile file,
                             HttpSession session,
                             RedirectAttributes redirectAttributes) throws Exception {
        if (file.isEmpty()) {
            redirectAttributes.addFlashAttribute("error", "Please select a file to upload");
            return "redirect:/";
        }

        List<String> sheetNames = excelToJasperService.getSheetNames(file.getInputStream());

        session.setAttribute("fileBytes", file.getBytes());
        session.setAttribute("fileName", file.getOriginalFilename());
        session.setAttribute("sheetNames", sheetNames);

        // Clear previous download if any
        session.removeAttribute("jrxmlBytes");
        session.removeAttribute("jrxmlFileName");

        return "redirect:/";
    }

    @PostMapping("/convert")
    public String convert(
            @RequestParam("sheetName") String sheetName,
            @RequestParam("headerStartRow") int headerStartRow,
            @RequestParam("headerRowCount") int headerRowCount,
            HttpSession session,
            RedirectAttributes redirectAttributes) throws Exception {

        byte[] fileBytes = (byte[]) session.getAttribute("fileBytes");
        String fileName = (String) session.getAttribute("fileName");

        if (fileBytes == null) {
            redirectAttributes.addFlashAttribute("error", "No file uploaded");
            return "redirect:/";
        }

        String jrxmlFileName = fileName.replace(".xlsx", "_" + sheetName + ".jrxml");

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        excelToJasperService.convert(
                new ByteArrayInputStream(fileBytes),
                sheetName,
                baos,
                headerStartRow,
                headerRowCount
        );

        session.setAttribute("jrxmlBytes", baos.toByteArray());
        session.setAttribute("jrxmlFileName", jrxmlFileName);

        redirectAttributes.addFlashAttribute("success", "Conversion successful! You can now download the file.");
        return "redirect:/";
    }

    @GetMapping("/download")
    public ResponseEntity<byte[]> download(HttpSession session) {
        byte[] jrxmlBytes = (byte[]) session.getAttribute("jrxmlBytes");
        String jrxmlFileName = (String) session.getAttribute("jrxmlFileName");

        if (jrxmlBytes == null || jrxmlFileName == null) {
            return ResponseEntity.notFound().build();
        }

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + jrxmlFileName + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(jrxmlBytes.length)
                .body(jrxmlBytes);
    }

    @PostMapping("/generate-shipping-label")
    public ResponseEntity<byte[]> generateShippingLabel(@RequestParam("image") MultipartFile image) throws Exception {
        // Logic currently generates a template based on the recognized pattern of shipping labels
        // In a real OCR scenario, we would analyze the 'image' here.

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        shippingLabelService.generateShippingLabelJRXML(baos);

        String fileName = image.getOriginalFilename();
        String jrxmlName = (fileName != null && fileName.contains("."))
                ? fileName.substring(0, fileName.lastIndexOf(".")) + ".jrxml"
                : "shipping_label_template.jrxml";

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + jrxmlName + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(baos.toByteArray());
    }

    @GetMapping("/reset")
    public String reset(HttpSession session) {
        session.invalidate();
        return "redirect:/";
    }
}
