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

import javax.servlet.http.HttpSession;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

@Controller
public class ExcelToJasperController {

    @Autowired
    private ExcelToJasperService excelToJasperService;

    @GetMapping("/")
    public String index() {
        return "index";
    }

    @PostMapping("/list-sheets")
    public String listSheets(@RequestParam("file") MultipartFile file,
                             Model model,
                             HttpSession session) throws Exception {
        if (file.isEmpty()) {
            model.addAttribute("error", "Please select a file to upload");
            return "index";
        }

        List<String> sheetNames = excelToJasperService.getSheetNames(file.getInputStream());
        model.addAttribute("sheetNames", sheetNames);
        model.addAttribute("fileName", file.getOriginalFilename());

        session.setAttribute("fileBytes", file.getBytes());
        session.setAttribute("fileName", file.getOriginalFilename());

        return "index";
    }

    @PostMapping("/convert")
    public ResponseEntity<byte[]> convert(
            @RequestParam("sheetName") String sheetName,
            @RequestParam("headerStartRow") int headerStartRow,
            @RequestParam("headerRowCount") int headerRowCount,
            HttpSession session) throws Exception {

        byte[] fileBytes = (byte[]) session.getAttribute("fileBytes");
        String fileName = (String) session.getAttribute("fileName");

        if (fileBytes == null) {
            return ResponseEntity.badRequest().build();
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

        byte[] jrxmlBytes = baos.toByteArray();

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + jrxmlFileName + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .contentLength(jrxmlBytes.length)
                .body(jrxmlBytes);
    }
}
