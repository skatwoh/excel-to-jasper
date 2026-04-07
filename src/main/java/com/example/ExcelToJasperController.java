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

    private static final int FREE_LIMIT = 5;
    private static final String VALID_LICENSE = "PRO-2024-EXCEL-JASPER";

    @GetMapping("/")
    public String index(HttpSession session, Model model) {
        // Initialize usage count if not exists
        if (session.getAttribute("usageCount") == null) {
            session.setAttribute("usageCount", 0);
        }

        model.addAttribute("usageCount", session.getAttribute("usageCount"));
        model.addAttribute("freeLimit", FREE_LIMIT);

        String license = (String) session.getAttribute("license");
        boolean isPro = VALID_LICENSE.equals(license);
        model.addAttribute("isPro", isPro);
        model.addAttribute("license", license);

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

    @PostMapping("/set-license")
    public String setLicense(@RequestParam("license") String license, HttpSession session) {
        session.setAttribute("license", license.trim());
        return "redirect:/";
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

        int usageCount = (int) session.getAttribute("usageCount");
        String license = (String) session.getAttribute("license");
        boolean isPro = VALID_LICENSE.equals(license);

        if (!isPro && usageCount >= FREE_LIMIT) {
            redirectAttributes.addFlashAttribute("error", "Trial limit reached (" + FREE_LIMIT + "). Please enter a valid license key.");
            return "redirect:/";
        }

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

        session.setAttribute("usageCount", usageCount + 1);

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

    @GetMapping("/reset")
    public String reset(HttpSession session) {
        session.invalidate();
        return "redirect:/";
    }
}
