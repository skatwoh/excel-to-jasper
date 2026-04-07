package com.example;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
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
    private static int globalUsageCount = 0;

    @GetMapping("/")
    public String index(HttpSession session, Model model, @RequestParam(value = "sheet", required = false) String sheet) throws Exception {
        if (session.getAttribute("usageCount") == null) session.setAttribute("usageCount", 0);

        model.addAttribute("usageCount", session.getAttribute("usageCount"));
        model.addAttribute("freeLimit", FREE_LIMIT);

        String license = (String) session.getAttribute("license");
        model.addAttribute("isPro", LicenseManager.isValid(license));
        model.addAttribute("license", license);

        if (session.getAttribute("sheetNames") != null) {
            model.addAttribute("sheetNames", session.getAttribute("sheetNames"));
            model.addAttribute("fileName", session.getAttribute("fileName"));

            String selectedSheet = sheet != null ? sheet : (String) session.getAttribute("selectedSheet");
            if (selectedSheet == null) selectedSheet = ((List<String>) session.getAttribute("sheetNames")).get(0);

            session.setAttribute("selectedSheet", selectedSheet);
            model.addAttribute("selectedSheet", selectedSheet);

            byte[] fileBytes = (byte[]) session.getAttribute("fileBytes");
            model.addAttribute("mappings", excelToJasperService.analyzeSheet(new ByteArrayInputStream(fileBytes), selectedSheet));
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
    public String listSheets(@RequestParam("file") MultipartFile file, HttpSession session, RedirectAttributes ra) throws Exception {
        if (file.isEmpty()) {
            ra.addFlashAttribute("error", "Please select a file");
            return "redirect:/";
        }
        session.setAttribute("fileBytes", file.getBytes());
        session.setAttribute("fileName", file.getOriginalFilename());
        session.setAttribute("sheetNames", excelToJasperService.getSheetNames(file.getInputStream()));
        return "redirect:/";
    }

    // Wrapper class for mapping list binding
    public static class MappingForm {
        private List<ExcelToJasperService.ColumnMapping> mappings;
        public List<ExcelToJasperService.ColumnMapping> getMappings() { return mappings; }
        public void setMappings(List<ExcelToJasperService.ColumnMapping> m) { this.mappings = m; }
    }

    @PostMapping("/convert")
    public String convert(
            @RequestParam("sheetName") String sheetName,
            @RequestParam("headerColor") String headerColorHex,
            @ModelAttribute MappingForm form,
            HttpSession session,
            RedirectAttributes ra) throws Exception {

        int usageCount = (int) session.getAttribute("usageCount");
        if (!LicenseManager.isValid((String) session.getAttribute("license")) && usageCount >= FREE_LIMIT) {
            ra.addFlashAttribute("error", "Trial limit reached (" + FREE_LIMIT + ").");
            return "redirect:/";
        }

        String fileName = (String) session.getAttribute("fileName");
        String jrxmlFileName = fileName.replace(".xlsx", "_" + sheetName + ".jrxml");
        java.awt.Color headerColor = java.awt.Color.decode(headerColorHex);

        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        excelToJasperService.generateJRXML(form.getMappings(), baos, headerColor);

        session.setAttribute("jrxmlBytes", baos.toByteArray());
        session.setAttribute("jrxmlFileName", jrxmlFileName);
        session.setAttribute("usageCount", usageCount + 1);
        globalUsageCount++;

        ra.addFlashAttribute("success", "Report generated successfully!");
        return "redirect:/";
    }

    @GetMapping("/download")
    public ResponseEntity<byte[]> download(HttpSession session) {
        byte[] b = (byte[]) session.getAttribute("jrxmlBytes");
        String n = (String) session.getAttribute("jrxmlFileName");
        if (b == null) return ResponseEntity.notFound().build();
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + n + "\"")
                .contentType(MediaType.APPLICATION_OCTET_STREAM).body(b);
    }

    @GetMapping("/reset")
    public String reset(HttpSession session) {
        session.invalidate();
        return "redirect:/";
    }
}
