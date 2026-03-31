package com.example;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@SpringBootTest
class ExcelToJasperServiceTest {

    @Autowired
    private ExcelToJasperService excelToJasperService;

    @Test
    void testGetSheetNames() throws Exception {
        File file = new File("sample.xlsx");
        if (!file.exists()) return;

        try (FileInputStream fis = new FileInputStream(file)) {
            List<String> names = excelToJasperService.getSheetNames(fis);
            assertNotNull(names);
            assertFalse(names.isEmpty());
        }
    }

    @Test
    void testConvert() throws Exception {
        File file = new File("sample.xlsx");
        if (!file.exists()) return;

        try (FileInputStream fis = new FileInputStream(file)) {
            List<String> names = excelToJasperService.getSheetNames(new FileInputStream(file));
            String sheetName = names.get(0);

            ByteArrayOutputStream baos = new ByteArrayOutputStream();
            excelToJasperService.convert(new FileInputStream(file), sheetName, baos, 0, 2);

            String result = baos.toString("UTF-8");
            assertTrue(result.contains("<jasperReport"));
            assertTrue(result.contains("name=\"PRO_REPORT\""));
        }
    }
}
