package com.Library.ConversionXLSXtoDATABASE;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api")
public class ExcelController {

    private final ExcelToDatabaseService excelToDatabaseService;

    public ExcelController(ExcelToDatabaseService excelToDatabaseService) {
        this.excelToDatabaseService = excelToDatabaseService;
    }

    @GetMapping("/upload-xlsx")
    public String uploadExcel(@RequestParam String filePath) {
        filePath = filePath.replace("\\", "/"); // Fix Windows path issue
        System.out.println("Received File Path: " + filePath);

        excelToDatabaseService.processExcelFile(filePath);
        return "XLSX processed and data inserted successfully!";
    }
}