package com.example.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class ExcelService {

    @Autowired
    EmailService emailService;

    public List<List<String>> readExcel(MultipartFile file) throws IOException {
        List<List<String>> data = new ArrayList<>();
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            List<String> rowData = new ArrayList<>();
            for (Cell cell : row) {
                rowData.add(cell.toString());
            }
            data.add(rowData);
        }
        return data;
    }

    public void writeExcel(List<List<String>> data, File file) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Sheet1");

            for (int i = 0; i < data.size(); i++) {
                Row row = sheet.createRow(i);
                List<String> rowData = data.get(i);
                for (int j = 0; j < rowData.size(); j++) {
                    row.createCell(j).setCellValue(rowData.get(j));
                }
            }

            try (FileOutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
            }
        }
    }

    public void processFormData(Map<String, String> formData) throws Exception {
        List<List<String>> rows = new ArrayList<>();

        int rowCount = Integer.parseInt(formData.get("rows"));
        int colCount = Integer.parseInt(formData.get("cols"));

        for (int i = 0; i < rowCount; i++) {
            List<String> row = new ArrayList<>();
            for (int j = 0; j < colCount; j++) {
                row.add(formData.getOrDefault("cell_" + i + "_" + j, ""));
            }
            rows.add(row);
        }

        StringBuilder emailBody = new StringBuilder("Modified Excel Data:\n\n");
        for (List<String> row : rows) {
            emailBody.append(String.join(", ", row)).append("\n");
        }

        emailService.sendEmail("Excel Sheet Notification", emailBody.toString());
    }
}

