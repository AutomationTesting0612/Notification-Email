//package com.example.demo;
//
//import jakarta.annotation.PostConstruct;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.beans.factory.annotation.Value;
//import org.springframework.scheduling.annotation.Scheduled;
//import org.springframework.stereotype.Component;
//import org.springframework.web.multipart.MultipartFile;
//import org.springframework.mock.web.MockMultipartFile;
//
//import java.io.FileInputStream;
//import java.io.InputStream;
//import java.nio.file.Files;
//import java.nio.file.Paths;
//import java.util.ArrayList;
//import java.util.List;
//
//@Component
//public class ExcelMonitor {
//
//    @Autowired
//    private ExcelService excelService;
//
//    @Autowired
//    private EmailService emailService;
//
//    private List<List<String>> previousSnapshot = new ArrayList<>();
//
//    @Value("${excel.file.path}")
//    private String excelFilePath;
//
//    @PostConstruct
//    public void init() throws Exception {
//        previousSnapshot = readCurrentExcel();
//    }
//
//    @Scheduled(cron = "0 0/5 * * * ?") // Runs every 5 minutes
//    public void checkExcelChanges() {
//        try {
//            List<List<String>> currentSnapshot = readCurrentExcel();
//            List<List<String>> changedRows = new ArrayList<>();
//
//            for (int i = 0; i < currentSnapshot.size(); i++) {
//                List<String> newRow = currentSnapshot.get(i);
//                List<String> oldRow = i < previousSnapshot.size() ? previousSnapshot.get(i) : new ArrayList<>();
//
//                if (!newRow.equals(oldRow)) {
//                    changedRows.add(newRow);
//                }
//            }
//
//            if (!changedRows.isEmpty()) {
//                StringBuilder emailBody = new StringBuilder("Detected changes in Excel file:\n\n");
//                for (List<String> row : changedRows) {
//                    emailBody.append(String.join(", ", row)).append("\n");
//                }
//                emailService.sendEmail("Excel Sheet Auto Notification", emailBody.toString());
//                previousSnapshot = currentSnapshot; // Update snapshot
//            }
//
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//    }
//
//    private List<List<String>> readCurrentExcel() throws Exception {
//        try (InputStream is = new FileInputStream(excelFilePath)) {
//            MultipartFile multipartFile = new MockMultipartFile("file", Files.readAllBytes(Paths.get(excelFilePath)));
//            return excelService.readExcel(multipartFile);
//        }
//    }
//}
//
