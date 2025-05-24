package com.example.demo;

import com.fasterxml.jackson.databind.ObjectMapper;
import jakarta.mail.*;
import jakarta.mail.internet.InternetAddress;
import jakarta.mail.internet.MimeMessage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;

@Service
public class ExcelMonitorService {

    private static final String EXCEL_FILE = "datatest.xlsx";
    private static final String SNAPSHOT_FILE = "snapshot.json";

    @Scheduled(cron = "0 */1 * * * *") // every 1 minute
    public void checkExcelChanges() throws Exception {
        List<List<String>> currentData = readExcel(EXCEL_FILE);
        List<List<String>> oldData = loadSnapshot();

        String changes = detectChanges(oldData, currentData);

        if (!changes.isEmpty()) {
            sendEmail("🟢 Updated Record: ", changes);
            saveSnapshot(currentData);
        } else {
            System.out.println("✅ No changes detected at " + new Date());
        }
    }

    private List<List<String>> readExcel(String path) throws Exception {
        List<List<String>> data = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(new FileInputStream(path))) {
            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                List<String> rowData = new ArrayList<>();
                for (Cell cell : row) {
                    rowData.add(cell.toString().trim());
                }
                data.add(rowData);
            }
        }
        return data;
    }

    private List<List<String>> loadSnapshot() {
        ObjectMapper mapper = new ObjectMapper();
        try {
            return mapper.readValue(new File(SNAPSHOT_FILE), List.class);
        } catch (IOException e) {
            return new ArrayList<>();
        }
    }

    private void saveSnapshot(List<List<String>> data) throws IOException {
        ObjectMapper mapper = new ObjectMapper();
        mapper.writeValue(new File(SNAPSHOT_FILE), data);
    }

    private String detectChanges(List<List<String>> oldData, List<List<String>> newData) {
        StringBuilder changes = new StringBuilder();
        int max = Math.max(oldData.size(), newData.size());

        for (int i = 0; i < max; i++) {
            if (i >= oldData.size()) {
                changes.append("➕ Row Added: ").append(newData.get(i)).append("\n");
            } else if (i >= newData.size()) {
                changes.append("❌ Row Deleted: ").append(oldData.get(i)).append("\n");
            } else if (!oldData.get(i).equals(newData.get(i))) {
                changes.append("✏️ Row Updated:\n")
                        .append("   Old: ").append(oldData.get(i)).append("\n")
                        .append("   New: ").append(newData.get(i)).append("\n");
            }
        }
        return changes.toString();
    }

    private void sendEmail(String subject, String body) throws MessagingException {
        String from = "akhil.sharma0612@gmail.com";
        String password = "mxfx rflx djuf undm";
        String to = "technologylearning022@gmail.com";

        Properties props = new Properties();
        props.put("mail.smtp.host", "smtp.gmail.com");
        props.put("mail.smtp.port", "587");
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");

        Session session = Session.getInstance(props,
                new Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(from, password);
                    }
                });

        Message message = new MimeMessage(session);
        message.setFrom(new InternetAddress(from));
        message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(to));
        message.setSubject(subject);
        message.setText(body);

        Transport.send(message);
        System.out.println("📧 Email sent with Excel changes.");
    }
}

