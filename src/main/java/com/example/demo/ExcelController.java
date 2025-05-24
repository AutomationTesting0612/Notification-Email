package com.example.demo;

import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Controller
public class ExcelController {

    @Autowired
    ExcelService excelService;

    @Autowired
    EmailService emailService;

    @GetMapping("/")
    public String index() {
        return "upload";
    }

    @PostMapping("/upload")
    public String uploadExcel(@RequestParam("file") MultipartFile file, Model model, HttpSession session) throws Exception {
        List<List<String>> data = excelService.readExcel(file);
        model.addAttribute("data", data);
        session.setAttribute("originalData", data);
        return "upload";
    }

    @PostMapping("/submit")
    public String submitData(@RequestParam Map<String, String> formData, HttpSession session, Model model) throws Exception {
        @SuppressWarnings("unchecked")
        List<List<String>> originalData = (List<List<String>>) session.getAttribute("originalData");
        List<List<String>> updatedData = new ArrayList<>();

        int rowCount = Integer.parseInt(formData.get("rows"));
        int colCount = Integer.parseInt(formData.get("cols"));

        StringBuilder emailBody = new StringBuilder("🔔 *Changed Excel Data:*\n\n");

        for (int i = 0; i < rowCount; i++) {
            List<String> updatedRow = new ArrayList<>();
            boolean rowChanged = false;

            for (int j = 0; j < colCount; j++) {
                String key = "cell_" + i + "_" + j;
                String newVal = formData.getOrDefault(key, "");
                String oldVal = originalData != null && originalData.size() > i && originalData.get(i).size() > j
                        ? originalData.get(i).get(j) : "";

                updatedRow.add(newVal);

                if (!newVal.equals(oldVal)) {
                    rowChanged = true;
                }
            }

            updatedData.add(updatedRow);

            boolean isNewRow = originalData == null || i >= originalData.size();

            if (isNewRow) {
                emailBody.append("🟢 New Row: ").append(String.join(", ", updatedRow)).append("\n");
            } else if (rowChanged) {
                emailBody.append("🟡 Updated Row: ").append(String.join(", ", updatedRow)).append("\n");
            }
        }

        // Send email only if any changes occurred
        if (!emailBody.toString().equals("🔔 *Changed Excel Data:*\n\n")) {
            emailService.sendEmail("🔔 * Excel Sheet Notification", emailBody.toString());

            // Save updated Excel file
            File updatedFile = new File("datatest.xlsx");
            excelService.writeExcel(updatedData, updatedFile);
        }

        // Update session with latest data
        session.setAttribute("originalData", updatedData);
        model.addAttribute("data", updatedData);

        return "upload";
    }



}

