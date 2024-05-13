package com.example.timesheetsubmission.entries;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.mail.MessagingException;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

@RestController
public class ExcelController {

    @Autowired
    private JavaMailSender emailSender;

    @PostMapping("/uploadExcel")
    public ResponseEntity<String> uploadExcel(@RequestParam("file") MultipartFile file) {
        try {
            // Process the file
            Workbook workbook = WorkbookFactory.create(file.getInputStream());
            Sheet sheet = workbook.getSheetAt(0); // Assuming the timesheet is in the first sheet

            // Extract bytes from the workbook
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            workbook.write(bos);
            byte[] bytes = bos.toByteArray();

            // Email the file
            sendEmailWithAttachment(file.getOriginalFilename(), bytes);

            return ResponseEntity.ok("File uploaded successfully");
        } catch (IOException | MessagingException | EncryptedDocumentException e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to upload file");
        }
    }


    private void sendEmailWithAttachment(String fileName, byte[] fileBytes) throws MessagingException {
        jakarta.mail.internet.MimeMessage message = emailSender.createMimeMessage();
        MimeMessageHelper helper;
        try {
            helper = new MimeMessageHelper(message, true);
            helper.setFrom("mt0574576@gmail.com"); // Update with your email address
            helper.setTo("tristanmcurtis844@gmail.com"); // Update with recipient's email address
            helper.setSubject("Timesheet Uploaded");
            helper.setText("Please find the attached timesheet file.");

            helper.addAttachment(fileName, new ByteArrayResource(fileBytes));
            emailSender.send(message);
        } catch (jakarta.mail.MessagingException e) {
            e.printStackTrace();
        }
    }
}