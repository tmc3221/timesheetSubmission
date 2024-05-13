package com.example.timesheetsubmission.entries;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.google.common.net.HttpHeaders;

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

            return ResponseEntity.status(HttpStatus.FOUND).header(HttpHeaders.LOCATION, "/success.html").build();
           //return "redirect:/success"; // Redirect to the success page
        } catch (IOException | MessagingException | EncryptedDocumentException e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to upload file");
            //return "redirect:/error";
        }
    }


    private void sendEmailWithAttachment(String fileName, byte[] fileBytes) throws MessagingException {
        jakarta.mail.internet.MimeMessage message = emailSender.createMimeMessage();
        MimeMessageHelper helper;
        try {
            helper = new MimeMessageHelper(message, true);
            helper.setFrom("tristanmcurtis844@gmail.com"); // Update with your email address
            helper.setTo("tmcurti4@ncsu.edu"); // Update with recipient's email address
            helper.setSubject("Timesheet Uploaded");
            helper.setText("Please find the attached timesheet file.");

            helper.addAttachment(fileName, new ByteArrayResource(fileBytes));
            emailSender.send(message);
        } catch (jakarta.mail.MessagingException e) {
            e.printStackTrace();
        }
    }

    @GetMapping("/success")
    public String successPage() {
        return "success";
    }

    @ExceptionHandler(Exception.class)
    public ResponseEntity<String> handleException(Exception e) {
        e.printStackTrace();
        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Failed to upload file");
    }
}