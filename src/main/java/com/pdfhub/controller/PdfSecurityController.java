package com.pdfhub.controller;


import com.pdfhub.config.PdfConfig;
import com.pdfhub.service.PdfSecurityService;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

@RestController
@CrossOrigin(origins = "http://localhost:3000")
@RequestMapping("/pdf/security")
public class PdfSecurityController {

    @Autowired
    PdfSecurityService pdfSecurityService;

    @PostMapping("/addpassword")
    ResponseEntity<?> addPasswordToPdf(@RequestParam("files") MultipartFile file, @RequestParam("password") String password) throws IOException {
        File f = PdfConfig.convertMultipartFileToFile(file);
        PDDocument document = Loader.loadPDF(f);

        AccessPermission ap = new AccessPermission();
        StandardProtectionPolicy spp = new StandardProtectionPolicy(password, password, ap);
        document.protect(spp);
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        document.save(outputStream);
        document.close();
        //HttpHeaders headers = new HttpHeaders();
        // headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=protected.pdf");


        byte[] content = outputStream.toByteArray();
        HttpHeaders headers = new HttpHeaders();
        //headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=protected.pdf");

        headers.setContentDisposition(ContentDisposition.attachment()
                .filename(file.getOriginalFilename()+".pdf")
                .build());
        headers.setContentType(MediaType.APPLICATION_PDF);

        return new ResponseEntity<>(content, headers, HttpStatus.OK);

        // return new ResponseEntity<>(outputStream.toByteArray(), headers, HttpStatus.OK);


    }

    @PostMapping("/unlock")
    public ResponseEntity<byte[]> unlockPdf(@RequestParam("file") MultipartFile file, @RequestParam("password") String password) {
        try {
            // Load the PDF document with the password
            File f = PdfConfig.convertMultipartFileToFile(file);
            PDDocument document = Loader.loadPDF(f, password);

            // Remove the protection
            document.setAllSecurityToBeRemoved(true);

            // Save the document to a byte array output stream
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            document.save(outputStream);
            document.close();

            // Prepare the response entity with the unlocked PDF

            //HttpHeaders headers = new HttpHeaders();
            // headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=protected.pdf");


            byte[] content = outputStream.toByteArray();
            HttpHeaders headers = new HttpHeaders();
            //headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=protected.pdf");

            headers.setContentDisposition(ContentDisposition.attachment()
                    .filename("unlocked.pdf")
                    .build());
            headers.setContentType(MediaType.APPLICATION_PDF);

            return new ResponseEntity<>(content, headers, HttpStatus.OK);

        } catch (IOException e) {
            e.printStackTrace();
            return new ResponseEntity<>(HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}