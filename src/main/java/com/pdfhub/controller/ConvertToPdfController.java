package com.pdfhub.controller;

import com.itextpdf.text.DocumentException;
import com.pdfhub.service.ConvertToPdfService;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RestController
@RequestMapping("/api")
public class ConvertToPdfController {

    private final ConvertToPdfService convertToPdfService;
    public ConvertToPdfController(ConvertToPdfService  convertToPdfService){
        this.convertToPdfService=convertToPdfService;
    }

//    //this is working
//    @PostMapping("/convertJpgToPdf")
//    public ResponseEntity<byte[]> convertMultipleJpgToPdf(
//            @RequestParam("files")  MultipartFile[] files,
//            @RequestParam("orientation") String orientation,
//            @RequestParam("pageSize") String pageSize) throws IOException {
//
//        List<byte[]> imageBytesList = PdfConfig.convertFilesToBytes(files);
//        byte[] pdfBytes = convertToPdfService.convertImagesToPdf(imageBytesList, orientation, pageSize);
//
//        HttpHeaders headers = new HttpHeaders();
//        headers.setContentType(MediaType.APPLICATION_PDF);
//        headers.setContentDispositionFormData("attachment", "images.pdf");
//
//        return ResponseEntity.ok()
//                .headers(headers)
//                .body(pdfBytes);
//    }

    //this is working
    //jpg to pdf
    @PostMapping("/convertJpgToPdf")
    public ResponseEntity<byte[]> convertJpgToPdf(@RequestParam MultipartFile file) {
        try {
            byte[] pdfContent = convertToPdfService.convertJpgToPdf(file);

            // Set headers for PDF download
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted_image.pdf");
            headers.add(HttpHeaders.CONTENT_TYPE, "application/pdf");

            // Return the PDF as a byte array
            return new ResponseEntity<>(pdfContent, headers, HttpStatus.OK);
        } catch (IOException e) {
            e.printStackTrace();
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }




    @PostMapping("/convertwordtopdf")
    public ResponseEntity<byte[]> convertWordToPdf(@RequestParam("files") List<MultipartFile> files) {
        try {
            // Convert multiple Word files to PDF
            List<byte[]> pdfFiles = convertToPdfService.convertWordToPdf(files);
            ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
            ZipOutputStream zipOut = new ZipOutputStream(zipOutputStream);

            // Add each PDF to the ZIP file
            int index = 1;
            for (byte[] pdf : pdfFiles) {
                ZipEntry zipEntry = new ZipEntry("file" + index + ".pdf");
                zipOut.putNextEntry(zipEntry);
                zipOut.write(pdf);
                zipOut.closeEntry();
                index++;
            }

            // Close the ZIP output stream
            zipOut.close();
            // Set headers for PDF download (if needed for returning multiple PDFs)
            HttpHeaders headers = new HttpHeaders();
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=converted_files.zip");
            headers.add(HttpHeaders.CONTENT_TYPE, "application/zip");

            // You may want to compress all PDF files into a ZIP before sending them
            return new ResponseEntity<>(zipOutputStream .toByteArray(), headers, HttpStatus.OK);

        } catch (IOException e) {
            e.printStackTrace();
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        } catch (Docx4JException e) {
            throw new RuntimeException(e);
        }
    }


    @PostMapping("/convertexceltopdf")
    public ResponseEntity<List<byte[]>> convertExcelFilesToPdf(@RequestParam("files") MultipartFile[] files) {
        List<byte[]> pdfs = new ArrayList<>();

        try {
            for (MultipartFile file : files) {
                byte[] pdfBytes = convertToPdfService.convertExcelToPdf(file);
                pdfs.add(pdfBytes);  // Add each converted PDF byte array to the list
            }
            return ResponseEntity.ok(pdfs);  // Return the list of PDFs
        } catch (IOException | DocumentException e) {
            e.printStackTrace();
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body(null);  // Return an error response if conversion fails
        }
    }
    @PostMapping("/convertexceltopdf2")
    public ResponseEntity<byte[]> convertExcelToPdf(@RequestParam("file") MultipartFile file) throws IOException, DocumentException {
        byte[] pdfBytes = convertToPdfService.convertExcelToPdf2(file);


        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_PDF);
        headers.setContentDispositionFormData("attachment", "images.pdf");

        return ResponseEntity.ok()
                .headers(headers)
                .body(pdfBytes);


//		HttpHeaders headers = new HttpHeaders();
//		headers.add("Content-Disposition", "attachment; filename=converted.pdf");
//
//		return ResponseEntity.ok()
//				.headers(headers)
//				.body(pdfBytes);
    }

	/*
	Extracting structured data from images using AI in Java
	 */

//working
    @PostMapping("/powerpointtopdf")
    ResponseEntity<?> convertPowerPointToPdf(@RequestParam MultipartFile file) {
        try {
            // Convert the PowerPoint file to a PDF byte array
            byte[] pdfContent = convertToPdfService.convertPowerPointToPdf(file);

            // Set response headers for PDF content
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_PDF);
            headers.add(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + file.getOriginalFilename() + ".pdf");

            // Return the PDF as a byte array
            return new ResponseEntity<>(pdfContent, headers, HttpStatus.OK);
        } catch (IOException e) {
            e.printStackTrace();
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

}
