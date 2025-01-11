//package com.pdfhub.controller;
//
//import com.pdfhub.service.AiService;
//import org.springframework.beans.factory.annotation.Autowired;
//import org.springframework.core.io.Resource;
//import org.springframework.http.HttpStatus;
//import org.springframework.http.ResponseEntity;
//import org.springframework.web.bind.annotation.PostMapping;
//import org.springframework.web.bind.annotation.RequestMapping;
//import org.springframework.web.bind.annotation.RequestParam;
//import org.springframework.web.bind.annotation.RestController;
//import org.springframework.web.multipart.MultipartFile;
//
//
//
//import java.io.File;
//import java.io.IOException;
//
//@RestController
//@RequestMapping("/api")
//public class AIController {
//
//
//    @Autowired
//   private AiService aiService;
//
//
////    @PostMapping("/summarize")
////    public ResponseEntity<String> summarizePdf(@RequestParam("file") MultipartFile file) {
////        try {
////            // Save the file to a temporary location
////            File tempFile = File.createTempFile("uploaded", ".pdf");
////            file.transferTo(tempFile);
////
////            // Extract text from PDF
////            String text = aiService.extractTextFromPdf(tempFile);
////            // Summarize the text
////            String summary = aiService.summarizeText(text);
////
////            // Clean up the temporary file
////            tempFile.delete();
////
////            return ResponseEntity.status(HttpStatus.OK).body(summary);
////
////        } catch (IOException e) {
////            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing the PDF file.");
////        }
////    }
//
//
//    @PostMapping("/summarize")
//    public ResponseEntity<String> summarizePdf(@RequestParam("file") MultipartFile file) {
//        try {
//            // Extract text from PDF directly using MultipartFile's InputStream
//            String text = aiService.extractTextFromPdf(file.getInputStream());
//
//            // Summarize the text
//            String summary = aiService.summarizeText(text);
//
//            return ResponseEntity.status(HttpStatus.OK).body(summary);
//
//        } catch (IOException e) {
//            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing the PDF file.");
//        }
//    }
//
//
////@PostMapping("/translate")
////public ResponseEntity<?> translatePdf(
////        @RequestParam("file") MultipartFile file,
////        @RequestParam("language") String targetLanguage
////) {
////    try {
////        // Save the uploaded PDF file to a temporary location
////        File tempFile = File.createTempFile("uploaded", ".pdf");
////        file.transferTo(tempFile);
////
////        // Extract text from the PDF
////        String text = aiService.extractTextFromPdf(tempFile);
////
////        // Translate the extracted text
////        String translatedText = aiService.translateText(text, targetLanguage);
////
////        // Clean up the temporary file
////        tempFile.delete();
////
////        return ResponseEntity.ok(translatedText);
////    } catch (IOException e) {
////        return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing the PDF file.");
////    }
////}
//
//    ResponseEntity<?> generateQuestions(@RequestParam MultipartFile file){
//        aiService.generateQuestions(file);
//        return null;
//    }
//
//
//}