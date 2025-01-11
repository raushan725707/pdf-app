package com.pdfhub.controller;



import com.pdfhub.config.PdfConfig;
import com.pdfhub.service.OrganizedPdfService;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDPage;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;


import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import java.util.List;


import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;



@RestController
@CrossOrigin(origins = "http://localhost:3000")
@RequestMapping("/pdf")
public class OrganizePdfController {

	@Autowired
	private OrganizedPdfService organizedPdfService;


@PostMapping("/split/extract/extractallpage")
public ResponseEntity<?>splitExtractExtractAllPage(@RequestParam("file") MultipartFile file) throws  IOException{

  try{  File pdfFile = PdfConfig.convertMultipartFileToFile(file);

    // Load PDF document
    PDDocument document = Loader.loadPDF(pdfFile);
    int totalPages = document.getNumberOfPages();
    File tempDir = PdfConfig.createTempDir();
    for (int pageNum = 0; pageNum < totalPages; pageNum++) {
        // Create a new PDF document
        PDDocument singlePageDocument = new PDDocument();
        singlePageDocument.addPage((PDPage) document.getDocumentCatalog().getPages().get(pageNum));

        // Create a file for the single page
        File singlePageFile = new File(tempDir, "page_" + (pageNum + 1) + ".pdf");
        singlePageDocument.save(singlePageFile);
        singlePageDocument.close();
    }
    File zipFile = PdfConfig.createZipArchive(tempDir);

    // Prepare the response with the ZIP file
    byte[] zipBytes = PdfConfig.readBytes(zipFile);
    HttpHeaders headers = new HttpHeaders();
    headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
    headers.setContentDispositionFormData("attachment", "split_pages.zip");
    headers.setContentLength(zipBytes.length);

    return new ResponseEntity<>(zipBytes, headers, HttpStatus.OK);
}
catch (Exception e) {
        e.printStackTrace();
        return new ResponseEntity<>("Failed to split PDF: " + e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
    }



}

    @PostMapping("/split/extract/extractallpage/pagenumber")
    public ResponseEntity<?>splitExtractExtractAllPage(@RequestParam("file") MultipartFile file,@RequestParam("pageNumber") List<Integer> pageNumber) throws IOException{
        File pdfFile =PdfConfig.convertMultipartFileToFile(file);

        // Load PDF document
        PDDocument document = Loader.loadPDF(pdfFile);
        PDDocument pdf2=new PDDocument();
        PDFMergerUtility merger = new PDFMergerUtility();
        merger.setDestinationStream(new ByteArrayOutputStream());

        for (Integer pagenumber : pageNumber) {
            PDPage page=   document.getPage(pagenumber-1); // PDFBox pages are 0-based index
           // merger.appendDocument(pdf2, document.getPage(pagenumber));
            pdf2.addPage(page);
        }


        merger.appendDocument(pdf2, pdf2);

        // Prepare merged PDF in ByteArrayOutputStream
        ByteArrayOutputStream mergedPdfOutputStream = (ByteArrayOutputStream) merger.getDestinationStream();







       //  Convert ByteArrayOutputStream to byte array
        byte[] mergedPdfBytes = mergedPdfOutputStream.toByteArray();

        // Close documents
        document.close();
        pdf2.close();

        // Return merged PDF as ResponseEntity
        return ResponseEntity
                .ok()
                .header("Content-Disposition", "attachment; filename=merged.pdf")
                .body(mergedPdfBytes);




//        byte[] content = mergedPdfOutputStream.toByteArray();
//        HttpHeaders headers = new HttpHeaders();
//        headers.setContentDisposition(ContentDisposition.attachment()
//                .filename("extractpdfpage.pdf")
//                .build());
//        headers.setContentType(MediaType.APPLICATION_PDF);
//
//        return new ResponseEntity<>(content, headers, HttpStatus.OK);

}

    @PostMapping("/split/extract/extractallpage/customrange")
    public ResponseEntity<?>splitByCustomRange(@RequestParam("file") MultipartFile file) throws  IOException{



return  null;
    }




    @PostMapping("/split/extract/extractallpage/fixedrange")
    public ResponseEntity<?>splitByFixedRange(@RequestParam("file") MultipartFile file) throws  IOException{



        return  null;
    }






    
    
    
    
    
    

    
    @PostMapping("/mergepdfs")
    public ResponseEntity<?> mergePdfss(@RequestParam("files") List<MultipartFile> files) throws IOException {

        if (files.isEmpty() || files.size() < 2) {
            return ResponseEntity.badRequest().body("Please select at least 2 PDF files");
        }

        PDFMergerUtility pdfmerger = new PDFMergerUtility();

        // Using a temporary directory for merged file (optional)
        String tempDir = System.getProperty("java.io.tmpdir");
        String mergedFilePath = tempDir + File.separator + "merged.pdf";
       pdfmerger.setDestinationFileName(mergedFilePath);

        for (MultipartFile file : files) {
            pdfmerger.addSource(PdfConfig.convertMultipartFileToFile(file));
        }



        pdfmerger.mergeDocuments(null);

        byte[] content = Files.readAllBytes(Paths.get(pdfmerger.getDestinationFileName()));
        HttpHeaders headers = new HttpHeaders();
        headers.setContentDisposition(ContentDisposition.attachment()
                .filename("merged.pdf")
                .build());
        headers.setContentType(MediaType.APPLICATION_PDF);

        return new ResponseEntity<>(content, headers, HttpStatus.OK);

    }



    
    
    
    
    
    @PostMapping("/removepage")
    public ResponseEntity<byte[]> removePageFromPdf(@RequestParam("file") MultipartFile file,
                                               @RequestParam("pageNumber") List<Integer> pageNumber) throws IOException {


        File f =PdfConfig.convertMultipartFileToFile(file);
        //  PDDocument document = Loader.loadPDF(f);
        try (PDDocument document = Loader.loadPDF(f)) {

            // Remove pages based on pageNumber list
            pageNumber.forEach(number -> {
                if (number >= 1 && number <= document.getNumberOfPages()) {
                    try {
                        document.removePage(number - 1); // PDFBox pages are 0-indexed
                    } catch (Exception e) {
                        // Handle IOException (if needed)
                        e.printStackTrace();
                    }
                }
            });
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            document.save(byteArrayOutputStream);
            document.close();

            byte[] content = byteArrayOutputStream.toByteArray();
            HttpHeaders headers = new HttpHeaders();
            headers.setContentDisposition(ContentDisposition.attachment()
                    .filename("removepage.pdf")
                    .build());
            headers.setContentType(MediaType.APPLICATION_PDF);

            return new ResponseEntity<>(content, headers, HttpStatus.OK);

        }
    }


    @PostMapping("/organizepdf")
    ResponseEntity<?> organizePdf(@RequestParam("file") MultipartFile file) throws IOException{
    File f=PdfConfig.convertMultipartFileToFile(file);





    return null;
    }
    @PostMapping("/scanpdf")
    ResponseEntity<?> scanPdf(@RequestParam("file") MultipartFile file) throws IOException{
        File f=PdfConfig.convertMultipartFileToFile(file);
        return null;
    }


    }


