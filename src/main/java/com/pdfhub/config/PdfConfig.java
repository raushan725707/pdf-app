package com.pdfhub.config;


import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Component
public class PdfConfig {




    public static byte[] readBytes(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        byte[] bytes = new byte[(int) file.length()];
        fis.read(bytes);
        fis.close();
        return bytes;
    }


    public static File createZipArchive(File sourceDir) throws IOException {
        File zipFile = File.createTempFile("split_pages", ".zip");
        try (FileOutputStream fos = new FileOutputStream(zipFile);
             ZipOutputStream zipOut = new ZipOutputStream(fos)) {

            for (File file : sourceDir.listFiles()) {
                try (FileInputStream fis = new FileInputStream(file)) {
                    ZipEntry zipEntry = new ZipEntry(file.getName());
                    zipOut.putNextEntry(zipEntry);

                    byte[] bytes = new byte[1024];
                    int length;
                    while ((length = fis.read(bytes)) >= 0) {
                        zipOut.write(bytes, 0, length);
                    }

                    zipOut.closeEntry();
                }
            }
        }
        return zipFile;
    }




    public static File convertMultipartFileToFile(MultipartFile file) throws IOException{
        File f = new File(file.getOriginalFilename());
        FileOutputStream fos = new FileOutputStream(f);
        fos.write(file.getBytes());
        fos.close();
        return f;
    }




//    public static File createTempDir() throws IOException {
//        File tempDir = File.createTempFile("split_pdf_temp", Long.toString(System.nanoTime()));
//        tempDir.delete();
//        tempDir.mkdir();
//        return tempDir;
//    }

    public static File createTempDir() {
        File tempDir = null;
        try {
            tempDir = File.createTempFile("split_pdf_temp", Long.toString(System.nanoTime()));
            if (tempDir.delete() && tempDir.mkdir()) {
                return tempDir;
            } else {
                throw new IOException("Failed to create directory " + tempDir.getAbsolutePath());
            }
        } catch (IOException e) {
            e.printStackTrace();
            // Handle the exception, e.g., by returning null or throwing a runtime exception
            // throw new RuntimeException("Unable to create temporary directory", e);
        }
        return tempDir; // or return null if appropriate
    }



    public static List<byte[]> convertFilesToBytes(MultipartFile[] files)  {
        return List.of(files).stream()
                .map(file -> {
                    try {
                        return file.getBytes();
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                })
                .toList();
    }



}