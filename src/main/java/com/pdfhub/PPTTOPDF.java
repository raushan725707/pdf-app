package com.pdfhub;


import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.List;


public class PPTTOPDF {

    private InputStream inStream;
    private OutputStream outStream;
    private XSLFSlide[] slides;

    // Constructor that accepts InputStream (for PPT) and OutputStream (for PDF)
    public PPTTOPDF(InputStream inStream, OutputStream outStream) {
        this.inStream = inStream;
        this.outStream = outStream;
    }

    // Main method to convert PPT to PDF
    public void convertPPTTOPDF() throws Exception {
        // Process the slides and get their dimensions
        Dimension pgSize = processSlides();

        // Set zoom level to magnify the slides for better resolution
        double zoom = 2;
        AffineTransform at = new AffineTransform();
        at.setToScale(zoom, zoom);

        // Create a new PDF document for the output
        Document document = new Document();
        PdfWriter writer = PdfWriter.getInstance(document, outStream);
        document.open();

        // Loop through the slides and generate a PDF image for each slide
        for (int i = 0; i < slides.length; i++) {
            BufferedImage bufImg = new BufferedImage(
                    (int) Math.ceil(pgSize.width * zoom),
                    (int) Math.ceil(pgSize.height * zoom),
                    BufferedImage.TYPE_INT_RGB);

            Graphics2D graphics = bufImg.createGraphics();
            graphics.setTransform(at);
            graphics.setPaint(getSlideBGColor(i));
            graphics.fill(new Rectangle2D.Float(0, 0, pgSize.width, pgSize.height));

            // Try to draw the slide content on the image
            try {
                drawOntoThisGraphic(i, graphics);
            } catch (Exception e) {
                e.printStackTrace(); // If drawing fails, proceed with the next slide
            }

            // Convert BufferedImage to iText Image and add to the PDF document
            Image image = Image.getInstance(bufImg, null);
            document.setPageSize(new Rectangle(image.getScaledWidth(), image.getScaledHeight()));
            document.newPage();
            image.setAbsolutePosition(0, 0);
            document.add(image);
        }

        // Close the document and the PDF writer
        document.close();
        writer.close();
    }

    // Process the slides and return their dimensions
    private Dimension processSlides() throws IOException {
        XMLSlideShow ppt = new XMLSlideShow(inStream);
        slides = ppt.getSlides().toArray(new XSLFSlide[0]);
        return ppt.getPageSize();
    }

    // Draw the content of a slide onto the Graphics2D object
    private void drawOntoThisGraphic(int index, Graphics2D graphics) {
        slides[index].draw(graphics);
    }

    // Get the background color for a slide
    private Color getSlideBGColor(int index) {
        return slides[index].getBackground().getFillColor();
    }
}




//package com.pdfhub;
//
//import com.itextpdf.text.Image;
//import com.itextpdf.text.Rectangle;
//import org.apache.poi.xslf.usermodel.XMLSlideShow;
//import com.itextpdf.text.*;
//import com.itextpdf.text.pdf.PdfWriter;
//import org.apache.poi.xslf.usermodel.XSLFSlide;
//import java.awt.*;
//import java.awt.geom.AffineTransform;
//import java.awt.geom.Rectangle2D;
//import java.awt.image.BufferedImage;
//import java.io.*;
//import java.util.List;
//
//public class PPTTOPDF {
//
//    private InputStream inStream;
//    private OutputStream outStream;
//    private XSLFSlide[] slides;
//
//    // Constructor that accepts InputStream and OutputStream
//    public PPTTOPDF(InputStream inStream, OutputStream outStream) {
//        this.inStream = inStream;
//        this.outStream = outStream;
//    }
//
//    // Main method to convert PPT to PDF
//    public void convertPPTTOPDF(InputStream inStream, OutputStream outStream) throws Exception {
//        // Process the slides and get their dimensions
//        Dimension pgSize = processSlides();
//
//        // Set zoom level to magnify the slides
//        double zoom = 2; // magnify by 2 to improve resolution
//        AffineTransform at = new AffineTransform();
//        at.setToScale(zoom, zoom);
//
//        // Create a new PDF document for the output
//        Document document = new Document();
//        PdfWriter writer = PdfWriter.getInstance(document, outStream);
//        document.open();
//
//        // Loop through the slides and generate a PDF image for each slide
//        for (int i = 0; i < slides.length; i++) {
//            BufferedImage bufImg = new BufferedImage(
//                    (int) Math.ceil(pgSize.width * zoom),
//                    (int) Math.ceil(pgSize.height * zoom),
//                    BufferedImage.TYPE_INT_RGB);
//
//            Graphics2D graphics = bufImg.createGraphics();
//            graphics.setTransform(at);
//            graphics.setPaint(getSlideBGColor(i));
//            graphics.fill(new Rectangle2D.Float(0, 0, pgSize.width, pgSize.height));
//
//            try {
//                drawOntoThisGraphic(i, graphics);
//            } catch (Exception e) {
//                // If drawing fails, proceed with the next slide
//                e.printStackTrace();
//            }
//
//            // Convert BufferedImage to iText Image and add to the PDF document
//            Image image = Image.getInstance(bufImg, null);
//            document.setPageSize(new Rectangle(image.getScaledWidth(), image.getScaledHeight()));
//            document.newPage();
//            image.setAbsolutePosition(0, 0);
//            document.add(image);
//        }
//
//        // Close the document and the PDF writer
//        document.close();
//        writer.close();
//    }
//
//    // Process the slides and return their dimension
//    private Dimension processSlides() throws IOException {
//        XMLSlideShow ppt = new XMLSlideShow(inStream);
//        slides = ppt.getSlides().toArray(new XSLFSlide[0]);
//        return ppt.getPageSize();
//    }
//
//    // Return the number of slides
//    private int getNumSlides() {
//        return slides.length;
//    }
//
//    // Draw the content of a slide onto the Graphics2D object
//    private void drawOntoThisGraphic(int index, Graphics2D graphics) {
//        slides[index].draw(graphics);
//    }
//
//    // Get the background color for a slide
//    private Color getSlideBGColor(int index) {
//        return slides[index].getBackground().getFillColor();
//    }
//}
