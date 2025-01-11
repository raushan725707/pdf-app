package com.pdfhub.service;




import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.pdf.xobject.PdfImageXObject;
import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.pdf.*;
import com.itextpdf.text.pdf.PdfDocument;
import com.itextpdf.text.pdf.PdfWriter;
import com.pdfhub.PPTTOPDF;
import com.pdfhub.config.FontConvertor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

import com.itextpdf.text.Document;
import org.apache.poi.ss.usermodel.Cell;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.springframework.stereotype.Service;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
;
import java.util.Iterator;
import java.util.List;

import org.springframework.web.multipart.MultipartFile;

import javax.imageio.ImageIO;
import java.io.ByteArrayOutputStream;
import java.io.IOException;


@Service
public class ConvertToPdfService {

	public byte[] convertImagesToPdf(List<byte[]> imageBytesList, String orientation, String pageSizeType) throws IOException {
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		try (PDDocument document = new PDDocument()) {
			PDRectangle pageSize = getPageSize(pageSizeType, orientation);

			for (byte[] imageBytes : imageBytesList) {
				PDPage page = new PDPage(pageSize);
				document.addPage(page);
				PDImageXObject pdImage = PDImageXObject.createFromByteArray(document, imageBytes, null);

				try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
					float scale = Math.min(pageSize.getWidth() / pdImage.getWidth(), pageSize.getHeight() / pdImage.getHeight());
					float x = (pageSize.getWidth() - (pdImage.getWidth() * scale)) / 2;
					float y = (pageSize.getHeight() - (pdImage.getHeight() * scale)) / 2;
					contentStream.drawImage(pdImage, x, y, pdImage.getWidth() * scale, pdImage.getHeight() * scale);
				}
			}

			document.save(outputStream);
		}

		return outputStream.toByteArray();
	}

	private PDRectangle getPageSize(String pageSizeType, String orientation) {
		PDRectangle pageSize;
		switch (pageSizeType.toUpperCase()) {
			case "A4":
				pageSize = PDRectangle.A4;
				break;
			case "USLETTER":
				pageSize = PDRectangle.LETTER;
				break;
			case "SAME_SIZE":
				pageSize = PDRectangle.LETTER; // Assuming LETTER as default, since PDFBox doesn't have a DEFAULT page size
				break;
			default:
				pageSize = PDRectangle.A4; // Default to A4 if not specified
				break;
		}

		if ("LANDSCAPE".equalsIgnoreCase(orientation)) {
			pageSize = new PDRectangle(pageSize.getHeight(), pageSize.getWidth());
		}

		return pageSize;
	}

	public byte[] convertJpgToPdf(MultipartFile file) throws IOException {
		// Create a new PDF document
		PDDocument document = new PDDocument();

		// Create a page for the PDF
		PDPage page = new PDPage();
		document.addPage(page);

		// Load the image as a PDF object using the file's byte array
		PDImageXObject pdImage = PDImageXObject.createFromByteArray(document, file.getBytes(), file.getOriginalFilename());

		// Create content stream to write image to PDF
		PDPageContentStream contentStream = new PDPageContentStream(document, page);

		// Get image dimensions and position
		float x = 100; // x-position
		float y = 100; // y-position
		float width = pdImage.getWidth();
		float height = pdImage.getHeight();

		// Draw image on the page
		contentStream.drawImage(pdImage, x, y, width, height);

		// Close the content stream
		contentStream.close();

		// Save the PDF to a byte array
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
		document.save(byteArrayOutputStream);
		document.close();

		// Return the PDF as a byte array
		return byteArrayOutputStream.toByteArray();
	}

	public List<byte[]> convertWordToPdf(List<MultipartFile> files) throws IOException, Docx4JException {
		List<byte[]> pdfList = new ArrayList<>();

		for (MultipartFile file : files) {
			// Process each uploaded Word file (Word format: .docx)
			if (file.getOriginalFilename().endsWith(".docx")) {
				pdfList.add(convertSingleWordToPdf(file));
			} else {
				throw new IOException("Unsupported file type: " + file.getOriginalFilename());
			}
		}

		return pdfList;
	}

	public byte[] convertSingleWordToPdf(MultipartFile file) throws IOException, Docx4JException {
		// Load the Word document (.docx)
		InputStream inputStream = file.getInputStream();
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inputStream);

		// Create a ByteArrayOutputStream to hold the generated PDF
		ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

		// Convert the Word document to PDF using XSLFOPdfConversion
	//	PdfConversion converter = new XSLFOPdfConversion(wordMLPackage);
		//converter.convert(byteArrayOutputStream, null);  // Pass a null file for no output to disk

		// Return the PDF byte array
		return byteArrayOutputStream.toByteArray();
	}

	public byte[] convertExcelToPdf(MultipartFile file) throws IOException, DocumentException, DocumentException {
		// Load the Excel file using Apache POI
		Workbook workbook = new XSSFWorkbook(file.getInputStream());
		ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();

		// Create a PDF document using iText
		PdfDocument pdfDocument = new PdfDocument();
		Document document = new Document();

		// Iterate over each sheet in the Excel file
		for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			Iterator<Row> rowIterator = sheet.iterator();

			// Create a table in the PDF to hold Excel data
			Table pdfTable = new Table(sheet.getRow(0).getPhysicalNumberOfCells());

			// Iterate over each row in the sheet
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.iterator();

				// Iterate over each cell in the row and convert it to iText Cell
				while (cellIterator.hasNext()) {
					Cell cell = (Cell) cellIterator.next();
					String cellText = cell.toString(); // Convert POI Cell to String

					// Create an iText Cell and add it to the Table
					pdfTable.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(cellText)));
				}
			}

			// Add the table to the PDF document
			document.add((Element) pdfTable);
		}

		// Close the document and return the PDF as a byte array
		document.close();
		return pdfOutputStream.toByteArray();
	}

















//	public byte[] convertExcelToPdf2(MultipartFile file) throws IOException, DocumentException {
//		Workbook workbook = new XSSFWorkbook(file.getInputStream());  // Read the Excel file
//		ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
//
//		// Create a PDF document using iText
//		Document document = new Document();
//		PdfWriter.getInstance(document, pdfOutputStream);
//		document.open();
//
//		// Iterate over all sheets in the Excel file
//		for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
//			Sheet sheet = workbook.getSheetAt(sheetIndex);
//			PdfPTable table = new PdfPTable(sheet.getRow(0).getPhysicalNumberOfCells());
//
//			// Iterate over rows in the sheet
//			for (Row row : sheet) {
//				// Skip the header row if necessary
//				if (row.getRowNum() == 0) continue;
//
//				// Iterate over cells in each row
//				for (
//						org.apache.poi.ss.usermodel.Cell cell : row) {
//					String cellValue = getCellValue(cell);
//					PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue));
//
//					// Apply styling like font, background color, alignment, etc.
//					applyCellStyle(cell, pdfCell);
//
//					// Add the cell to the table
//					table.addCell(pdfCell);
//				}
//			}
//
//			// Add the table to the document
//			document.add(table);
//		}
//
//		// Close the document
//		document.close();
//		workbook.close();
//
//		return pdfOutputStream.toByteArray();  // Return the generated PDF as a byte array
//	}
//
//	// Get the cell value from Apache POI Cell
//	private String getCellValue(Cell cell) {
//		String cellValue = "";
//		switch (cell.getCellType()) {
//			case STRING:
//				cellValue = cell.getStringCellValue();
//				break;
//			case NUMERIC:
//				cellValue = String.valueOf(BigDecimal.valueOf(cell.getNumericCellValue()));
//				break;
//			case BLANK:
//				break;
//			default:
//				cellValue = "";
//		}
//		return cellValue;
//	}
//
//	// Apply cell style (font, background, and alignment)
//	private void applyCellStyle(Cell cell, PdfPCell pdfCell) {
//		CellStyle cellStyle = cell.getCellStyle();
//
//		// Font styling
//		Font font= FontConvertor.convertXSSFFontToITextFont()
//		Font font = getCellFont(cell);
//		pdfCell.setPhrase(new Phrase(pdfCell.getPhrase().getContent(), font));
//
//		// Background color
//		setCellBackgroundColor(cell, pdfCell);
//
//		// Alignment
//		setCellAlignment(cell, pdfCell);
//	}
//
//	// Get the font styling from Apache POI
//	private Font getCellFont(Cell cell) {
//		Font font = new Font();
//		CellStyle cellStyle = cell.getCellStyle();
//		Font cellFont = (Font) cell.getSheet()
//				.getWorkbook()
//				.getFontAt(cellStyle.getFontIndexAsInt());
//
//		if (cellFont.isItalic()) {
//			font.setStyle(Font.ITALIC);
//		}
//		if ((cellFont.getStyle() & Font.STRIKETHRU) != 0) {
//			font.setStyle(Font.STRIKETHRU);
//		}
//		if ((cellFont.getStyle() & Font.UNDERLINE) != 0) {
//			font.setStyle(Font.UNDERLINE); // Apply underline style if it's set
//		}
//		font.setSize(cellFont.getSize());
//
//		if (cellFont.isBold()) {
//			font.setStyle(Font.BOLD);
//		}
//
//		String fontFamilyName = cellFont.getFamilyname();
//		font.setFamily(fontFamilyName);
//
//		return font;
//	}
//
//	// Set the background color for the PdfPCell
//	private void setCellBackgroundColor(Cell cell, PdfPCell pdfCell) {
//		short bgColorIndex = cell.getCellStyle().getFillForegroundColor();
//		if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
//			XSSFColor bgColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
//			if (bgColor != null) {
//				byte[] rgb = bgColor.getRGB();
//				if (rgb != null && rgb.length == 3) {
//					pdfCell.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
//				}
//			}
//		}
//
//
//
//
//	}
//	void setCellAlignment(Cell cell, PdfPCell cellPdf) {
//		CellStyle cellStyle = cell.getCellStyle();
//
//		HorizontalAlignment horizontalAlignment = cellStyle.getAlignment();
//
//		switch (horizontalAlignment) {
//			case LEFT:
//				cellPdf.setHorizontalAlignment(Element.ALIGN_LEFT);
//				break;
//			case CENTER:
//				cellPdf.setHorizontalAlignment(Element.ALIGN_CENTER);
//				break;
//			case JUSTIFY:
//			case FILL:
//				cellPdf.setVerticalAlignment(Element.ALIGN_JUSTIFIED);
//				break;
//			case RIGHT:
//				cellPdf.setHorizontalAlignment(Element.ALIGN_RIGHT);
//				break;
//		}
//	}


	//working dont touch the code
	public byte[] convertExcelToPdf2(MultipartFile file) throws IOException, DocumentException {
		Workbook workbook = new XSSFWorkbook(file.getInputStream());  // Read the Excel file
		ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();

		// Create a PDF document using iText
		Document document = new Document();
		PdfWriter.getInstance(document, pdfOutputStream);
		document.open();

		// Iterate over all sheets in the Excel file
		for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			PdfPTable table = new PdfPTable(sheet.getRow(0).getPhysicalNumberOfCells());

			// Iterate over rows in the sheet
			for (Row row : sheet) {
				// Skip the header row if necessary
				if (row.getRowNum() == 0) continue;

				// Iterate over cells in each row
				for (Cell cell : row) {
					String cellValue = getCellValue(cell);
					PdfPCell pdfCell = new PdfPCell(new Phrase(cellValue));

					// Apply styling like font, background color, alignment, etc.
					applyCellStyle(cell, pdfCell);

					// Add the cell to the table
					table.addCell(pdfCell);
				}
			}

			// Add the table to the document
			document.add(table);
		}

		// Close the document and workbook
		document.close();
		workbook.close();

		return pdfOutputStream.toByteArray();  // Return the generated PDF as a byte array
	}

	// Get the cell value from Apache POI Cell
	private String getCellValue(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
			case STRING:
				cellValue = cell.getStringCellValue();
				break;
			case NUMERIC:
				cellValue = String.valueOf(BigDecimal.valueOf(cell.getNumericCellValue()));
				break;
			case BLANK:
				break;
			default:
				cellValue = "";
		}
		return cellValue;
	}

	// Apply cell style (font, background, and alignment)
	private void applyCellStyle(Cell cell, PdfPCell pdfCell) {
		CellStyle cellStyle = cell.getCellStyle();

		// Font styling
		Font font = getCellFont(cell);
		pdfCell.setPhrase(new Phrase(pdfCell.getPhrase().getContent(), font));

		// Background color
		setCellBackgroundColor(cell, pdfCell);

		// Alignment
		setCellAlignment(cell, pdfCell);
	}

	// Get the font styling from Apache POI
	private Font getCellFont(Cell cell) {
		Font font = new Font();
		CellStyle cellStyle = cell.getCellStyle();
		XSSFFont cellFont = (XSSFFont) cell.getSheet()
				.getWorkbook()
				.getFontAt(cellStyle.getFontIndexAsInt());

		// Set the font style flags
		int style = Font.NORMAL;
		if (cellFont.getBold()) style |= Font.BOLD;
		if (cellFont.getItalic()) style |= Font.ITALIC;
		if (cellFont.getStrikeout()) style |= Font.STRIKETHRU;
		if ( Font.UNDERLINE!= 0) style |= Font.UNDERLINE;

		font.setStyle(style);

		// Set the font size and family
		font.setSize(cellFont.getFontHeightInPoints());
		font.setFamily(cellFont.getFontName());

		return font;
	}

	// Set the background color for the PdfPCell
	private void setCellBackgroundColor(Cell cell, PdfPCell pdfCell) {
		short bgColorIndex = cell.getCellStyle().getFillForegroundColor();
		if (bgColorIndex != IndexedColors.AUTOMATIC.getIndex()) {
			XSSFColor bgColor = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
			if (bgColor != null) {
				byte[] rgb = bgColor.getRGB();
				if (rgb != null && rgb.length == 3) {
					pdfCell.setBackgroundColor(new BaseColor(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
				}
			}
		}
	}

	// Set cell alignment based on POI cell alignment
	private void setCellAlignment(Cell cell, PdfPCell pdfCell) {
		CellStyle cellStyle = cell.getCellStyle();
		HorizontalAlignment horizontalAlignment = cellStyle.getAlignment();

		switch (horizontalAlignment) {
			case LEFT:
				pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT);
				break;
			case CENTER:
				pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
				break;
			case RIGHT:
				pdfCell.setHorizontalAlignment(Element.ALIGN_RIGHT);
				break;
			default:
				pdfCell.setHorizontalAlignment(Element.ALIGN_LEFT);  // Default to left alignment
		}
	}




	//powerpoint to pdf
//	public byte[] convertPowerPointToPdf(MultipartFile file) throws IOException, DocumentException {
//		XMLSlideShow ppt = new XMLSlideShow(file.getInputStream());
//
//		// Create PDF OutputStream
//		ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
//
//		// Initialize the PDF Document and PdfWriter
//		PdfDocument pdfDoc = new PdfDocument();
//		Document document = new Document(pdfDoc.getPageSize());
//
//		// Iterate over each slide in the PowerPoint
//		for (XSLFSlide slide : ppt.getSlides()) {
//			// Convert each slide to an image (Render as image)
//			ByteArrayOutputStream slideImageStream = new ByteArrayOutputStream();
//			ImageIO.write(renderSlideAsImage(slide), "png", slideImageStream);
//			ByteArrayInputStream imageStream = new ByteArrayInputStream(slideImageStream.toByteArray());
//
//
//			Image pdfImage = new Image(imageStream);
//			document.add(pdfImage);
//
//
//		}
//
//
//		document.close();
//		return pdfOutputStream.toByteArray();
//	}
//
//	// Render slide as an image (can use Apache POI or external tools like Apache Batik)
//	private BufferedImage renderSlideAsImage(XSLFSlide slide) throws IOException {
//		// Use a library like Apache Batik or use Java2D to render the slide.
//		// Here we just get a simple screenshot of the slide content.
//		// In production, you might want a library that renders the slide into an image.
//		Dimension pageSize = slide.getSlideShow().getPageSize();
//		BufferedImage img = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_ARGB);
//		Graphics2D graphics = img.createGraphics();
//		slide.draw(graphics);
//		graphics.dispose();
//		return img;
//	}



//	public byte[] convertPowerPointToPdf(MultipartFile file) throws IOException, DocumentException {
//		// Initialize POI XMLSlideShow for PowerPoint
//		XMLSlideShow ppt = new XMLSlideShow(file.getInputStream());
//
//		// Create PDF OutputStream
//		ByteArrayOutputStream pdfOutputStream = new ByteArrayOutputStream();
//
//		// Initialize the PDF Document and PdfWriter
//		PdfDocument pdfDoc = new PdfDocument();
//		Document document = new Document(pdfDoc);
//
//		// Iterate over each slide in the PowerPoint
//		for (XSLFSlide slide : ppt.getSlides()) {
//			// Convert each slide to an image (Render as image)
//			BufferedImage slideImage = renderSlideAsImage(slide);
//
//			// Convert the BufferedImage to a ByteArrayOutputStream
//			ByteArrayOutputStream slideImageStream = new ByteArrayOutputStream();
//			ImageIO.write(slideImage, "png", slideImageStream);
//
//			// Create ImageData from byte array
//			byte[] imageByteArray = slideImageStream.toByteArray();
//			com.itextpdf.io.image.ImageData imageData = ImageDataFactory.create(imageByteArray);
//
//			// Create Image object from ImageData
//			Image pdfImage = new Image(imageData);
//
//			// Add the image to the PDF document
//			document.add(pdfImage);
//		}
//
//		// Close the document
//		document.close();
//		return pdfOutputStream.toByteArray(); // Return the PDF as byte array
//	}
//
//
//	// Render slide as an image (Using Graphics2D)
//	private BufferedImage renderSlideAsImage(XSLFSlide slide) throws IOException {
//		Dimension pageSize = slide.getSlideShow().getPageSize();
//		BufferedImage img = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_ARGB);
//		Graphics2D graphics = img.createGraphics();
//		slide.draw(graphics);  // Draw the slide on the graphics object (this renders the slide content)
//		graphics.dispose();
//		return img;
//	}

	public byte[] convertPowerPointToPdf(MultipartFile file) throws Exception {
		try (InputStream inputStream = file.getInputStream();
			 ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

			// Create the PPTTOPDF converter with the input and output streams
			PPTTOPDF pptToPdfConverter = new PPTTOPDF(inputStream, outputStream);

			// Convert PPT to PDF
			pptToPdfConverter.convertPPTTOPDF();

			// Return the generated PDF as a byte array
			return outputStream.toByteArray();
		}
	}


}