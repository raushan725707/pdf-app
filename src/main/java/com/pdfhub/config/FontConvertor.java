package com.pdfhub.config;

import org.springframework.stereotype.Component;
import org.apache.poi.xssf.usermodel.XSSFFont;
import com.itextpdf.text.Font;
import com.itextpdf.text.FontFactory;

@Component
public class FontConvertor {

    public static Font convertXSSFFontToITextFont(XSSFFont poiFont) {
        // Get the font properties from XSSFFont
        String fontName = poiFont.getFontName();
        float fontSize = poiFont.getFontHeightInPoints();
        boolean isBold = poiFont.getBold();
        boolean isItalic = poiFont.getItalic();
        boolean isStrikeout = poiFont.getStrikeout();

        // Create an iText Font object based on the properties from XSSFFont
        Font iTextFont = FontFactory.getFont(fontName, fontSize);

        // Set bold, italic, and strikeout based on XSSFFont properties
        int style = Font.NORMAL;
        if (isBold) style |= Font.BOLD;
        if (isItalic) style |= Font.ITALIC;
        if (isStrikeout) style |= Font.STRIKETHRU;

        iTextFont.setStyle(style);

        return iTextFont;
    }
}
