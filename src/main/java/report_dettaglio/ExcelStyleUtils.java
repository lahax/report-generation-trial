package report_dettaglio;

import org.apache.poi.ss.usermodel.*;

class ExcelStyleUtils extends ReportDettaglio{
    static CellStyle createTitleStyle(Workbook workbook, Font whiteBoldFont) {
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(whiteBoldFont);
        headerStyle.setFillForegroundColor(IndexedColors.DARK_TEAL.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        return headerStyle;
    }

    static CellStyle createWhiteStyle(Workbook workbook) {
        CellStyle whiteStyle = workbook.createCellStyle();
        whiteStyle.setLocked(false);
        whiteStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return whiteStyle;
    }

    static CellStyle createWhiteBorderedStyle(Workbook workbook) {
        CellStyle whiteBorderedStyle = workbook.createCellStyle();
        whiteBorderedStyle.setLocked(false);
        whiteBorderedStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteBorderedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        whiteBorderedStyle.setBorderBottom(BorderStyle.MEDIUM);
        whiteBorderedStyle.setBorderTop(BorderStyle.MEDIUM);
        whiteBorderedStyle.setBorderRight(BorderStyle.MEDIUM);
        whiteBorderedStyle.setBorderLeft(BorderStyle.MEDIUM);
        whiteBorderedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        whiteBorderedStyle.setAlignment(HorizontalAlignment.CENTER);
        return whiteBorderedStyle;
    }

    static CellStyle createPercentageStyle(Workbook workbook) {
        CellStyle percentageStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        percentageStyle.setDataFormat(dataFormat.getFormat("0.00\\%"));
        percentageStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        percentageStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        percentageStyle.setBorderBottom(BorderStyle.MEDIUM);
        percentageStyle.setBorderTop(BorderStyle.MEDIUM);
        percentageStyle.setBorderRight(BorderStyle.MEDIUM);
        percentageStyle.setBorderLeft(BorderStyle.MEDIUM);
        percentageStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        percentageStyle.setAlignment(HorizontalAlignment.CENTER);
        return percentageStyle;
    }

    static CellStyle createHeaderStyle(Workbook workbook, Font blackBoldFont) {
        CellStyle lightBlueStyle = workbook.createCellStyle();
        lightBlueStyle.setFont(blackBoldFont);
        lightBlueStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        lightBlueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        lightBlueStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        lightBlueStyle.setAlignment(HorizontalAlignment.CENTER);
        lightBlueStyle.setBorderBottom(BorderStyle.MEDIUM);
        lightBlueStyle.setBorderTop(BorderStyle.MEDIUM);
        lightBlueStyle.setBorderRight(BorderStyle.MEDIUM);
        lightBlueStyle.setBorderLeft(BorderStyle.MEDIUM);
        return lightBlueStyle;
    }

    static CellStyle createOkStyle(Workbook workbook, Font whiteFont) {
        CellStyle okStyle = workbook.createCellStyle();
        okStyle.setFont(whiteFont);
        okStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        okStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        okStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        okStyle.setAlignment(HorizontalAlignment.CENTER);
        okStyle.setBorderBottom(BorderStyle.MEDIUM);
        okStyle.setBorderTop(BorderStyle.MEDIUM);
        okStyle.setBorderRight(BorderStyle.MEDIUM);
        okStyle.setBorderLeft(BorderStyle.MEDIUM);
        return okStyle;
    }

    static CellStyle createKoStyle(Workbook workbook, Font whiteFont) {
        CellStyle koStyle = workbook.createCellStyle();
        koStyle.setFont(whiteFont);
        koStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        koStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        koStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        koStyle.setAlignment(HorizontalAlignment.CENTER);
        koStyle.setBorderBottom(BorderStyle.MEDIUM);
        koStyle.setBorderTop(BorderStyle.MEDIUM);
        koStyle.setBorderRight(BorderStyle.MEDIUM);
        koStyle.setBorderLeft(BorderStyle.MEDIUM);
        return koStyle;
    }

    static Font createWhiteBoldFont(Workbook workbook) {
        Font whiteBoldFont = workbook.createFont();
        whiteBoldFont.setFontHeightInPoints((short) 16);
        whiteBoldFont.setFontName("Arial");
        whiteBoldFont.setColor(IndexedColors.WHITE.getIndex());
        whiteBoldFont.setBold(true);
        return whiteBoldFont;
    }

    static Font createBlackBoldFont(Workbook workbook) {
        Font blackBoldFont = workbook.createFont();
        blackBoldFont.setFontHeightInPoints((short) 10);
        blackBoldFont.setFontName("Arial");
        blackBoldFont.setColor(IndexedColors.BLACK.getIndex());
        blackBoldFont.setBold(true);
        return blackBoldFont;
    }

    static Font createWhiteFont(Workbook workbook){
        Font whiteFont = workbook.createFont();
        whiteFont.setFontName("Arial");
        whiteFont.setColor(IndexedColors.WHITE.getIndex());
        return whiteFont;
    }

}
