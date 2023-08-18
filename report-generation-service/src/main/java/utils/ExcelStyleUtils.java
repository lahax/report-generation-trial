package utils;

import org.apache.poi.ss.usermodel.*;

class ExcelStyleUtils extends ReportGenerator {
    static CellStyle createTitleStyle(Workbook workbook, Font whiteBoldFont) {
        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(whiteBoldFont);
        titleStyle.setFillForegroundColor(IndexedColors.DARK_TEAL.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleStyle.setBorderTop(BorderStyle.MEDIUM);
        titleStyle.setBorderRight(BorderStyle.MEDIUM);
        titleStyle.setBorderLeft(BorderStyle.MEDIUM);
        return titleStyle;
    }
    static CellStyle createWhiteStyle(Workbook workbook) {
        CellStyle whiteStyle = workbook.createCellStyle();
        whiteStyle.setLocked(false);
        whiteStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return whiteStyle;
    }
    static CellStyle createWhiteBorderedStyle(Workbook workbook, Font blackFont) {
        CellStyle whiteBorderedStyle = workbook.createCellStyle();
        whiteBorderedStyle.setLocked(false);
        whiteBorderedStyle.setFont(blackFont);
        whiteBorderedStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteBorderedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        whiteBorderedStyle.setBorderBottom(BorderStyle.MEDIUM);
        whiteBorderedStyle.setBorderTop(BorderStyle.MEDIUM);
        whiteBorderedStyle.setBorderRight(BorderStyle.MEDIUM);
        whiteBorderedStyle.setBorderLeft(BorderStyle.MEDIUM);
        whiteBorderedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        whiteBorderedStyle.setAlignment(HorizontalAlignment.CENTER);
        //whiteBorderedStyle.setShrinkToFit(true);
        return whiteBorderedStyle;
    }
    static CellStyle createPercentageStyle(Workbook workbook, Font blackFont) {
        CellStyle percentageStyle = workbook.createCellStyle();
        DataFormat dataFormat = workbook.createDataFormat();
        percentageStyle.setDataFormat(dataFormat.getFormat("0.00\\%"));
        percentageStyle.setFont(blackFont);
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
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(blackBoldFont);
        headerStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setBorderBottom(BorderStyle.MEDIUM);
        headerStyle.setBorderTop(BorderStyle.MEDIUM);
        headerStyle.setBorderRight(BorderStyle.MEDIUM);
        headerStyle.setBorderLeft(BorderStyle.MEDIUM);
        return headerStyle;
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
    static CellStyle createOkDelayedStyle(Workbook workbook, Font blackFont) {
        CellStyle okDelayedStyle = workbook.createCellStyle();
        okDelayedStyle.setFont(blackFont);
        okDelayedStyle.setFillForegroundColor(IndexedColors.GOLD.getIndex());
        okDelayedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        okDelayedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        okDelayedStyle.setAlignment(HorizontalAlignment.CENTER);
        okDelayedStyle.setBorderBottom(BorderStyle.MEDIUM);
        okDelayedStyle.setBorderTop(BorderStyle.MEDIUM);
        okDelayedStyle.setBorderRight(BorderStyle.MEDIUM);
        okDelayedStyle.setBorderLeft(BorderStyle.MEDIUM);
        return okDelayedStyle;
    }
    static CellStyle createPassedStyle(Workbook workbook, Font greenFont) {
        CellStyle passedStyle = workbook.createCellStyle();
        passedStyle.setFont(greenFont);
        passedStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        passedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        passedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        passedStyle.setAlignment(HorizontalAlignment.CENTER);
        passedStyle.setBorderBottom(BorderStyle.MEDIUM);
        passedStyle.setBorderTop(BorderStyle.MEDIUM);
        passedStyle.setBorderRight(BorderStyle.MEDIUM);
        passedStyle.setBorderLeft(BorderStyle.MEDIUM);
        return passedStyle;
    }
    static CellStyle createFailedStyle(Workbook workbook, Font redFont) {
        CellStyle failedStyle = workbook.createCellStyle();
        failedStyle.setFont(redFont);
        failedStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        failedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        failedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        failedStyle.setAlignment(HorizontalAlignment.CENTER);
        failedStyle.setBorderBottom(BorderStyle.MEDIUM);
        failedStyle.setBorderTop(BorderStyle.MEDIUM);
        failedStyle.setBorderRight(BorderStyle.MEDIUM);
        failedStyle.setBorderLeft(BorderStyle.MEDIUM);
        return failedStyle;
    }
    static CellStyle createDelayedStyle(Workbook workbook, Font yellowFont) {
        CellStyle delayedStyle = workbook.createCellStyle();
        delayedStyle.setFont(yellowFont);
        delayedStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        delayedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        delayedStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        delayedStyle.setAlignment(HorizontalAlignment.CENTER);
        delayedStyle.setBorderBottom(BorderStyle.MEDIUM);
        delayedStyle.setBorderTop(BorderStyle.MEDIUM);
        delayedStyle.setBorderRight(BorderStyle.MEDIUM);
        delayedStyle.setBorderLeft(BorderStyle.MEDIUM);
        return delayedStyle;
    }
    static CellStyle createLightBlueStyle(Workbook workbook) {
        CellStyle lightBlueStyle = workbook.createCellStyle();
        lightBlueStyle.setLocked(false);
        lightBlueStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        lightBlueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return lightBlueStyle;
    }
    static CellStyle createDarkBlueStyle(Workbook workbook) {
        CellStyle darkBlueStyle = workbook.createCellStyle();
        darkBlueStyle.setLocked(false);
        darkBlueStyle.setFillForegroundColor(IndexedColors.DARK_TEAL.getIndex());
        darkBlueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return darkBlueStyle;
    }
    static CellStyle createDetailOkStyle(Workbook workbook, Font blackFont) {
        CellStyle detailOkStyle = workbook.createCellStyle();
        detailOkStyle.setLocked(false);
        detailOkStyle.setFont(blackFont);
        detailOkStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        detailOkStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        detailOkStyle.setBorderBottom(BorderStyle.MEDIUM);
        detailOkStyle.setBorderTop(BorderStyle.MEDIUM);
        detailOkStyle.setBorderRight(BorderStyle.MEDIUM);
        detailOkStyle.setBorderLeft(BorderStyle.MEDIUM);
        detailOkStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        detailOkStyle.setAlignment(HorizontalAlignment.CENTER);
        return detailOkStyle;
    }
    static CellStyle createHeaderCollaudoStyle(Workbook workbook, Font blackBoldFont) {
        CellStyle headerCollaudoStyle = workbook.createCellStyle();
        headerCollaudoStyle.setFont(blackBoldFont);
        headerCollaudoStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        headerCollaudoStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCollaudoStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCollaudoStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCollaudoStyle.setBorderBottom(BorderStyle.THIN);
        headerCollaudoStyle.setBorderTop(BorderStyle.THIN);
        headerCollaudoStyle.setBorderRight(BorderStyle.THIN);
        headerCollaudoStyle.setBorderLeft(BorderStyle.THIN);
        return headerCollaudoStyle;
    }
    static CellStyle createWhiteBorderedCollaudoStyle(Workbook workbook) {
        CellStyle whiteBorderedCollaudoStyle = workbook.createCellStyle();
        whiteBorderedCollaudoStyle.setLocked(false);
        whiteBorderedCollaudoStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        whiteBorderedCollaudoStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        whiteBorderedCollaudoStyle.setBorderBottom(BorderStyle.THIN);
        whiteBorderedCollaudoStyle.setBorderTop(BorderStyle.THIN);
        whiteBorderedCollaudoStyle.setBorderRight(BorderStyle.THIN);
        whiteBorderedCollaudoStyle.setBorderLeft(BorderStyle.THIN);
        whiteBorderedCollaudoStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        whiteBorderedCollaudoStyle.setAlignment(HorizontalAlignment.CENTER);
        return whiteBorderedCollaudoStyle;
    }
    static CellStyle createOkStyleCollaudo(Workbook workbook, Font greenFont) {
        CellStyle okStyleCollaudo = workbook.createCellStyle();
        okStyleCollaudo.setFont(greenFont);
        okStyleCollaudo.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        okStyleCollaudo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        okStyleCollaudo.setVerticalAlignment(VerticalAlignment.CENTER);
        okStyleCollaudo.setAlignment(HorizontalAlignment.CENTER);
        okStyleCollaudo.setBorderBottom(BorderStyle.THIN);
        okStyleCollaudo.setBorderTop(BorderStyle.THIN);
        okStyleCollaudo.setBorderRight(BorderStyle.THIN);
        okStyleCollaudo.setBorderLeft(BorderStyle.THIN);
        return okStyleCollaudo;
    }
    static CellStyle createKoStyleCollaudo(Workbook workbook, Font redFont) {
        CellStyle koStyleCollaudo = workbook.createCellStyle();
        koStyleCollaudo.setFont(redFont);
        koStyleCollaudo.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        koStyleCollaudo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        koStyleCollaudo.setVerticalAlignment(VerticalAlignment.CENTER);
        koStyleCollaudo.setAlignment(HorizontalAlignment.CENTER);
        koStyleCollaudo.setBorderBottom(BorderStyle.THIN);
        koStyleCollaudo.setBorderTop(BorderStyle.THIN);
        koStyleCollaudo.setBorderRight(BorderStyle.THIN);
        koStyleCollaudo.setBorderLeft(BorderStyle.THIN);
        return koStyleCollaudo;
    }
    static Font createWhiteBoldFont(Workbook workbook) {
        Font whiteBoldFont = workbook.createFont();
        whiteBoldFont.setFontHeightInPoints((short) 12);
        whiteBoldFont.setFontName("Arial");
        whiteBoldFont.setColor(IndexedColors.WHITE.getIndex());
        whiteBoldFont.setBold(true);
        return whiteBoldFont;
    }
    static Font createWhiteBoldFont16(Workbook workbook) {
        Font whiteBoldFont16 = workbook.createFont();
        whiteBoldFont16.setFontHeightInPoints((short) 16);
        whiteBoldFont16.setFontName("Arial");
        whiteBoldFont16.setColor(IndexedColors.WHITE.getIndex());
        whiteBoldFont16.setBold(true);
        return whiteBoldFont16;
    }
    static Font createBlackBoldFont(Workbook workbook) {
        Font blackBoldFont = workbook.createFont();
        blackBoldFont.setFontHeightInPoints((short) 10);
        blackBoldFont.setFontName("Arial");
        blackBoldFont.setColor(IndexedColors.BLACK.getIndex());
        blackBoldFont.setBold(true);
        return blackBoldFont;
    }
    static Font createWhiteFont(Workbook workbook) {
        Font whiteFont = workbook.createFont();
        whiteFont.setFontHeightInPoints((short) 10);
        whiteFont.setFontName("Arial");
        whiteFont.setColor(IndexedColors.WHITE.getIndex());
        whiteFont.setBold(true);
        return whiteFont;
    }
    static Font createBlackFont(Workbook workbook){
        Font blackFont = workbook.createFont();
        blackFont.setFontHeightInPoints((short) 10);
        blackFont.setFontName("Arial");
        blackFont.setColor(IndexedColors.BLACK.getIndex());
        return blackFont;
    }
    static Font createGreenFont(Workbook workbook){
        Font greenFont = workbook.createFont();
        greenFont.setFontHeightInPoints((short) 10);
        greenFont.setFontName("Arial");
        greenFont.setColor(IndexedColors.GREEN.getIndex());
        greenFont.setBold(true);
        return greenFont;
    }
    static Font createYellowFont(Workbook workbook){
        Font yellowFont = workbook.createFont();
        yellowFont.setFontHeightInPoints((short) 10);
        yellowFont.setFontName("Arial");
        yellowFont.setColor(IndexedColors.GOLD.getIndex());
        yellowFont.setBold(true);
        return yellowFont;
    }
    static Font createRedFont(Workbook workbook){
        Font redFont = workbook.createFont();
        redFont.setFontHeightInPoints((short) 10);
        redFont.setFontName("Arial");
        redFont.setColor(IndexedColors.DARK_RED.getIndex());
        redFont.setBold(true);
        return redFont;
    }
}