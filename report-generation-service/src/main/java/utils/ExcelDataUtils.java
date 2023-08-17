package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import shared.SummaryReportData;
import static utils.ExcelStyleUtils.*;

class ExcelDataUtils {
    private static CellStyle titleStyle;
    private static CellStyle headerStyle;
    private static CellStyle whiteStyle;
    private static CellStyle whiteBorderedStyle;
    private static CellStyle percentageStyle;
    private static CellStyle okStyle;
    private static CellStyle koStyle;

    public static void initializeStyles(Workbook workbook) {
        titleStyle = createTitleStyle(workbook, createWhiteBoldFont(workbook));
        headerStyle = createHeaderStyle(workbook, createBlackBoldFont(workbook));
        whiteStyle = createWhiteStyle(workbook);
        whiteBorderedStyle = createWhiteBorderedStyle(workbook);
        percentageStyle = createPercentageStyle(workbook);
        okStyle = createOkStyle(workbook, createWhiteFont(workbook));
        koStyle = createKoStyle(workbook, createWhiteFont(workbook));
    }

    static void addTitle(XSSFSheet sheet, int rowIndex, String title) {
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue(title);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 3);
        sheet.addMergedRegion(cellRangeAddress);
        row.setRowStyle(whiteStyle);
        cell.setCellStyle(titleStyle);
    }

    static void addHeaderData(XSSFSheet sheet, int rowIndex, String label, Object value) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
        Cell cell = row.createCell(0);
        cell.setCellValue(label);
        cell.setCellStyle(headerStyle);

        cell = row.createCell(1);
        cell.setCellStyle(headerStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 1);
        sheet.addMergedRegion(cellRangeAddress);

        cell = row.createCell(2);
        if (value instanceof String) {
            cell.setCellValue((String) value);
            cell.setCellStyle(whiteBorderedStyle);
        } else if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
            cell.setCellStyle(whiteBorderedStyle);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
            cell.setCellStyle(percentageStyle);
        }

        cell = row.createCell(3);
        cell.setCellStyle(whiteBorderedStyle);
        cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 2, 3);
        sheet.addMergedRegion(cellRangeAddress);
    }

    static void addWhiteLine(XSSFSheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
    }

    static void addDetailsHeader(XSSFSheet sheet, int rowIndex, String title1, String title2, String title3, String title4) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);

        Cell cell = row.createCell(0);
        cell.setCellValue(title1);
        cell.setCellStyle(headerStyle);

        cell = row.createCell(1);
        cell.setCellValue(title2);
        cell.setCellStyle(headerStyle);

        cell = row.createCell(2);
        cell.setCellValue(title3);
        cell.setCellStyle(headerStyle);

        cell = row.createCell(3);
        cell.setCellValue(title4);
        cell.setCellStyle(headerStyle);
    }

    static void addQuantityData(XSSFSheet sheet, int rowIndex, SummaryReportData summaryReportData, double toleranceThreshold) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);

        Cell cell = row.createCell(0);
        cell.setCellValue(summaryReportData.getFieldDataTotal());
        cell.setCellStyle(whiteBorderedStyle);

        cell = row.createCell(1);
        cell.setCellValue(summaryReportData.getDashboardDataTotal());
        cell.setCellStyle(whiteBorderedStyle);

        double recievedValuesPercentage = ((double) summaryReportData.getDashboardDataTotal() / (double) summaryReportData.getFieldDataTotal()) * 100;

        cell = row.createCell(2);
        cell.setCellValue(recievedValuesPercentage);
        cell.setCellStyle(percentageStyle);

        cell = row.createCell(3);
        if (summaryReportData.getIsOk().equals(true)) {
            cell.setCellValue("OK");
            cell.setCellStyle(okStyle);
        } else {
            cell.setCellValue("KO");
            cell.setCellStyle(koStyle);
        }
    }

    static void addRecapDescriptionData(XSSFSheet sheet, int rowIndex, SummaryReportData summaryReportData, double toleranceThreshold) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);

        Cell cell = row.createCell(0);
        cell.setCellValue(summaryReportData.getFieldDataTotal());
        cell.setCellStyle(whiteBorderedStyle);

        cell = row.createCell(1);
        cell.setCellValue(summaryReportData.getDashboardDataTotal().toString() + '/' + summaryReportData.getFieldDataTotal().toString());
        cell.setCellStyle(whiteBorderedStyle);

        double recievedValuesPercentage = ((double) summaryReportData.getDashboardDataTotal() / (double) summaryReportData.getFieldDataTotal()) * 100;

        cell = row.createCell(2);
        cell.setCellValue(recievedValuesPercentage);
        cell.setCellStyle(percentageStyle);

        cell = row.createCell(3);
        if (summaryReportData.getIsOk().equals(true)) {
            cell.setCellValue("OK");
            cell.setCellStyle(okStyle);
        } else {
            cell.setCellValue("KO");
            cell.setCellStyle(koStyle);
        }
    }

    static void addDescriptionData(XSSFSheet sheet, int rowIndex, SummaryReportData summaryReportData, double toleranceThreshold) {
        for (int i = 0; i < summaryReportData.getMetricDescriptions().size(); i++) {
            Row row = sheet.createRow(rowIndex + i);
            row.setRowStyle(whiteStyle);

            Cell cell = row.createCell(0);
            cell.setCellValue(summaryReportData.getMetricDescriptions().get(i).getMetricName());
            cell.setCellStyle(whiteBorderedStyle);

            cell = row.createCell(1);
            cell.setCellValue(summaryReportData.getMetricDescriptions().get(i).getMetricDashboardQuantity().toString() + '/' + summaryReportData.getMetricDescriptions().get(i).getMetricFieldQuantity().toString());
            cell.setCellStyle(whiteBorderedStyle);

            double recievedValuesPercentage = ((double) summaryReportData.getMetricDescriptions().get(i).getMetricDashboardQuantity() / (double) summaryReportData.getMetricDescriptions().get(i).getMetricFieldQuantity()) * 100;

            cell = row.createCell(2);
            cell.setCellValue(recievedValuesPercentage);
            cell.setCellStyle(percentageStyle);

            cell = row.createCell(3);
            if (summaryReportData.getMetricDescriptions().get(i).getIsOK().equals(true)) {
                cell.setCellValue("OK");
                cell.setCellStyle(okStyle);
            } else {
                cell.setCellValue("KO");
                cell.setCellStyle(koStyle);
            }
        }
    }
}
