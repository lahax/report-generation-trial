package utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import shared.SchedaInfo;
import shared.SummaryReportData;
import shared.MetricComparison;
import shared.TestResult;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import static utils.ExcelStyleUtils.*;

class ExcelDataUtils {
    private static CellStyle titleStyle;
    private static CellStyle whiteStyle;
    private static CellStyle whiteBorderedStyle;
    private static CellStyle percentageStyle;
    private static CellStyle headerStyle;
    private static CellStyle okStyle;
    private static CellStyle koStyle;
    private static CellStyle okDelayedStyle;
    private static CellStyle passedStyle;
    private static CellStyle failedStyle;
    private static CellStyle delayedStyle;
    private static CellStyle lightBlueStyle;
    private static CellStyle darkBlueStyle;
    private static CellStyle detailOkStyle;
    private static CellStyle headerCollaudoStyle;
    private static CellStyle whiteBorderedCollaudoStyle;
    private static CellStyle okStyleCollaudo;
    private static CellStyle koStyleCollaudo;
    private static CellStyle titleStyle16;
    public static void initializeStyles(Workbook workbook) {
        titleStyle = createTitleStyle(workbook, createWhiteBoldFont(workbook));
        titleStyle16 = createTitleStyle(workbook, createWhiteBoldFont16(workbook));
        headerStyle = createHeaderStyle(workbook, createBlackBoldFont(workbook));
        whiteStyle = createWhiteStyle(workbook);
        whiteBorderedStyle = createWhiteBorderedStyle(workbook, createBlackFont(workbook));
        percentageStyle = createPercentageStyle(workbook, createBlackFont(workbook));
        okStyle = createOkStyle(workbook, createWhiteFont(workbook));
        koStyle = createKoStyle(workbook, createWhiteFont(workbook));
        okDelayedStyle = createOkDelayedStyle(workbook, createBlackFont(workbook));
        lightBlueStyle = createLightBlueStyle(workbook);
        darkBlueStyle = createDarkBlueStyle(workbook);
        passedStyle = createPassedStyle(workbook, createGreenFont(workbook));
        failedStyle = createFailedStyle(workbook, createRedFont(workbook));
        delayedStyle = createDelayedStyle(workbook, createYellowFont(workbook));
        detailOkStyle = createDetailOkStyle(workbook, createBlackFont(workbook));
        headerCollaudoStyle = createHeaderCollaudoStyle(workbook, createBlackBoldFont(workbook));
        whiteBorderedCollaudoStyle = createWhiteBorderedCollaudoStyle(workbook);
        okStyleCollaudo = createOkStyleCollaudo(workbook, createGreenFont(workbook));
        koStyleCollaudo = createKoStyleCollaudo(workbook, createRedFont(workbook));
    }
    static void addTitle(XSSFSheet sheet, int rowIndex, String title) {
        Row row = sheet.createRow(rowIndex);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);
        Cell cell4 = row.createCell(3);

        Cell cell5 = row.createCell(4);
        Cell cell6 = row.createCell(5);
        cell1.setCellValue(title);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 3);
        sheet.addMergedRegion(cellRangeAddress);
        row.setRowStyle(whiteStyle);
        cell1.setCellStyle(titleStyle);
        cell2.setCellStyle(titleStyle);
        cell3.setCellStyle(titleStyle);
        cell4.setCellStyle(titleStyle);
        cell5.setCellStyle(titleStyle);
        cell6.setCellStyle(titleStyle);
    }
    static void addTitleB(XSSFSheet sheet, int rowIndex, String title) {
        Row row = sheet.createRow(rowIndex);
        Cell cell1 = row.createCell(0);
        Cell cell2 = row.createCell(1);
        Cell cell3 = row.createCell(2);
        Cell cell4 = row.createCell(3);

        //test 6
        Cell cell5 = row.createCell(4);
        Cell cell6 = row.createCell(5);

        cell1.setCellValue(title);
        //test 6
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 5);
        sheet.addMergedRegion(cellRangeAddress);
        row.setRowStyle(whiteStyle);
        cell1.setCellStyle(titleStyle16);
        cell2.setCellStyle(titleStyle16);
        cell3.setCellStyle(titleStyle16);
        cell4.setCellStyle(titleStyle16);
        //test 6
        cell5.setCellStyle(titleStyle16);
        cell6.setCellStyle(titleStyle16);
    }
    static void addHeaderData(XSSFSheet sheet, int rowIndex, String label, Object value) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
        Cell cell = row.createCell(0);
        cell.setCellValue(label);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(1);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(2);
        cell.setCellStyle(headerStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 2);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(3);
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
        cell = row.createCell(4);
        cell.setCellStyle(whiteBorderedStyle);
        cell = row.createCell(5);
        cell.setCellStyle(whiteBorderedStyle);
        cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 3, 5);
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
        cell.setCellStyle(headerStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 1);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(2);
        cell.setCellValue(title2);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(3);
        cell.setCellStyle(headerStyle);
        cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 2, 3);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(4);
        cell.setCellValue(title3);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(5);
        cell.setCellValue(title4);
        cell.setCellStyle(headerStyle);
    }
    static void addDetailsHeader2(XSSFSheet sheet, int rowIndex, String title1, String title2, String title3, String title4) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
        Cell cell = row.createCell(0);
        cell.setCellValue(title1);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(1);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(2);
        cell.setCellStyle(headerStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 2);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(3);
        cell.setCellValue(title2);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(4);
        cell.setCellValue(title3);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(5);
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
        cell.setCellStyle(whiteBorderedStyle);
        cell = row.createCell(2);
        cell.setCellStyle(whiteBorderedStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 1);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(2);
        cell.setCellValue(summaryReportData.getDashboardDataTotal());
        cell.setCellStyle(whiteBorderedStyle);
        cell = row.createCell(3);
        cell.setCellStyle(whiteBorderedStyle);
        cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 2, 3);
        sheet.addMergedRegion(cellRangeAddress);
        double recievedValuesPercentage = ((double) summaryReportData.getDashboardDataTotal() / (double) summaryReportData.getFieldDataTotal()) * 100;
        cell = row.createCell(4);
        cell.setCellValue(recievedValuesPercentage);
        cell.setCellStyle(percentageStyle);
        cell = row.createCell(5);
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
        cell.setCellStyle(whiteBorderedStyle);
        cell = row.createCell(2);
        cell.setCellStyle(whiteBorderedStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 2);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(3);
        cell.setCellValue(summaryReportData.getDashboardDataTotal().toString() + '/' + summaryReportData.getFieldDataTotal().toString());
        cell.setCellStyle(whiteBorderedStyle);
        double recievedValuesPercentage = ((double) summaryReportData.getDashboardDataTotal() / (double) summaryReportData.getFieldDataTotal()) * 100;
        cell = row.createCell(4);
        cell.setCellValue(recievedValuesPercentage);
        cell.setCellStyle(percentageStyle);
        cell = row.createCell(5);
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
            cell.setCellStyle(whiteBorderedStyle);
            cell = row.createCell(2);
            cell.setCellStyle(whiteBorderedStyle);
            CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex + i, rowIndex + i, 0, 2);
            sheet.addMergedRegion(cellRangeAddress);
            cell = row.createCell(3);
            cell.setCellValue(summaryReportData.getMetricDescriptions().get(i).getMetricDashboardQuantity().toString() + '/' + summaryReportData.getMetricDescriptions().get(i).getMetricFieldQuantity().toString());
            cell.setCellStyle(whiteBorderedStyle);
            double recievedValuesPercentage = ((double) summaryReportData.getMetricDescriptions().get(i).getMetricDashboardQuantity() / (double) summaryReportData.getMetricDescriptions().get(i).getMetricFieldQuantity()) * 100;
            cell = row.createCell(4);
            cell.setCellValue(recievedValuesPercentage);
            cell.setCellStyle(percentageStyle);
            cell = row.createCell(5);
            if (summaryReportData.getMetricDescriptions().get(i).getIsOK().equals(true)) {
                cell.setCellValue("OK");
                cell.setCellStyle(okStyle);
            } else {
                cell.setCellValue("KO");
                cell.setCellStyle(koStyle);
            }
        }
    }
    static void addTitleDetail(XSSFSheet sheet, int rowIndex, String title) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
        Cell cell = row.createCell(0);
        cell.setCellValue(title);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 12);
        sheet.addMergedRegion(cellRangeAddress);
        cell.setCellStyle(titleStyle);
    }
    static void addCompareHeaderDetailedReport(XSSFSheet sheet, int rowIndex, String title1, String title2) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
        Cell cell = row.createCell(0);
        cell.setCellStyle(lightBlueStyle);
        cell = row.createCell(1);
        cell.setCellStyle(headerStyle);
        CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 0, 1);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(2);
        cell.setCellValue(title1);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(3);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(4);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(5);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(6);
        cell.setCellStyle(headerStyle);
        cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 2, 6);
        sheet.addMergedRegion(cellRangeAddress);
        cell = row.createCell(7);
        cell.setCellStyle(darkBlueStyle);
        cell = row.createCell(8);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title2);
        cell = row.createCell(9);
        cell.setCellValue(title1);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(10);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(11);
        cell.setCellStyle(headerStyle);
        cell = row.createCell(12);
        cell.setCellStyle(headerStyle);
        cellRangeAddress = new CellRangeAddress(rowIndex, rowIndex, 8, 12);
        sheet.addMergedRegion(cellRangeAddress);
    }
    static void addCompareHeaderDetailedReport2(XSSFSheet sheet, int rowIndex, String title1, String title2, String title3, String title4, String title5, String title6, String title7) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(whiteStyle);
        Cell cell = row.createCell(0);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title1);
        cell = row.createCell(1);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title2);
        cell = row.createCell(2);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title3);
        cell = row.createCell(3);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title4);
        cell = row.createCell(4);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title5);
        cell = row.createCell(5);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title6);
        cell = row.createCell(6);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title7);
        cell = row.createCell(7);
        cell.setCellStyle(darkBlueStyle);
        cell = row.createCell(8);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title3);
        cell = row.createCell(9);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title4);
        cell = row.createCell(10);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title5);
        cell = row.createCell(11);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title6);
        cell = row.createCell(12);
        cell.setCellStyle(headerStyle);
        cell.setCellValue(title7);
    }
    static void addCompareRecord(XSSFSheet sheet, ArrayList<MetricComparison> records){
        int sizeListRecords = records.size();
        int rowIndex = 3;
        for (int i = 0; i < sizeListRecords; i++) {
            Row insertRecordInRow = sheet.createRow(rowIndex + i);
            insertRecordInRow.setRowStyle(whiteStyle);
            Cell cellIndex = insertRecordInRow.createCell(0);
            cellIndex.setCellValue(i + 1);
            cellIndex.setCellStyle(detailOkStyle);
            Cell cellEsito = insertRecordInRow.createCell(1);
            if (Objects.equals(records.get(i).getTestResult(), TestResult.PASSED)) {
                cellEsito.setCellValue("OK");
                cellEsito.setCellStyle(passedStyle);
            } else if (Objects.equals(records.get(i).getTestResult(), TestResult.FAILED)) {
                cellEsito.setCellValue("KO");
                cellEsito.setCellStyle(failedStyle);
            } else if (Objects.equals(records.get(i).getTestResult(), TestResult.DELAYED)) {
                cellEsito.setCellValue("OK");
                cellEsito.setCellStyle(delayedStyle);
            } else {
                cellEsito.setCellValue("ERROR");
            }
            Cell cellTimestampCampo = insertRecordInRow.createCell(2);
            Cell cellTimestampCentro = insertRecordInRow.createCell(8);
            Cell cellTimestampRawCampo = insertRecordInRow.createCell(3);
            Cell cellTimestampRawCentro = insertRecordInRow.createCell(9);
            Cell cellValoreCampo = insertRecordInRow.createCell(4);
            Cell cellValoreCentro = insertRecordInRow.createCell(10);
            Cell cellQualityCodeCampo = insertRecordInRow.createCell(5);
            Cell cellQualityCodeCentro = insertRecordInRow.createCell(11);
            Cell cell = insertRecordInRow.createCell(7);
            cell.setCellStyle(darkBlueStyle);
            Cell dashboardValuesNull = insertRecordInRow.createCell(8);
            Cell dashboardValuesNull2 = insertRecordInRow.createCell(9);
            Cell dashboardValuesNull3 = insertRecordInRow.createCell(10);
            Cell dashboardValuesNull4 = insertRecordInRow.createCell(11);
            Cell dashboardValuesNull5 = insertRecordInRow.createCell(12);
            Cell cellRicezioneCampo = insertRecordInRow.createCell(6);
            Cell cellRicezioneCentro = insertRecordInRow.createCell(12);
            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
            String sdfTimestampCampo = sdf.format(records.get(i).getFieldTimestamp());
            cellTimestampCampo.setCellValue(sdfTimestampCampo);
            //TODO time stamp raw and
            cellTimestampRawCampo.setCellValue(records.get(i).getFieldTimestamp());
            cellValoreCampo.setCellValue(records.get(i).getFieldValue());
            if (Objects.equals(records.get(i).getFieldQualityCode(), true)){
                cellQualityCodeCampo.setCellValue(0);
            } else if (Objects.equals(records.get(i).getFieldQualityCode(), false)){
                cellQualityCodeCampo.setCellValue(1);
            } else {
                cellQualityCodeCampo.setCellValue("Error 415");
            }
            String sdfReceivedOnCampo = sdf.format(records.get(i).getFieldReceivedOn());
            cellRicezioneCampo.setCellValue(sdfReceivedOnCampo);
            if(records.get(i).getDashboardTimestamp()==null){
                cellTimestampCampo.setCellStyle(koStyle);
                //TODO time stamp raw
                cellTimestampRawCampo.setCellStyle(koStyle);
                cellValoreCampo.setCellStyle(koStyle);
                cellQualityCodeCampo.setCellStyle(koStyle);
                cellRicezioneCampo.setCellStyle(koStyle);
                dashboardValuesNull.setCellValue("N.D.");
                CellRangeAddress cellRangeAddress = new CellRangeAddress(rowIndex + i, rowIndex + i, 8, 12);
                sheet.addMergedRegion(cellRangeAddress);
                dashboardValuesNull.setCellStyle(koStyle);
                dashboardValuesNull2.setCellStyle(koStyle);
                dashboardValuesNull3.setCellStyle(koStyle);
                dashboardValuesNull4.setCellStyle(koStyle);
                dashboardValuesNull5.setCellStyle(koStyle);
            } else if (records.get(i).getDashboardTimestamp()!=null){
                if (Objects.equals(records.get(i).getFieldTimestamp(), records.get(i).getDashboardTimestamp())){
                    cellTimestampCampo.setCellStyle(detailOkStyle);
                    String sdfTimestampCentro = sdf.format(records.get(i).getDashboardTimestamp());
                    cellTimestampCentro.setCellValue(sdfTimestampCentro);
                    cellTimestampCentro.setCellStyle(detailOkStyle);
                    //TODO timestampRAW
                    cellTimestampRawCampo.setCellStyle(detailOkStyle);
                    cellTimestampRawCentro.setCellValue(records.get(i).getDashboardTimestamp());
                    cellTimestampRawCentro.setCellStyle(detailOkStyle);
                } else {
                    cellTimestampCampo.setCellValue("ERROR 445");
                    cellTimestampCentro.setCellValue("ERROR 446");
                }
                cellValoreCentro.setCellValue(records.get(i).getDashboardValue());
                if (Objects.equals(records.get(i).getValueCorresponds(), true)) {
                    cellValoreCampo.setCellStyle(detailOkStyle);
                    cellValoreCentro.setCellStyle(detailOkStyle);
                }
                if (Objects.equals(records.get(i).getValueCorresponds(), false)) {
                    cellValoreCampo.setCellStyle(koStyle);
                    cellValoreCentro.setCellStyle(koStyle);
                }
                if (Objects.equals(records.get(i).getDashboardQualityCode(), true)){
                    cellQualityCodeCentro.setCellValue(0);
                } else if (Objects.equals(records.get(i).getDashboardQualityCode(), false)){
                    cellQualityCodeCentro.setCellValue(1);
                } else {
                    cellQualityCodeCentro.setCellValue("Error 462");
                }
                if(Objects.equals(records.get(i).getQualityCodeCorresponds(),true)) {
                    cellQualityCodeCampo.setCellStyle(detailOkStyle);
                    cellQualityCodeCentro.setCellStyle(detailOkStyle);
                } else if (Objects.equals(records.get(i).getQualityCodeCorresponds(),false)) {
                    cellQualityCodeCampo.setCellStyle(koStyle);
                    cellQualityCodeCentro.setCellStyle(koStyle);
                } else {
                    cellQualityCodeCentro.setCellValue("Error 471");
                }
                String sdfReceivedOnCentro = sdf.format(records.get(i).getDashboardReceivedOn());
                cellRicezioneCentro.setCellValue(sdfReceivedOnCentro);
                if(Objects.equals(records.get(i).getIsDelayed(),true)) {
                    cellRicezioneCampo.setCellStyle(okDelayedStyle);
                    cellRicezioneCentro.setCellStyle(okDelayedStyle);
                }else if(Objects.equals(records.get(i).getIsDelayed(),false)){
                    cellRicezioneCampo.setCellStyle(detailOkStyle);
                    cellRicezioneCentro.setCellStyle(detailOkStyle);
                }else{
                    cellRicezioneCampo.setCellValue("ERROR 482");
                    cellRicezioneCentro.setCellValue("ERROR 483");
                }
            } else {
                cellTimestampCampo.setCellValue("ERROR 486");
                cellTimestampCentro.setCellValue("ERROR 487");
            }
        }
    }
    static void addHeaderSchedaCollaudo(XSSFSheet sheet){
        Row row = sheet.createRow(0);
        row.setRowStyle(whiteStyle);
        List<String> header = List.of("    CA    ","    IMPIANTO    ", "    IOA_TYPE    ", "    IOA_DEV    ", "    IOA_NUM    ", "    IOA    ", "    SCCT    ", "    PSE    ", "    DESCRIZIONE    ", "    TESTPOINT    ", "    SERIAL    ", "    ESITO P2P grandezze non aggregate    ", "    ESITO P2P grandezze aggregate    ", "    Misura prima del test    ", "    Misura dopo il test    ", "    Data di Esecuzione Test    ", "    Note    ");
        Cell cell;
        for (int i = 0; i < header.size(); i++){
            cell = row.createCell(i);
            cell.setCellValue(header.get(i));
            cell.setCellStyle(headerCollaudoStyle);
        }
    }
    static void addSchedaCollaudoRecords(XSSFSheet sheet, ArrayList<SchedaInfo> records){
        int sizeListRecords = records.size();
        int rowIndex = 1;
        Row row;
        for (int i = 0; i < sizeListRecords; i++) {
            row = sheet.createRow(rowIndex + i);
            row.setRowStyle(whiteStyle);
            Cell ca = row.createCell(0);
            ca.setCellValue(records.get(i).getCa());
            ca.setCellStyle(whiteBorderedCollaudoStyle);
            Cell impianto = row.createCell(1);
            impianto.setCellValue(records.get(i).getImpianto());
            impianto.setCellStyle(whiteBorderedCollaudoStyle);
            Cell ioaType = row.createCell(2);
            ioaType.setCellValue(records.get(i).getIoaType());
            ioaType.setCellStyle(whiteBorderedCollaudoStyle);
            Cell ioaDev = row.createCell(3);
            ioaDev.setCellValue(records.get(i).getIoaDev());
            ioaDev.setCellStyle(whiteBorderedCollaudoStyle);
            Cell ioaNum = row.createCell(4);
            ioaNum.setCellValue(records.get(i).getIoaNum());
            ioaNum.setCellStyle(whiteBorderedCollaudoStyle);
            Cell ioa = row.createCell(5);
            ioa.setCellValue(records.get(i).getIoa());
            ioa.setCellStyle(whiteBorderedCollaudoStyle);
            Cell scct = row.createCell(6);
            scct.setCellValue(records.get(i).getScct());
            scct.setCellStyle(whiteBorderedCollaudoStyle);
            Cell pse = row.createCell(7);
            pse.setCellValue(records.get(i).getPse());
            pse.setCellStyle(whiteBorderedCollaudoStyle);
            Cell descrizione = row.createCell(8);
            descrizione.setCellValue(records.get(i).getDescrizione());
            descrizione.setCellStyle(whiteBorderedCollaudoStyle);
            Cell testpoint = row.createCell(9);
            testpoint.setCellValue(records.get(i).getTestpoint());
            testpoint.setCellStyle(whiteBorderedCollaudoStyle);
            Cell serial = row.createCell(10);
            serial.setCellValue(records.get(i).getSerial());
            serial.setCellStyle(whiteBorderedCollaudoStyle);
            Cell esitoP2PNonAggregate = row.createCell(11);
            if(records.get(i).getP2pResult().equals(true)){
                esitoP2PNonAggregate.setCellValue("OK");
                esitoP2PNonAggregate.setCellStyle(okStyleCollaudo);
            }else if(records.get(i).getP2pResult().equals(false)){
                esitoP2PNonAggregate.setCellValue("KO");
                esitoP2PNonAggregate.setCellStyle(koStyleCollaudo);
            } else {
                esitoP2PNonAggregate.setCellValue("Error 550");
            }
            Cell esitoP2PAggregate = row.createCell(12);
            esitoP2PAggregate.setCellValue("");
            esitoP2PAggregate.setCellStyle(whiteBorderedCollaudoStyle);
            Cell misuraPrimaTest = row.createCell(13);
            misuraPrimaTest.setCellValue(records.get(i).getMeasureBeforeTest());
            misuraPrimaTest.setCellStyle(whiteBorderedCollaudoStyle);
            Cell misuraDopoTest = row.createCell(14);
            misuraDopoTest.setCellValue(records.get(i).getMeasureAfterTest());
            misuraDopoTest.setCellStyle(whiteBorderedCollaudoStyle);
            SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
            Cell data = row.createCell(15);
            data.setCellValue(sdf.format(records.get(i).getP2PStartTime()));
            data.setCellStyle(whiteBorderedCollaudoStyle);
            Cell note = row.createCell(16);
            note.setCellValue(records.get(i).getNote());
            note.setCellStyle(whiteBorderedCollaudoStyle);
        }
        sheet.setAutoFilter(new CellRangeAddress(0, sizeListRecords, 0, 16));
        for (int i = 0; i<16; i++){
            sheet.autoSizeColumn(i);
        }
    }
}