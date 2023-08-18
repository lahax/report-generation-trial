package utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import reportGenerationService.dataSimulator.DataSimulator;
import shared.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

public class ReportGenerator {

    public static void main(String[] args) {
        //Dati usati per testare che la generazione di excel funzioni
        TestPointInfo testPointInfo = new TestPointInfo(
                "NORD",
                "PESCHIERA - BOLGIANO CS (23495D1)",
                "23495D1-CUSTOM-CIRCUITO",
                "CIRCUITO",
                "AGGREGATI_CIRCUITO",
                "PRYSMIAN",
                "1:1:3:4:6:6889",
                Timestamp.valueOf("2023-06-13 05:43:11.000"),
                Timestamp.valueOf("2023-06-14 21:12:06.000"),
                3,
                4,
                92.00,
                95.00
        );

        ArrayList<SummaryReportMetricDescription> telemetryReportMetricDescription = new ArrayList<>();
        telemetryReportMetricDescription.add(new SummaryReportMetricDescription("Corrente di Schermo Fase 4", 200, 187, true));
        telemetryReportMetricDescription.add(new SummaryReportMetricDescription("Corrente di Schermo Fase 8", 200, 198, true));
        telemetryReportMetricDescription.add(new SummaryReportMetricDescription("Corrente di Schermo Fase 12", 278, 147, false));

        SummaryReportData telemetryReportData = new SummaryReportData(678, 532, telemetryReportMetricDescription, false);

        ArrayList<SummaryReportMetricDescription> telesignalReportMetricDescription = new ArrayList<>();
        telesignalReportMetricDescription.add(new SummaryReportMetricDescription("Aggregato Allarme Pressione Olio Cavo Bassa", 200, 184, true));
        telesignalReportMetricDescription.add(new SummaryReportMetricDescription("Aggregato Allarme Pressione Olio Cavo Bassissima", 300, 289, true));
        telesignalReportMetricDescription.add(new SummaryReportMetricDescription("Aggregato Allarme Pressione Olio Cavo Alta", 100, 99, true));
        telesignalReportMetricDescription.add(new SummaryReportMetricDescription("Aggregato Anomaila Apparati Monitoraggio Cavo", 600, 578, true));

        SummaryReportData telesignalReportData = new SummaryReportData(1200, 1150, telesignalReportMetricDescription, true);


        Workbook workbookReportDettaglio = createReport(testPointInfo, telemetryReportData, telesignalReportData);
        Workbook workbookSchedaCollaudo = createScheda(testPointInfo, telemetryReportMetricDescription, telesignalReportMetricDescription);

        //TODO mettere il current Time and Date nel file name,
        // va bene questo format?
        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        SimpleDateFormat sdfReport = new SimpleDateFormat("yyyy-MM-dd_HH.mm.ss");
        String filePathReport = "C:\\Users\\haxla\\Documents\\concept\\terna_report\\"+sdfReport.format(timestamp)+"_prova_report.xlsx";
        String filePathScheda = "C:\\Users\\haxla\\Documents\\concept\\terna_report\\"+sdfReport.format(timestamp)+"_prova_scheda_collaudo.xlsx";

        try (FileOutputStream fileOutputStream = new FileOutputStream(filePathReport)) {
            workbookReportDettaglio.write(fileOutputStream);
            System.out.println("Workbook Report salvato con successo!");
        } catch (IOException e) {
            e.printStackTrace();
        }
        try (FileOutputStream fileOutputStream = new FileOutputStream(filePathScheda)) {
            workbookSchedaCollaudo.write(fileOutputStream);
            System.out.println("Workbook Scheda salvato con successo!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    //bisognerà passargli una classe contenente arrayList con dati e risultati del test
    public static Workbook createReport(TestPointInfo testPointInfo, SummaryReportData telemetrySummary, SummaryReportData telesignalSummary) {
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy - HH:mm:ss");
        SimpleDateFormat sdf2 = new SimpleDateFormat("dd/MM/yyyy");
        Workbook workbook = new XSSFWorkbook();
        ExcelDataUtils.initializeStyles(workbook);
        //Sheet 1
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Titolo Sheet");
        ExcelDataUtils.addTitleB(sheet, 0, "Resoconto P2P " + sdf2.format(testPointInfo.getP2PStartTime()) + " " + testPointInfo.getDeviceID());
        ExcelDataUtils.addWhiteLine(sheet, 1);
        ExcelDataUtils.addHeaderData(sheet, 2, "DT", testPointInfo.getDT());
        ExcelDataUtils.addHeaderData(sheet, 3, "Circuito", testPointInfo.getCircuit());
        ExcelDataUtils.addHeaderData(sheet, 4, "Test Point", testPointInfo.getTestPoint());
        ExcelDataUtils.addHeaderData(sheet, 5, "Nome Test Point", testPointInfo.getTestPointName());
        ExcelDataUtils.addHeaderData(sheet, 6, "Position Code Type", testPointInfo.getPositionCodeType());
        ExcelDataUtils.addHeaderData(sheet, 7, "Vendor", testPointInfo.getVendor());
        ExcelDataUtils.addHeaderData(sheet, 8, "Device ID", testPointInfo.getDeviceID());
        ExcelDataUtils.addHeaderData(sheet, 9, "Inizio P2P", sdf.format(testPointInfo.getP2PStartTime()));
        ExcelDataUtils.addHeaderData(sheet, 10, "Fine P2P", sdf.format(testPointInfo.getP2PEndTime()));
        ExcelDataUtils.addHeaderData(sheet, 11, "Telemisure", testPointInfo.getTelemetryQuantity());
        ExcelDataUtils.addHeaderData(sheet, 12, "Telesegnali", testPointInfo.getTeleSignalQuantity());
        ExcelDataUtils.addHeaderData(sheet, 13, "Soglia di tolleranza impostata per Telemisure", testPointInfo.getTelemetryThreshold());
        ExcelDataUtils.addHeaderData(sheet, 14, "Soglia di tolleranza impostata per Telesegnali", testPointInfo.getTeleSignalThreshold());
        ExcelDataUtils.addWhiteLine(sheet, 15);
        int rowIndex = 16;
        if (!testPointInfo.getTelemetryQuantity().equals(0)) {
            ExcelDataUtils.addTitle(sheet, rowIndex, "Quantità Dati - Telemisure");
            rowIndex++;
            ExcelDataUtils.addDetailsHeader(sheet, rowIndex, "Dati campo", "Dal centro", "Ricevuti", "Esito");
            rowIndex++;
            ExcelDataUtils.addQuantityData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex++;
            ExcelDataUtils.addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }
        if (!testPointInfo.getTeleSignalQuantity().equals(0)) {
            ExcelDataUtils.addTitle(sheet, rowIndex, "Quantità Dati - Telesegnali");
            rowIndex++;
            ExcelDataUtils.addDetailsHeader(sheet, rowIndex, "Dati campo", "Dal centro", "Ricevuti", "Esito");
            rowIndex++;
            ExcelDataUtils.addQuantityData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
            rowIndex++;
            ExcelDataUtils.addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }
        if (!testPointInfo.getTelemetryQuantity().equals(0)) {
            ExcelDataUtils.addTitle(sheet, rowIndex, "Coerenza Dati - Telemisure");
            rowIndex++;
            ExcelDataUtils.addDetailsHeader2(sheet, rowIndex, "Esaminati", "Corrispondenti", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addRecapDescriptionData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex++;
            ExcelDataUtils.addDetailsHeader2(sheet, rowIndex, "Nome", "Descrizione", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addDescriptionData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex += telemetrySummary.getMetricDescriptions().size();
            ExcelDataUtils.addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }
        if (!testPointInfo.getTeleSignalQuantity().equals(0)) {
            ExcelDataUtils.addTitle(sheet, rowIndex, "Coerenza Dati - Telesegnali");
            rowIndex++;
            ExcelDataUtils.addDetailsHeader2(sheet, rowIndex, "Esaminati", "Corrispondenti", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addRecapDescriptionData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
            rowIndex++;
            ExcelDataUtils.addDetailsHeader2(sheet, rowIndex, "Nome", "Descrizione", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addDescriptionData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
        }
        for (int i = 0; i <5; i++){
            sheet.autoSizeColumn(i, true);
        }
        //Sheet 2
        XSSFSheet sheetDetail = (XSSFSheet) workbook.createSheet("Detail");
        ExcelDataUtils.addTitleDetail(sheetDetail, 0, "Confronto Dati - Aggregato Allarme Pressione Olio Cavo Bassa");
        ExcelDataUtils.addCompareHeaderDetailedReport(sheetDetail,1,"Dal Campo", "Dal Centro");
        ExcelDataUtils.addCompareHeaderDetailedReport2(sheetDetail, 2, "Index", "Esito", "Timestamp", "Timestamp (RAW)", "Valore", "Quality Code", "Ricezione");
        DataSimulator mc = new DataSimulator();
        ArrayList<MetricComparison> records = mc.generateRandomData("Aggregato Allarme Pressione Olio Cavo Bassa");
        ExcelDataUtils.addCompareRecord(sheetDetail, records);
        for (int i = 0; i <13; i++){
            sheetDetail.autoSizeColumn(i);
        }
        sheetDetail.setColumnWidth(7, 2 * 256);
        return workbook;
    }
    public static Workbook createScheda(TestPointInfo testPointInfo, ArrayList<SummaryReportMetricDescription> telemetryReportMetricDescription, ArrayList<SummaryReportMetricDescription> telesignalReportMetricDescription){
        Workbook workbook = new XSSFWorkbook();
        ExcelDataUtils.initializeStyles(workbook);
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Scheda Collaudo");
        ExcelDataUtils.addHeaderSchedaCollaudo(sheet);
        ArrayList<SchedaInfo> records = DataSimulator.getCollaudoData(testPointInfo, telemetryReportMetricDescription, telesignalReportMetricDescription);
        ExcelDataUtils.addSchedaCollaudoRecords(sheet, records);
        return workbook;
    }
}