package utils;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import shared.SummaryReportData;
import shared.SummaryReportMetricDescription;
import shared.TestPointInfo;
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


        Workbook workbook = createSummaryReportTemplate(testPointInfo, telemetryReportData, telesignalReportData);
        //TODO mettere il current Time and Date nel file name
        String filePath = "C:\\Users\\Exacon\\Documents\\Prove\\prova_report.xlsx";

        try (FileOutputStream fileOutputStream = new FileOutputStream(filePath)) {
            workbook.write(fileOutputStream);
            System.out.println("Workbook salvato con successo!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    //bisognerà passargli una classe contenente arrayList con dati e risultati del test
    public static Workbook createSummaryReportTemplate(TestPointInfo testPointInfo, SummaryReportData telemetrySummary, SummaryReportData telesignalSummary) {
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy - HH:mm:ss");
        Workbook workbook = new XSSFWorkbook();
        ExcelDataUtils.initializeStyles(workbook);


        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Titolo Sheet");
        sheet.setColumnWidth(0, 25 * 256);
        sheet.setColumnWidth(1, 25 * 256);
        sheet.setColumnWidth(2, 25 * 256);
        sheet.setColumnWidth(3, 25 * 256);

        ExcelDataUtils.addTitle(sheet, 0, "Resoconto P2P Data - DeviceID");

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
            ExcelDataUtils.addDetailsHeader(sheet, rowIndex, "Esaminati", "Corrispondenti", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addRecapDescriptionData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex++;
            ExcelDataUtils.addDetailsHeader(sheet, rowIndex, "Nome", "Descrizione", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addDescriptionData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex += telemetrySummary.getMetricDescriptions().size();
            ExcelDataUtils.addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }

        if (!testPointInfo.getTeleSignalQuantity().equals(0)) {
            ExcelDataUtils.addTitle(sheet, rowIndex, "Coerenza Dati - Telesegnali");
            rowIndex++;
            ExcelDataUtils.addDetailsHeader(sheet, rowIndex, "Esaminati", "Corrispondenti", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addRecapDescriptionData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
            rowIndex++;
            ExcelDataUtils.addDetailsHeader(sheet, rowIndex, "Nome", "Descrizione", "Percentuale", "Esito");
            rowIndex++;
            ExcelDataUtils.addDescriptionData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
        }

        return workbook;
    }
}
