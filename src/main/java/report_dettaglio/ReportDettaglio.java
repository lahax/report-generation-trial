package report_dettaglio;

import java.io.FileOutputStream;

import dataObjects.SummaryReportData;
import dataObjects.SummaryReportMetricDescription;
import dataObjects.TestPointInfo;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import static report_dettaglio.ExcelDataUtils.*;

public class ReportDettaglio {


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
        initializeStyles(workbook);


        XSSFSheet sheet = (XSSFSheet) workbook.createSheet("Titolo Sheet");
        sheet.setColumnWidth(0, 25 * 256);
        sheet.setColumnWidth(1, 25 * 256);
        sheet.setColumnWidth(2, 25 * 256);
        sheet.setColumnWidth(3, 25 * 256);

        addTitle(sheet, 0, "Resoconto P2P Data - DeviceID");

        addWhiteLine(sheet, 1);

        addHeaderData(sheet, 2, "DT", testPointInfo.getDT());
        addHeaderData(sheet, 3, "Circuito", testPointInfo.getCircuit());
        addHeaderData(sheet, 4, "Test Point", testPointInfo.getTestPoint());
        addHeaderData(sheet, 5, "Nome Test Point", testPointInfo.getTestPointName());
        addHeaderData(sheet, 6, "Position Code Type", testPointInfo.getPositionCodeType());
        addHeaderData(sheet, 7, "Vendor", testPointInfo.getVendor());
        addHeaderData(sheet, 8, "Device ID", testPointInfo.getDeviceID());
        addHeaderData(sheet, 9, "Inizio P2P", sdf.format(testPointInfo.getP2PStartTime()));
        addHeaderData(sheet, 10, "Fine P2P", sdf.format(testPointInfo.getP2PEndTime()));
        addHeaderData(sheet, 11, "Telemisure", testPointInfo.getTelemetryQuantity());
        addHeaderData(sheet, 12, "Telesegnali", testPointInfo.getTeleSignalQuantity());
        addHeaderData(sheet, 13, "Soglia di tolleranza impostata per Telemisure", testPointInfo.getTelemetryThreshold());
        addHeaderData(sheet, 14, "Soglia di tolleranza impostata per Telesegnali", testPointInfo.getTeleSignalThreshold());

        addWhiteLine(sheet, 15);

        int rowIndex = 16;

        if (!testPointInfo.getTelemetryQuantity().equals(0)) {
            addTitle(sheet, rowIndex, "Quantità Dati - Telemisure");
            rowIndex++;
            addDetailsHeader(sheet, rowIndex, "Dati campo", "Dal centro", "Ricevuti", "Esito");
            rowIndex++;
            addQuantityData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex++;
            addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }

        if (!testPointInfo.getTeleSignalQuantity().equals(0)) {
            addTitle(sheet, rowIndex, "Quantità Dati - Telesegnali");
            rowIndex++;
            addDetailsHeader(sheet, rowIndex, "Dati campo", "Dal centro", "Ricevuti", "Esito");
            rowIndex++;
            addQuantityData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
            rowIndex++;
            addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }

        if (!testPointInfo.getTelemetryQuantity().equals(0)) {
            addTitle(sheet, rowIndex, "Coerenza Dati - Telemisure");
            rowIndex++;
            addDetailsHeader(sheet, rowIndex, "Esaminati", "Corrispondenti", "Percentuale", "Esito");
            rowIndex++;
            addRecapDescriptionData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex++;
            addDetailsHeader(sheet, rowIndex, "Nome", "Descrizione", "Percentuale", "Esito");
            rowIndex++;
            addDescriptionData(sheet, rowIndex, telemetrySummary, testPointInfo.getTelemetryThreshold());
            rowIndex += telemetrySummary.getMetricDescriptions().size();
            addWhiteLine(sheet, rowIndex);
            rowIndex++;
        }

        if (!testPointInfo.getTeleSignalQuantity().equals(0)) {
            addTitle(sheet, rowIndex, "Coerenza Dati - Telesegnali");
            rowIndex++;
            addDetailsHeader(sheet, rowIndex, "Esaminati", "Corrispondenti", "Percentuale", "Esito");
            rowIndex++;
            addRecapDescriptionData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
            rowIndex++;
            addDetailsHeader(sheet, rowIndex, "Nome", "Descrizione", "Percentuale", "Esito");
            rowIndex++;
            addDescriptionData(sheet, rowIndex, telesignalSummary, testPointInfo.getTeleSignalThreshold());
        }

        return workbook;
    }
}
