package reportGenerationService.dataSimulator;

import shared.*;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Random;

public class DataSimulator {
    private static final int NUM_ELEMENTS = 100;
    //metricType è il parametro che andrà usato come titolo nel report di dettaglio, nei test va passato manualmente come stringa
    public static ArrayList<MetricComparison> generateRandomData(String metricType) {

        ArrayList<MetricComparison> metricComparisons = new ArrayList<>();
        Random random = new Random();
        MetricComparison metricComparison;

        for (int i = 0; i < NUM_ELEMENTS; i++) {
            Timestamp base = getRandomTimestamp();
            metricComparison = new MetricComparison();
            metricComparisons.add(metricComparison);

            //testResult: indicato nella String richiesta dalla funzione
            metricComparison.setMetricType(metricType);
            //testResult random
            metricComparison.setTestResult(getRandomTestResult());

            //setFieldTimestamp
            metricComparison.setFieldTimestamp(base);
            //setDashboardTimestamp
            metricComparison.setDashboardTimestamp(base);
            Timestamp later = new Timestamp(base.getTime() + 2000);
            //setFieldReceivedOn
            metricComparison.setFieldReceivedOn(later);
            //setdashboardReceivedOn
            metricComparison.setDashboardReceivedOn(later);

            //setFieldValue && setDashBoardValue
            int randomFieldValue = random.nextInt(2);
            switch (randomFieldValue) {
                case 0:
                    metricComparison.setFieldValue(0);
                    metricComparison.setDashboardValue(0);
                    break;
                case 1:
                    metricComparison.setFieldValue(1);
                    metricComparison.setDashboardValue(1);
                default:
                    break;
            }

            //setFieldQualityCode && setDashboardQualityCode
            int randomFieldQualityCode = random.nextInt(2);
            switch (randomFieldQualityCode) {
                case 0:
                    metricComparison.setFieldQualityCode(true);
                    metricComparison.setDashboardQualityCode(true);
                    break;
                case 1:
                    metricComparison.setFieldQualityCode(false);
                    metricComparison.setDashboardQualityCode(false);
                default:
                    break;
            }

            if (metricComparison.getTestResult().equals(TestResult.PASSED)) {
                Timestamp newLaterPassed = new Timestamp(later.getTime() + 30000);
                metricComparison.setTimestampCorresponds(true);
                metricComparison.setValueCorresponds(true);
                metricComparison.setQualityCodeCorresponds(true);
                metricComparison.setIsDelayed(false);
                metricComparison.setDashboardReceivedOn(newLaterPassed);

            } else if (metricComparison.getTestResult().equals(TestResult.DELAYED)) {
                Timestamp newLaterDelayed = new Timestamp(later.getTime() + 120000);
                metricComparison.setTimestampCorresponds(true);
                metricComparison.setValueCorresponds(true);
                metricComparison.setQualityCodeCorresponds(true);
                metricComparison.setIsDelayed(true);
                metricComparison.setDashboardReceivedOn(newLaterDelayed);
            } else {
                int randomValueFalse = random.nextInt(3); //almeno uno dei tre booleani viene settato a false -> test FAILED
                switch (randomValueFalse) {
                    case 0:
                        metricComparison.setTimestampCorresponds(false);
                        metricComparison.setDashboardTimestamp(null);
                        metricComparison.setDashboardValue(null);
                        metricComparison.setDashboardQualityCode(null);
                        metricComparison.setDashboardReceivedOn(null);
                        break;
                    case 1:
                        metricComparison.setValueCorresponds(false);
                        metricComparison.setFieldValue(0);
                        metricComparison.setDashboardValue(1);
                        metricComparison.setIsDelayed(random.nextBoolean());
                        break;
                    case 2:
                        metricComparison.setQualityCodeCorresponds(false);
                        metricComparison.setFieldQualityCode(true);
                        metricComparison.setDashboardQualityCode(false);
                        metricComparison.setIsDelayed(random.nextBoolean());
                        break;
                    default:
                        break;
                }
                if(metricComparison.getFieldValue().equals(metricComparison.getDashboardValue())){
                    metricComparison.setValueCorresponds(true);
                }else{
                    metricComparison.setValueCorresponds(false);
                }
                if(metricComparison.getFieldQualityCode().equals(metricComparison.getDashboardQualityCode())){
                    metricComparison.setQualityCodeCorresponds(true);
                }else{
                    metricComparison.setQualityCodeCorresponds(false);
                }
            }
        }
        return metricComparisons;
    }
    private static TestResult getRandomTestResult() {
        Random random = new Random();
        int randomNum = random.nextInt(3);
        switch (randomNum) {
            case 0:
                return TestResult.PASSED;
            case 1:
                return TestResult.FAILED;
            case 2:
            default:
                return TestResult.DELAYED;
        }
    }
    private static Timestamp getRandomTimestamp() {
        Random random = new Random();
        long offset = Timestamp.valueOf("2020-01-01 00:00:00").getTime();
        long end = Timestamp.valueOf("2023-01-01 00:00:00").getTime();

        // Genera un timestamp casuale tra 2020-01-01 e 2023-01-01
        long diff = end - offset + 1;
        long randomTimestamp = offset + (long) (random.nextDouble() * diff);

        return new Timestamp(randomTimestamp);
    }
    public static ArrayList<SchedaInfo> getCollaudoData(TestPointInfo testPointInfo, ArrayList<SummaryReportMetricDescription> telemetryReportMetricDescription, ArrayList<SummaryReportMetricDescription> telesignalReportMetricDescription) {

        ArrayList<SchedaInfo> testingInfos = new ArrayList<>();
        SchedaInfo schedaInfo;

        if (telemetryReportMetricDescription.size()!=0){
            for(int i = 0; i < telemetryReportMetricDescription.size(); i++){
                String deviceID = testPointInfo.getDeviceID();
                String descrizione = telemetryReportMetricDescription.get(i).getMetricName(); //circuit_desc, parte del device id e metric name
                String testpoint = testPointInfo.getTestPoint(); // position_code_desc
                String serial = testPointInfo.getDeviceID(); //serial è la seconda parte del device id
                Boolean result = telemetryReportMetricDescription.get(i).getIsOK();
                Timestamp p2pStart = testPointInfo.getP2PStartTime();
                schedaInfo = new SchedaInfo("8094", "CTNT", 0,0, 0,16, null, null, descrizione, testpoint, serial, result, null, deviceID, 0, 0, p2pStart, null);
                testingInfos.add(schedaInfo);
            }
        }
        if (telesignalReportMetricDescription.size()!=0){
            for(int i = 0; i < telesignalReportMetricDescription.size(); i++){
                String deviceID = testPointInfo.getDeviceID();
                String descrizione = telesignalReportMetricDescription.get(i).getMetricName(); //circuit_desc, parte del device id e metric name
                String testpoint = testPointInfo.getTestPoint(); // position_code_desc
                String serial = testPointInfo.getDeviceID(); //serial è la seconda parte del device id
                Boolean result = telesignalReportMetricDescription.get(i).getIsOK();
                Timestamp p2pStart = testPointInfo.getP2PStartTime();
                schedaInfo = new SchedaInfo("8094", "CTNT", 0,0, 0,16, null, null, descrizione, testpoint, serial, result, null, deviceID, 0, 0, p2pStart, null);
                testingInfos.add(schedaInfo);
            }
        }
        return testingInfos;
    }
}