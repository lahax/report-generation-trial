package reportGenerationService.dataSimulator;

import shared.MetricComparison;
import shared.TestResult;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Random;

public class DataSimulator {
    private static final int NUM_ELEMENTS = 100;

    //metricType è il parametro che andrà usato come titolo nel report di dettaglio, nei test va passato manualmente come stringa
    private static ArrayList<MetricComparison> generateRandomData(String metricType) {
        ArrayList<MetricComparison> metricComparisons = new ArrayList<>();

        Random random = new Random();

        for (int i = 0; i < NUM_ELEMENTS; i++) {
            MetricComparison metricComparison = new MetricComparison();
            metricComparison.setMetricType(metricType);
            metricComparison.setTestResult(getRandomTestResult());

            if (metricComparison.getTestResult().equals(TestResult.PASSED)) {
                metricComparison.setTimestampCorresponds(true);
                metricComparison.setValueCorresponds(true);
                metricComparison.setQualityCodeCorresponds(true);
                metricComparison.setIsDelayed(false);
                metricComparison.setFieldReceivedOn(getRandomTimestamp());

                Timestamp dashboardRecievedOn = metricComparison.getFieldReceivedOn();
                dashboardRecievedOn.setTime(dashboardRecievedOn.getTime() + 30000); // 30.000 millisecondi = 30 secondi
                metricComparison.setDashboardReceivedOn(dashboardRecievedOn);

            } else if (metricComparison.getTestResult().equals(TestResult.DELAYED)) {
                metricComparison.setTimestampCorresponds(true);
                metricComparison.setValueCorresponds(true);
                metricComparison.setQualityCodeCorresponds(true);
                metricComparison.setIsDelayed(true);

                Timestamp dashboardRecievedOn = metricComparison.getFieldReceivedOn();
                dashboardRecievedOn.setTime(dashboardRecievedOn.getTime() + 120000); // 120.000 millisecondi = 2 minuti
                metricComparison.setDashboardReceivedOn(dashboardRecievedOn);


            } else {
                int randomValue = random.nextInt(3); //almeno uno dei tre booleani viene settato a false -> test FAILED
                switch (randomValue) {
                    case 0:
                        metricComparison.setTimestampCorresponds(false);
                        break;
                    case 1:
                        metricComparison.setValueCorresponds(false);
                        break;
                    case 2:
                        metricComparison.setQualityCodeCorresponds(false);
                        break;
                    default:
                        break;
                }
                metricComparison.setIsDelayed(random.nextBoolean());
                if (metricComparison.getIsDelayed().equals(true)) {
                    Timestamp dashboardRecievedOn = metricComparison.getFieldReceivedOn();
                    dashboardRecievedOn.setTime(dashboardRecievedOn.getTime() + 120000); // 120.000 millisecondi = 2 minuti
                    metricComparison.setDashboardReceivedOn(dashboardRecievedOn);
                } else {
                    Timestamp dashboardRecievedOn = metricComparison.getFieldReceivedOn();
                    dashboardRecievedOn.setTime(dashboardRecievedOn.getTime() + 30000); // 30.000 millisecondi = 30 secondi
                    metricComparison.setDashboardReceivedOn(dashboardRecievedOn);
                }
            }

            // Se timestampCorresponds è true, imposta fieldTimestamp e dashboardTimestamp con timestamp casuali
            if (metricComparison.getTimestampCorresponds()) {
                metricComparison.setDashboardTimestamp(getRandomTimestamp());
                metricComparison.setFieldTimestamp(metricComparison.getDashboardTimestamp());
            } else {
                metricComparison.setDashboardTimestamp(getRandomTimestamp());
                metricComparison.setFieldTimestamp(getRandomTimestamp());
            }

            // Se valueCorresponds è true, imposta fieldValue e dashboardValue con valori casuali
            if (metricComparison.getValueCorresponds()) {
                metricComparison.setDashboardValue(random.nextInt());
                metricComparison.setFieldValue(metricComparison.getDashboardValue());
            } else {
                metricComparison.setDashboardValue(random.nextInt());
                metricComparison.setFieldValue(random.nextInt());
            }

            // Se qualityCodeCorresponds è true, imposta fieldQualityCode e dashboardQualityCode con valori casuali
            if (metricComparison.getQualityCodeCorresponds()) {
                metricComparison.setDashboardQualityCode(random.nextBoolean());
                metricComparison.setFieldQualityCode(metricComparison.getDashboardQualityCode());
            } else {
                metricComparison.setDashboardQualityCode(random.nextBoolean());
                metricComparison.setFieldQualityCode(random.nextBoolean());
            }

            metricComparisons.add(metricComparison);
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
}
