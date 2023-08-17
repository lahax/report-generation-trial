package data_comparison.services;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;
import org.springframework.stereotype.Service;
import shared.MetricComparison;
import shared.TestResult;

@Service
public class DataComparisonService {
    private static long getTimeDifferenceInSeconds(Timestamp timestamp1, Timestamp timestamp2) {
        long milliseconds1 = timestamp1.getTime();
        long milliseconds2 = timestamp2.getTime();
        long differenceMilliseconds = Math.abs(milliseconds1 - milliseconds2);
        return TimeUnit.MILLISECONDS.toSeconds(differenceMilliseconds);
    }

    public ArrayList<MetricComparison> processMetrics(ArrayList<MetricComparison> metricList) {
        for (MetricComparison metric : metricList) {
            metric.setTimestampCorresponds(metric.getFieldTimestamp().equals(metric.getDashboardTimestamp()));
            metric.setValueCorresponds(metric.getFieldValue().equals(metric.getDashboardValue()));
            metric.setQualityCodeCorresponds(metric.getFieldQualityCode().equals(metric.getDashboardQualityCode()));

            // Imposta isdelayed a true se la differenza è maggiore di 60 secondi, altrimenti a false
            metric.setIsDelayed(getTimeDifferenceInSeconds(metric.getDashboardReceivedOn(), metric.getFieldReceivedOn()) > 60);

            if (metric.getTimestampCorresponds() && metric.getValueCorresponds() && metric.getQualityCodeCorresponds() && !metric.getIsDelayed()) {
                metric.setTestResult(TestResult.PASSED);
            } else if (metric.getTimestampCorresponds() && metric.getValueCorresponds() && metric.getQualityCodeCorresponds() && metric.getIsDelayed()) {
                metric.setTestResult(TestResult.DELAYED);
            } else {
                metric.setTestResult(TestResult.FAILED);
            }
        }

        return metricList;
    }
}
