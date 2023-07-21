package dataObjects;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class SummaryReportMetricDescription {
    private String metricName;
    private Integer metricFieldQuantity;
    private Integer metricDashboardQuantity;
    private Boolean isOK;
}
