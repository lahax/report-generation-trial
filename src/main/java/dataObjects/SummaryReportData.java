package dataObjects;

import lombok.AllArgsConstructor;
import lombok.Data;
import java.util.ArrayList;

//questa classe va istanziata due volte: una per le telemisure e una per i telesegnali

@Data
@AllArgsConstructor
public class SummaryReportData {
    private Integer fieldDataTotal;
    private Integer dashboardDataTotal;
    private ArrayList<SummaryReportMetricDescription> metricDescriptions;
    private Boolean isOk;
}
