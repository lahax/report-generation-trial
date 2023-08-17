package shared;

import lombok.AllArgsConstructor;
import lombok.Data;

@Data
@AllArgsConstructor
public class Metric {
    private String metricType;
    private Double value;
    private Integer qualityCode;
    private Long timestamp;
    private Long timestampReceivedOn;
}
