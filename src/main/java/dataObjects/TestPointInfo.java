package dataObjects;

import lombok.AllArgsConstructor;
import lombok.Data;
import java.sql.Timestamp;

@Data
@AllArgsConstructor
public class TestPointInfo {
    private String DT;
    private String circuit;
    private String testPoint;
    private String testPointName;
    private String positionCodeType;
    private String vendor;
    private String deviceID;
    private Timestamp P2PStartTime;
    private Timestamp P2PEndTime;
    private Integer telemetryQuantity;
    private Integer teleSignalQuantity;
    private double telemetryThreshold;
    private double teleSignalThreshold;
}
