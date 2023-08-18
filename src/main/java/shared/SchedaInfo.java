package shared;

import lombok.*;
import java.sql.Timestamp;

@Data
@AllArgsConstructor
@NoArgsConstructor
//Oggetto per la creazione di ogni riga della scheda collaudo
public class SchedaInfo {
    private String ca;
    private String impianto;
    private int ioaType;
    private int ioaDev;
    private int ioaNum;
    private int ioa;
    private String scct; //empty
    private String pse; //empty
    private String descrizione;
    private String testpoint;
    private String serial;
    private Boolean p2pResult;
    private String p2pResultAggregate; //empty
    private String deviceID;
    private double measureBeforeTest; //empty?
    private double measureAfterTest; //empty?
    private Timestamp P2PStartTime;
    private String note;
}
