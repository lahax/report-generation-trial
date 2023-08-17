package shared;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.sql.Timestamp;

//struttura dati necessaria per la generazione del report di dettaglio: contiene tutti i campi necessari
//dal data comparison service dovrebbe arrivare un ArrayList di questa classe e lo usiamo per il report di dettaglio della singola metrica
@Data
@AllArgsConstructor
@NoArgsConstructor //necessario per il DataSimulator
public class MetricComparison {
    private String MetricType;

    //esito
    private TestResult testResult;

    //dati di campo
    private Timestamp fieldTimestamp;
    private Object fieldValue;
    private Boolean fieldQualityCode;
    private Timestamp fieldReceivedOn;

    //dati dal centro (dashboard)
    private Timestamp dashboardTimestamp;
    private Object dashboardValue;
    private Boolean dashboardQualityCode;
    private Timestamp dashboardReceivedOn;

    //booleani per ogni confronto
    private Boolean timestampCorresponds; //se false manca il dato di campo, mettere tutta la riga in rosso
    private Boolean valueCorresponds; //se false campo corrispondente in rosso
    private Boolean qualityCodeCorresponds;// se false campo corrispondente in rosso
    private Boolean isDelayed; //se true vuol dire che dashboardRecievedOn è maggiore di fieldRecievedOn di più di 60 secondi, ma per il resto i campi corrispondono tutti
}
