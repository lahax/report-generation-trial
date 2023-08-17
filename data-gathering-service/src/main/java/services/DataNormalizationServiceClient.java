package services;

import org.springframework.stereotype.Component;

@Component
public class DataNormalizationServiceClient {

    private final DataNormalizationService normalizationService;

    public DataNormalizationServiceClient(DataNormalizationService normalizationService) {
        this.normalizationService = normalizationService;
    }

    // Metodo per inviare i messaggi al Data Normalization Service
    public void processReceivedMessage(String mqttMessage) {
        // Passa il messaggio ricevuto al DataNormalizationService per la normalizzazione
        normalizationService.normalizePayload(mqttMessage);
    }
}