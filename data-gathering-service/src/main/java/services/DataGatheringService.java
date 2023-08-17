package services;

import org.eclipse.paho.client.mqttv3.IMqttClient;
import org.eclipse.paho.client.mqttv3.IMqttMessageListener;
import org.eclipse.paho.client.mqttv3.MqttException;
import org.eclipse.paho.client.mqttv3.MqttMessage;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class DataGatheringService {
    private final IMqttClient mqttClient;
    private final DataNormalizationServiceClient normalizationServiceClient;

    @Autowired
    public DataGatheringService(IMqttClient mqttClient, DataNormalizationServiceClient normalizationServiceClient) {
        this.mqttClient = mqttClient;
        this.normalizationServiceClient = normalizationServiceClient;
    }

    public void startMqttSubscription() {
        try {
            mqttClient.connect();

            // Sottoscrivi al topic MQTT
            String topic = "mqtt/topic"; // Sostituisci con il tuo topic MQTT
            mqttClient.subscribe(topic, new MqttMessageListener());
        } catch (MqttException e) {
            // Gestisci l'eccezione in caso di errore di connessione o sottoscrizione
            e.printStackTrace();
        }
    }

    // Listener per i messaggi MQTT ricevuti
    private class MqttMessageListener implements IMqttMessageListener {
        @Override
        public void messageArrived(String topic, MqttMessage message) throws Exception {
            String mqttMessage = new String(message.getPayload());

            // Invia il messaggio ricevuto al DataNormalizationServiceClient
            normalizationServiceClient.processReceivedMessage(mqttMessage);
        }
    }
}
