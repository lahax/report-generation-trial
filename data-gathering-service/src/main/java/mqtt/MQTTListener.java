package mqtt;

import org.eclipse.paho.client.mqttv3.MqttCallback;
import org.eclipse.paho.client.mqttv3.MqttMessage;
import org.springframework.stereotype.Component;

@Component
public class MQTTListener implements MqttCallback {
    @Override
    public void connectionLost(Throwable cause) {
        // Logica in caso di perdita di connessione al broker
    }

    @Override
    public void messageArrived(String topic, MqttMessage message) throws Exception {
        // Qui puoi inviare il messaggio al Data Normalization Service
    }

    @Override
    public void deliveryComplete(org.eclipse.paho.client.mqttv3.IMqttDeliveryToken token) {
        // Logica quando la consegna del messaggio è completata
    }
}
