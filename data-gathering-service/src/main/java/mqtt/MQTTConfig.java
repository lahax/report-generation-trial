package mqtt;

import org.eclipse.paho.client.mqttv3.IMqttClient;
import org.eclipse.paho.client.mqttv3.MqttClient;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

@Configuration
public class MQTTConfig {

    @Autowired
    private MQTTListener mqttListener;

    @Bean
    public IMqttClient mqttClient() throws Exception {
        IMqttClient client = new MqttClient("tcp://broker-url:port", "unique-client-id");
        client.setCallback(mqttListener);
        client.connect();
        client.subscribe("topic-to-subscribe");
        // Puoi aggiungere altre sottoscrizioni se necessario
        return client;
    }
}
