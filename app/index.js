const mqtt = require('mqtt');

// Connect to local Mosquitto broker
const client = mqtt.connect('mqtt://localhost:1883');

client.on('connect', () => {
    console.log('✅ Connected to MQTT broker');

    client.subscribe('#', (err) => {
        if (!err) {
            console.log('📡 Subscribed to all topics');
        } else {
            console.error('❌ Subscription error:', err);
        }
    });
});

// Listen for messages
client.on('message', (topic, message) => {
    try {
        const payload = JSON.parse(message.toString());
        if (Array.isArray(payload.data)) {
            payload.data.forEach(item => {
                const mac = item.mac;
                const rssi = item.rssi;
                const deviceInfo  = payload.device_info || {};


                if (mac && rssi !== undefined) {
                    console.log(`DEVICE MAC: ${deviceInfo.mac}|📍 MAC: ${mac} | RSSI: ${rssi}`);
                }
            });
        }

    } catch (err) {
        console.error('❌ Invalid JSON received:', err.message);
    }
});

// Error handling
client.on('error', (error) => {
    console.error('❌ MQTT Error:', error);
});
