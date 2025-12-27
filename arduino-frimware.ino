#define VOLTAGE_PIN A0
#define CURRENT_PIN A1
#define FREQ_PIN 2

// calibration values - adjust these for your setup
const float V_REF = 5.0;
const float V_DIVIDER = 11.0;  // 100k/10k divider ratio
const float I_SENSITIVITY = 0.185;  // ACS712-5A
const float I_OFFSET = 2.5;

volatile unsigned long pulses = 0;
unsigned long lastCalc = 0;
float freq = 0;

// sampling buffers
const int BUF_SIZE = 100;
float vBuffer[BUF_SIZE];
float iBuffer[BUF_SIZE];
int bufIdx = 0;

void countPulse() {
    pulses++;
}

void setup() {
    Serial.begin(115200);
    
    pinMode(VOLTAGE_PIN, INPUT);
    pinMode(CURRENT_PIN, INPUT);
    pinMode(FREQ_PIN, INPUT);
    
    attachInterrupt(digitalPinToInterrupt(FREQ_PIN), countPulse, RISING);
    
    // wait a bit for things to settle
    delay(100);
    Serial.println("BOARD_TESTER_READY");
}

float readVoltage() {
    long total = 0;
    for (int i = 0; i < 10; i++) {
        total += analogRead(VOLTAGE_PIN);
        delayMicroseconds(100);
    }
    float avg = total / 10.0;
    return (avg / 1023.0) * V_REF * V_DIVIDER;
}

float readCurrent() {
    long total = 0;
    for (int i = 0; i < 10; i++) {
        total += analogRead(CURRENT_PIN);
        delayMicroseconds(100);
    }
    float avg = total / 10.0;
    float v = (avg / 1023.0) * V_REF;
    return (v - I_OFFSET) / I_SENSITIVITY;
}

float calcRMS(float* buf, int len) {
    float sum = 0;
    for (int i = 0; i < len; i++) {
        sum += buf[i] * buf[i];
    }
    return sqrt(sum / len);
}

float calcPeakToPeak(float* buf, int len) {
    float minV = buf[0];
    float maxV = buf[0];
    for (int i = 1; i < len; i++) {
        if (buf[i] < minV) minV = buf[i];
        if (buf[i] > maxV) maxV = buf[i];
    }
    return maxV - minV;
}

void sendData() {
    float v = readVoltage();
    float i = readCurrent();
    
    // store in buffer
    vBuffer[bufIdx] = v;
    iBuffer[bufIdx] = i;
    bufIdx = (bufIdx + 1) % BUF_SIZE;
    
    float p = v * i;
    float r = (i != 0) ? (v / i) : 0;
    float vRms = calcRMS(vBuffer, BUF_SIZE);
    float vPp = calcPeakToPeak(vBuffer, BUF_SIZE);
    
    // wavelength calc (speed of light / freq)
    float wl = (freq > 0) ? (299792458.0 / freq) : 0;
    
    // send as json
    Serial.print("{\"V\":");
    Serial.print(v, 3);
    Serial.print(",\"I\":");
    Serial.print(i, 4);
    Serial.print(",\"P\":");
    Serial.print(p, 3);
    Serial.print(",\"R\":");
    Serial.print(r, 2);
    Serial.print(",\"F\":");
    Serial.print(freq, 1);
    Serial.print(",\"WL\":");
    Serial.print(wl, 2);
    Serial.print(",\"Vrms\":");
    Serial.print(vRms, 3);
    Serial.print(",\"Vpp\":");
    Serial.print(vPp, 3);
    Serial.println("}");
}

void handleSerial() {
    if (Serial.available()) {
        String cmd = Serial.readStringUntil('\n');
        cmd.trim();
        
        if (cmd == "PING") {
            Serial.println("PONG");
        }
        else if (cmd == "INFO") {
            Serial.println("{\"device\":\"Board Tester\",\"version\":\"3.0\"}");
        }
        else if (cmd == "RESET") {
            bufIdx = 0;
            pulses = 0;
            Serial.println("OK");
        }
    }
}

void loop() {
    // calculate frequency every second
    if (millis() - lastCalc >= 1000) {
        noInterrupts();
        freq = pulses;
        pulses = 0;
        interrupts();
        lastCalc = millis();
    }
    
    // send data every 50ms
    static unsigned long lastSend = 0;
    if (millis() - lastSend >= 50) {
        sendData();
        lastSend = millis();
    }
    
    handleSerial();
}