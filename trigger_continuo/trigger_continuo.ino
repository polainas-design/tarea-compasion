// ============================================================================
// ARDUINO UNO - TRIGGER START/STOP GRABACIÓN CONTINUA PARA DELSYS TRIGNO EMG
// ============================================================================
//
// Modo de operación:
//   - Comando 'H': pulso de 10ms en Pin 2 (START grabación continua)
//   - Comando 'L': pulso de 10ms en Pin 3 (STOP grabación continua)
//   - Se envía UN SOLO START al inicio de la tarea y UN SOLO STOP al final
//   - Los eventos individuales se registran mediante timestamps en PsychoPy
//
// CONEXIONES:
//   Pin 2  --> centro BNC 1 (configurar como Start Trigger en Trigno Discover)
//   Pin 3  --> centro BNC 2 (configurar como Stop Trigger en Trigno Discover)
//   GND    --> malla de ambos BNC
//
// ============================================================================

const int PIN_START = 2;
const int PIN_STOP = 3;
const long BAUD_RATE = 115200;

const unsigned long PULSE_WIDTH_US = 10000;  // 10ms

void setup() {
  pinMode(PIN_START, OUTPUT);
  pinMode(PIN_STOP, OUTPUT);
  digitalWrite(PIN_START, LOW);
  digitalWrite(PIN_STOP, LOW);

  Serial.begin(BAUD_RATE);

  // Sin pulso de prueba para evitar disparar el Trigno al encender
  Serial.println("READY");
}

void loop() {
  if (Serial.available() > 0) {
    int code = Serial.read();

    // Ignorar caracteres de control
    if (code <= 0 || code == 10 || code == 13) {
      return;
    }

    if (code == 'H') {
      // START: pulso en Pin 2
      digitalWrite(PIN_START, HIGH);
      unsigned long t_on = micros();
      delayMicroseconds(PULSE_WIDTH_US);
      digitalWrite(PIN_START, LOW);
      Serial.print("START:T");
      Serial.println(t_on);
    }
    else if (code == 'L') {
      // STOP: pulso en Pin 3
      digitalWrite(PIN_STOP, HIGH);
      unsigned long t_off = micros();
      delayMicroseconds(PULSE_WIDTH_US);
      digitalWrite(PIN_STOP, LOW);
      Serial.print("STOP:T");
      Serial.println(t_off);
    }
  }
}
