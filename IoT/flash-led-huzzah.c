#include <Arduino.h>

int LED_Pin = 0;
int status = 1;

void setup(){
    pinMode(LED_Pin, OUTPUT);
}

void loop() {

  // put your main code here, to run repeatedly:

if (status) {
    digitalWrite(LED_Pin, HIGH);
} else {
    digitalWrite(LED_Pin, LOW);
}

status = 1-status;
delay (1000);

}
