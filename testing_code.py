import time

import serial

BAUD_RATE = 9600
SERIAL_PORT = "COM3"

ser = serial.Serial(SERIAL_PORT, BAUD_RATE, serial.EIGHTBITS, serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE, timeout=5)
print("Serial port open:", ser.is_open)
print(ser.read())
ser.write(f"1 2 3 3 5 6\n".encode())
time.sleep(1)
print("Closing")
ser.close()
