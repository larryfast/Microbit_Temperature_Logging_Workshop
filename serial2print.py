import serial
import time
import re

ser = serial.Serial('/dev/ttyACM0', 115200, timeout = 1)

while True:
    data = ser.readline().decode('utf-8')
    #data = ser.read()
    d1 = re.sub(' ','.',data)
    d2 = data.maketrans(" ",".")
    print(f'{d1}', end='')


