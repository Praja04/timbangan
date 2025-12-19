import serial
import time

# Ganti dengan port Anda
PORT = 'COM11'
BAUDRATE = 9600

print("Testing format timbangan...")
print("Letakkan beban 29 gram di timbangan")
print("-" * 50)

ser = serial.Serial(PORT, BAUDRATE, timeout=1)
time.sleep(0.5)

print("\nData yang diterima (10 sample):\n")

for i in range(10):
    if ser.in_waiting > 0:
        raw = ser.readline()
        
        # Tampilkan dalam berbagai format
        print(f"Sample {i+1}:")
        print(f"  HEX   : {raw.hex()}")
        print(f"  ASCII : '{raw.decode('ascii', errors='replace').strip()}'")
        print(f"  Bytes : {list(raw)}")
        print()
        
    time.sleep(0.2)

ser.close()

print("-" * 50)
print("\nAnalisis:")
print("- Apakah ada koma (,) di data?")
print("- Apakah formatnya ST,GS,+xxxx.xx g ?")
print("- Jika tidak, setting format timbangan masih salah!")