"""
Contoh-contoh penggunaan program pembaca timbangan
"""

# ============================================================
# CONTOH 1: Program Paling Sederhana
# ============================================================
import serial
import time

def simple_read():
    """Contoh paling sederhana membaca timbangan"""
    # Ganti 'COM3' dengan port Anda
    ser = serial.Serial(
        port='COM3',
        baudrate=9600,
        bytesize=8,
        parity='N',
        stopbits=1,
        timeout=1
    )
    
    print("Membaca data timbangan...")
    print("Tekan Ctrl+C untuk berhenti\n")
    
    try:
        while True:
            if ser.in_waiting > 0:
                data = ser.readline().decode('ascii').strip()
                print(f"Data: {data}")
    except KeyboardInterrupt:
        print("\nSelesai")
    finally:
        ser.close()


# ============================================================
# CONTOH 2: Dengan Parsing Data
# ============================================================
import re

def read_with_parsing():
    """Membaca dan parsing data"""
    ser = serial.Serial('COM3', 9600, timeout=1)
    
    print("Membaca dengan parsing...\n")
    
    try:
        while True:
            if ser.in_waiting > 0:
                raw = ser.readline().decode('ascii').strip()
                
                # Parse format: ST,GS,+0210.000  g
                match = re.match(r'([A-Z]{2}),([A-Z]{2}),([+-]?\d+\.?\d*)\s*([a-zA-Z]+)', raw)
                
                if match:
                    status = match.group(1)
                    weight = float(match.group(3))
                    unit = match.group(4)
                    
                    stable = "✓" if status == "ST" else "✗"
                    print(f"{stable} Berat: {weight:>10.3f} {unit}")
                    
    except KeyboardInterrupt:
        print("\nSelesai")
    finally:
        ser.close()


# ============================================================
# CONTOH 3: Simpan ke File
# ============================================================
from datetime import datetime

def save_to_file():
    """Membaca dan simpan ke file"""
    ser = serial.Serial('COM3', 9600, timeout=1)
    filename = f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    print(f"Menyimpan data ke {filename}...")
    print("Tekan Ctrl+C untuk berhenti\n")
    
    with open(filename, 'w') as f:
        f.write("Timestamp,Status,Weight,Unit\n")
        
        try:
            while True:
                if ser.in_waiting > 0:
                    raw = ser.readline().decode('ascii').strip()
                    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    
                    # Parse data
                    match = re.match(r'([A-Z]{2}),([A-Z]{2}),([+-]?\d+\.?\d*)\s*([a-zA-Z]+)', raw)
                    if match:
                        status = match.group(1)
                        weight = match.group(3)
                        unit = match.group(4)
                        
                        line = f"{timestamp},{status},{weight},{unit}\n"
                        f.write(line)
                        print(f"✓ {timestamp} - {weight} {unit}")
                        
        except KeyboardInterrupt:
            print(f"\n✓ Data tersimpan di {filename}")
        finally:
            ser.close()


# ============================================================
# CONTOH 4: Monitoring dengan Threshold
# ============================================================
def monitor_with_alert():
    """Monitor berat dengan peringatan jika melebihi threshold"""
    ser = serial.Serial('COM3', 9600, timeout=1)
    
    # Set threshold
    MIN_WEIGHT = 100.0
    MAX_WEIGHT = 500.0
    
    print(f"Monitoring berat...")
    print(f"Alert jika < {MIN_WEIGHT}g atau > {MAX_WEIGHT}g\n")
    
    try:
        while True:
            if ser.in_waiting > 0:
                raw = ser.readline().decode('ascii').strip()
                match = re.match(r'([A-Z]{2}),([A-Z]{2}),([+-]?\d+\.?\d*)\s*([a-zA-Z]+)', raw)
                
                if match:
                    status = match.group(1)
                    weight = float(match.group(3))
                    unit = match.group(4)
                    
                    # Cek threshold
                    if weight < MIN_WEIGHT:
                        print(f"⚠️  TERLALU RINGAN: {weight} {unit}")
                    elif weight > MAX_WEIGHT:
                        print(f"⚠️  TERLALU BERAT: {weight} {unit}")
                    else:
                        print(f"✓ OK: {weight} {unit}")
                        
    except KeyboardInterrupt:
        print("\nSelesai")
    finally:
        ser.close()


# ============================================================
# CONTOH 5: Rata-rata dari N Pembacaan
# ============================================================
def average_reading(n_samples=10):
    """Ambil rata-rata dari N pembacaan"""
    ser = serial.Serial('COM3', 9600, timeout=1)
    
    readings = []
    print(f"Mengambil {n_samples} pembacaan untuk rata-rata...\n")
    
    try:
        while len(readings) < n_samples:
            if ser.in_waiting > 0:
                raw = ser.readline().decode('ascii').strip()
                match = re.match(r'([A-Z]{2}),([A-Z]{2}),([+-]?\d+\.?\d*)\s*([a-zA-Z]+)', raw)
                
                if match and match.group(1) == 'ST':  # Hanya ambil data stabil
                    weight = float(match.group(3))
                    unit = match.group(4)
                    readings.append(weight)
                    print(f"{len(readings)}. {weight} {unit}")
        
        # Hitung rata-rata
        average = sum(readings) / len(readings)
        print(f"\n✓ Rata-rata: {average:.3f} {unit}")
        print(f"  Min: {min(readings):.3f} {unit}")
        print(f"  Max: {max(readings):.3f} {unit}")
        
    finally:
        ser.close()


# ============================================================
# CONTOH 6: Kirim Command ke Timbangan
# ============================================================
def send_commands():
    """Mengirim command ke timbangan"""
    ser = serial.Serial('COM3', 9600, timeout=1)
    time.sleep(0.5)  # Tunggu koneksi stabil
    
    # Command untuk AND GX series:
    # Q  - Request data immediately
    # R  - Re-zero
    # CAL - Calibration
    # ON - Display ON
    # OFF - Display OFF
    
    print("Mengirim command 'Q' untuk request data...")
    ser.write(b'Q\r\n')
    time.sleep(0.1)
    
    if ser.in_waiting > 0:
        response = ser.readline().decode('ascii').strip()
        print(f"Response: {response}")
    
    print("\nMengirim command 'R' untuk re-zero...")
    ser.write(b'R\r\n')
    time.sleep(0.5)
    
    print("✓ Selesai")
    ser.close()


# ============================================================
# CONTOH 7: Real-time Graph (memerlukan matplotlib)
# ============================================================
def plot_realtime():
    """Menampilkan grafik real-time (perlu install matplotlib)"""
    try:
        import matplotlib.pyplot as plt
        import matplotlib.animation as animation
    except ImportError:
        print("Install matplotlib terlebih dahulu: pip install matplotlib")
        return
    
    ser = serial.Serial('COM3', 9600, timeout=1)
    
    fig, ax = plt.subplots()
    x_data, y_data = [], []
    line, = ax.plot(x_data, y_data)
    
    ax.set_xlabel('Sample')
    ax.set_ylabel('Weight (g)')
    ax.set_title('Real-time Weight Monitor')
    
    def update(frame):
        if ser.in_waiting > 0:
            raw = ser.readline().decode('ascii').strip()
            match = re.match(r'([A-Z]{2}),([A-Z]{2}),([+-]?\d+\.?\d*)\s*([a-zA-Z]+)', raw)
            
            if match and match.group(1) == 'ST':
                weight = float(match.group(3))
                x_data.append(len(x_data))
                y_data.append(weight)
                
                # Keep only last 50 points
                if len(x_data) > 50:
                    x_data.pop(0)
                    y_data.pop(0)
                
                line.set_data(x_data, y_data)
                ax.relim()
                ax.autoscale_view()
        
        return line,
    
    ani = animation.FuncAnimation(fig, update, interval=100, blit=True)
    plt.show()
    ser.close()


# ============================================================
# Main Menu untuk Testing
# ============================================================
if __name__ == "__main__":
    print("="*60)
    print("Contoh Program Pembaca Timbangan")
    print("="*60)
    print("\nPilih contoh yang ingin dijalankan:")
    print("1. Program paling sederhana")
    print("2. Dengan parsing data")
    print("3. Simpan ke file")
    print("4. Monitor dengan alert threshold")
    print("5. Rata-rata dari N pembacaan")
    print("6. Kirim command ke timbangan")
    print("7. Real-time graph")
    print("\nCATATAN: Ganti 'COM3' dengan port Anda di kode!")
    print("\nPilihan (1-7): ", end="")
    
    choice = input().strip()
    
    if choice == '1':
        simple_read()
    elif choice == '2':
        read_with_parsing()
    elif choice == '3':
        save_to_file()
    elif choice == '4':
        monitor_with_alert()
    elif choice == '5':
        n = int(input("Jumlah sample (default 10): ") or 10)
        average_reading(n)
    elif choice == '6':
        send_commands()
    elif choice == '7':
        plot_realtime()
    else:
        print("Pilihan tidak valid")