"""
Program Pembaca Timbangan AND GX-4000
Menggunakan Python dengan PySerial
"""

import serial
import serial.tools.list_ports
import time
import csv
from datetime import datetime
import re

class ScaleReader:
    def __init__(self, port=None, baudrate=9600, timeout=1):
        """
        Inisialisasi pembaca timbangan
        
        Args:
            port: COM port (contoh: 'COM3' di Windows, '/dev/ttyUSB0' di Linux)
            baudrate: Kecepatan komunikasi (default: 9600)
            timeout: Timeout dalam detik (default: 1)
        """
        self.port = port
        self.baudrate = baudrate
        self.timeout = timeout
        self.serial_conn = None
        self.is_connected = False
        self.data_history = []
        
    def list_available_ports(self):
        """Menampilkan daftar COM port yang tersedia"""
        ports = serial.tools.list_ports.comports()
        available_ports = []
        
        print("\n=== Daftar COM Port Tersedia ===")
        for i, port in enumerate(ports, 1):
            print(f"{i}. {port.device} - {port.description}")
            available_ports.append(port.device)
        
        return available_ports
    
    def connect(self):
        """Membuka koneksi serial ke timbangan"""
        try:
            self.serial_conn = serial.Serial(
                port=self.port,
                baudrate=self.baudrate,
                bytesize=serial.SEVENBITS,
                parity=serial.PARITY_EVEN,
                stopbits=serial.STOPBITS_ONE,
                timeout=self.timeout
            )
            
            # Bersihkan buffer
            self.serial_conn.reset_input_buffer()
            self.serial_conn.reset_output_buffer()
            
            self.is_connected = True
            print(f"✓ Terhubung ke {self.port} pada {self.baudrate} bps")
            return True
            
        except serial.SerialException as e:
            print(f"✗ Error koneksi: {e}")
            self.is_connected = False
            return False
    
    def disconnect(self):
        """Menutup koneksi serial"""
        if self.serial_conn and self.serial_conn.is_open:
            self.serial_conn.close()
            self.is_connected = False
            print("✓ Koneksi ditutup")
    
    def parse_data(self, raw_data):
        """
        Parsing data dari timbangan
        Format A&D Standard: ST,GS,+0210.000  g  
        
        Returns:
            dict: {'status': 'ST', 'type': 'GS', 'weight': 210.000, 'unit': 'g', 'raw': '...'}
        """
        try:
            # Bersihkan whitespace
            data = raw_data.strip()
            
            if not data:
                return None
            
            # Pattern untuk A&D standard format
            # Format: ST,GS,+0210.000  g
            pattern = r'([A-Z]{2}),([A-Z]{2}),([+-]?\d+\.?\d*)\s*([a-zA-Z]+)'
            match = re.match(pattern, data)
            
            if match:
                status = match.group(1)  # ST = Stable, US = Unstable
                weight_type = match.group(2)  # GS = Gross, NT = Net
                weight = float(match.group(3))
                unit = match.group(4)
                
                return {
                    'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'status': status,
                    'type': weight_type,
                    'weight': weight,
                    'unit': unit,
                    'raw': data
                }
            else:
                # Coba format sederhana (hanya angka)
                simple_pattern = r'([+-]?\d+\.?\d*)'
                simple_match = re.search(simple_pattern, data)
                if simple_match:
                    return {
                        'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        'status': 'UK',  # Unknown
                        'type': 'UK',
                        'weight': float(simple_match.group(1)),
                        'unit': 'g',
                        'raw': data
                    }
                
        except Exception as e:
            print(f"Error parsing data: {e}")
            return None
    
    def read_continuous(self, callback=None, duration=None):
        """
        Membaca data secara kontinyu
        
        Args:
            callback: Function yang dipanggil setiap ada data baru
            duration: Durasi pembacaan dalam detik (None = unlimited)
        """
        if not self.is_connected:
            print("✗ Belum terhubung ke timbangan")
            return
        
        print("\n=== Mulai Membaca Data (Tekan Ctrl+C untuk berhenti) ===\n")
        
        start_time = time.time()
        
        try:
            while True:
                # Cek durasi
                if duration and (time.time() - start_time) > duration:
                    break
                
                # Baca data dari serial
                if self.serial_conn.in_waiting > 0:
                    raw_data = self.serial_conn.readline().decode('ascii', errors='ignore')
                    
                    # Parse data
                    parsed = self.parse_data(raw_data)
                    
                    if parsed:
                        # Simpan ke history
                        self.data_history.append(parsed)
                        
                        # Tampilkan
                        print(f"[{parsed['timestamp']}] "
                              f"{parsed['status']} | "
                              f"Berat: {parsed['weight']:>10.3f} {parsed['unit']} | "
                              f"Status: {'Stabil' if parsed['status'] == 'ST' else 'Tidak Stabil'}")
                        
                        # Callback jika ada
                        if callback:
                            callback(parsed)
                
                time.sleep(0.01)  # Delay kecil untuk CPU
                
        except KeyboardInterrupt:
            print("\n\n✓ Pembacaan dihentikan oleh user")
    
    def read_single(self):
        """Membaca satu data saja"""
        if not self.is_connected:
            print("✗ Belum terhubung ke timbangan")
            return None
        
        try:
            if self.serial_conn.in_waiting > 0:
                raw_data = self.serial_conn.readline().decode('ascii', errors='ignore')
                return self.parse_data(raw_data)
        except Exception as e:
            print(f"Error membaca data: {e}")
            return None
    
    def save_to_csv(self, filename=None):
        """Menyimpan history data ke file CSV"""
        if not self.data_history:
            print("✗ Tidak ada data untuk disimpan")
            return
        
        if not filename:
            filename = f"timbangan_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=['timestamp', 'status', 'type', 'weight', 'unit', 'raw'])
                writer.writeheader()
                writer.writerows(self.data_history)
            
            print(f"✓ Data berhasil disimpan ke {filename} ({len(self.data_history)} baris)")
        except Exception as e:
            print(f"✗ Error menyimpan file: {e}")
    
    def get_statistics(self):
        """Menghitung statistik dari data yang terbaca"""
        if not self.data_history:
            print("✗ Tidak ada data untuk dihitung")
            return None
        
        weights = [d['weight'] for d in self.data_history]
        
        stats = {
            'count': len(weights),
            'min': min(weights),
            'max': max(weights),
            'avg': sum(weights) / len(weights),
            'range': max(weights) - min(weights)
        }
        
        # Hitung standard deviation
        avg = stats['avg']
        variance = sum((x - avg) ** 2 for x in weights) / len(weights)
        stats['std_dev'] = variance ** 0.5
        
        return stats
    
    def print_statistics(self):
        """Menampilkan statistik"""
        stats = self.get_statistics()
        if stats:
            print("\n=== Statistik Data ===")
            print(f"Jumlah Data  : {stats['count']}")
            print(f"Minimum      : {stats['min']:.3f}")
            print(f"Maximum      : {stats['max']:.3f}")
            print(f"Rata-rata    : {stats['avg']:.3f}")
            print(f"Range        : {stats['range']:.3f}")
            print(f"Std Deviasi  : {stats['std_dev']:.3f}")


def main():
    """Fungsi utama program"""
    print("="*60)
    print("Program Pembaca Timbangan AND GX-4000")
    print("="*60)
    
    # Buat instance ScaleReader
    reader = ScaleReader()
    
    # Tampilkan port yang tersedia
    ports = reader.list_available_ports()
    
    if not ports:
        print("\n✗ Tidak ada COM port yang tersedia")
        return
    
    # Pilih port
    print("\nMasukkan nomor port atau nama port (contoh: COM3):")
    port_input = input("> ").strip()
    
    # Cek apakah input adalah nomor atau nama port
    if port_input.isdigit():
        port_idx = int(port_input) - 1
        if 0 <= port_idx < len(ports):
            selected_port = ports[port_idx]
        else:
            print("✗ Nomor port tidak valid")
            return
    else:
        selected_port = port_input
    
    # Set port dan baud rate
    reader.port = selected_port
    
    print("\nMasukkan baud rate (default: 9600):")
    baudrate_input = input("> ").strip()
    if baudrate_input:
        try:
            reader.baudrate = int(baudrate_input)
        except ValueError:
            print("✗ Baud rate tidak valid, menggunakan 9600")
    
    # Koneksi ke timbangan
    if not reader.connect():
        return
    
    # Menu utama
    while True:
        try:
            print("\n" + "="*60)
            print("Menu:")
            print("1. Baca data kontinyu")
            print("2. Baca satu data")
            print("3. Simpan data ke CSV")
            print("4. Tampilkan statistik")
            print("5. Hapus history data")
            print("6. Keluar")
            print("="*60)
            
            choice = input("Pilih menu (1-6): ").strip()
            
            if choice == '1':
                try:
                    print("\nMasukkan durasi (detik) atau kosongkan untuk unlimited:")
                    duration_input = input("> ").strip()
                    duration = int(duration_input) if duration_input else None
                    
                    reader.read_continuous(duration=duration)
                except ValueError:
                    print("✗ Durasi harus berupa angka")
                    
            elif choice == '2':
                print("\nMembaca satu data...")
                data = reader.read_single()
                if data:
                    print(f"\nData terbaca:")
                    print(f"  Waktu   : {data['timestamp']}")
                    print(f"  Status  : {data['status']}")
                    print(f"  Berat   : {data['weight']} {data['unit']}")
                    print(f"  Raw     : {data['raw']}")
                else:
                    print("✗ Tidak ada data")
                
            elif choice == '3':
                print("\nMasukkan nama file (kosongkan untuk auto):")
                filename = input("> ").strip()
                reader.save_to_csv(filename if filename else None)
                
            elif choice == '4':
                reader.print_statistics()
                
            elif choice == '5':
                reader.data_history = []
                print("✓ History data telah dihapus")
                
            elif choice == '6':
                break
            
            else:
                print("✗ Pilihan tidak valid")
                
        except KeyboardInterrupt:
            print("\n\n✓ Program dihentikan oleh user")
            break
        except Exception as e:
            print(f"\n✗ Error: {e}")
            print("Program akan dilanjutkan...")
    
    # Tutup koneksi
    reader.disconnect()
    print("\nProgram selesai. Terima kasih!")


if __name__ == "__main__":
    main()