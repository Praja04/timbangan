"""
Program untuk mencoba SEMUA kombinasi setting serial
untuk menemukan yang BENAR untuk timbangan AND GX-4000
"""

import serial
import serial.tools.list_ports
import time

def list_ports():
    """List available COM ports"""
    ports = serial.tools.list_ports.comports()
    print("\n" + "="*70)
    print("DAFTAR COM PORT:")
    print("="*70)
    for i, port in enumerate(ports, 1):
        print(f"{i}. {port.device} - {port.description}")
    print("="*70)
    return [p.device for p in ports]

def test_all_combinations(port, duration=5):
    """
    Test SEMUA kombinasi setting untuk menemukan yang benar
    """
    
    # Kombinasi yang akan di-test
    baudrates = [2400, 4800, 9600, 19200]
    databits = [7, 8]
    parities = [
        ('None', serial.PARITY_NONE),
        ('Even', serial.PARITY_EVEN),
        ('Odd', serial.PARITY_ODD)
    ]
    stopbits = [1, 2]
    
    print("\n" + "="*70)
    print("TESTING SEMUA KOMBINASI SETTING SERIAL")
    print("="*70)
    print(f"Port: {port}")
    print(f"Durasi test per kombinasi: {duration} detik")
    print("="*70)
    
    results = []
    total_tests = len(baudrates) * len(databits) * len(parities) * len(stopbits)
    current_test = 0
    
    for baud in baudrates:
        for databit in databits:
            for parity_name, parity_val in parities:
                for stopbit in stopbits:
                    current_test += 1
                    
                    config = {
                        'baudrate': baud,
                        'databits': databit,
                        'parity': parity_name,
                        'stopbits': stopbit
                    }
                    
                    print(f"\n[{current_test}/{total_tests}] Testing: {baud} bps, {databit} data bits, Parity {parity_name}, {stopbit} stop bit(s)")
                    print("-" * 70)
                    
                    try:
                        ser = serial.Serial(
                            port=port,
                            baudrate=baud,
                            bytesize=databit,
                            parity=parity_val,
                            stopbits=stopbit,
                            timeout=1
                        )
                        
                        ser.reset_input_buffer()
                        time.sleep(0.5)
                        
                        # Kumpulkan data selama durasi yang ditentukan
                        samples = []
                        clean_samples = 0
                        corrupted_samples = 0
                        start = time.time()
                        
                        while (time.time() - start) < duration:
                            if ser.in_waiting > 0:
                                try:
                                    raw_bytes = ser.readline()
                                    
                                    # Cek apakah data clean (tidak ada byte > 127)
                                    is_clean = all(b < 128 for b in raw_bytes)
                                    
                                    try:
                                        decoded = raw_bytes.decode('ascii', errors='strict').strip()
                                        
                                        if is_clean and decoded:
                                            samples.append(decoded)
                                            clean_samples += 1
                                            print(f"  ✓ CLEAN: '{decoded}'")
                                        else:
                                            corrupted_samples += 1
                                            
                                    except UnicodeDecodeError:
                                        corrupted_samples += 1
                                        
                                except Exception as e:
                                    corrupted_samples += 1
                            
                            time.sleep(0.01)
                        
                        ser.close()
                        
                        # Evaluasi hasil
                        success_rate = (clean_samples / max(1, clean_samples + corrupted_samples)) * 100
                        
                        result = {
                            'config': config,
                            'clean': clean_samples,
                            'corrupted': corrupted_samples,
                            'success_rate': success_rate,
                            'samples': samples[:3]  # Simpan 3 sample pertama
                        }
                        
                        results.append(result)
                        
                        print(f"  → Clean: {clean_samples}, Corrupt: {corrupted_samples}, Success Rate: {success_rate:.1f}%")
                        
                        # Jika success rate tinggi, tandai sebagai kandidat
                        if success_rate > 80 and clean_samples > 0:
                            print(f"  🌟 KANDIDAT BAGUS! Success rate: {success_rate:.1f}%")
                        
                    except Exception as e:
                        print(f"  ✗ Error: {e}")
                    
                    time.sleep(0.2)  # Jeda antar test
    
    # Tampilkan ringkasan
    print("\n" + "="*70)
    print("RINGKASAN HASIL TESTING")
    print("="*70)
    
    # Sort berdasarkan success rate
    results.sort(key=lambda x: x['success_rate'], reverse=True)
    
    print("\nTOP 5 KONFIGURASI TERBAIK:")
    print("-" * 70)
    
    for i, result in enumerate(results[:5], 1):
        cfg = result['config']
        print(f"\n#{i}. Success Rate: {result['success_rate']:.1f}%")
        print(f"    Config: {cfg['baudrate']} bps, {cfg['databits']} data bits, "
              f"Parity {cfg['parity']}, {cfg['stopbits']} stop bit(s)")
        print(f"    Clean: {result['clean']}, Corrupt: {result['corrupted']}")
        
        if result['samples']:
            print(f"    Sample data:")
            for sample in result['samples']:
                print(f"      - '{sample}'")
    
    # Rekomendasi
    if results and results[0]['success_rate'] > 80:
        best = results[0]['config']
        print("\n" + "="*70)
        print("🎯 REKOMENDASI SETTING:")
        print("="*70)
        print(f"Baud Rate : {best['baudrate']} bps")
        print(f"Data Bits : {best['databits']}")
        print(f"Parity    : {best['parity']}")
        print(f"Stop Bits : {best['stopbits']}")
        print("="*70)
    else:
        print("\n⚠️  Tidak ada konfigurasi dengan success rate tinggi.")
        print("Kemungkinan masalah:")
        print("  1. Kabel RS232 bermasalah")
        print("  2. Setting timbangan perlu diubah")
        print("  3. Port COM salah")

def test_single_config(port, baudrate, databits, parity_str, stopbits, duration=30):
    """Test satu konfigurasi spesifik dengan detail"""
    
    parity_map = {
        'None': serial.PARITY_NONE,
        'Even': serial.PARITY_EVEN,
        'Odd': serial.PARITY_ODD
    }
    
    parity = parity_map.get(parity_str, serial.PARITY_NONE)
    
    print("\n" + "="*70)
    print("TESTING KONFIGURASI SPESIFIK")
    print("="*70)
    print(f"Port      : {port}")
    print(f"Baud Rate : {baudrate} bps")
    print(f"Data Bits : {databits}")
    print(f"Parity    : {parity_str}")
    print(f"Stop Bits : {stopbits}")
    print(f"Durasi    : {duration} detik")
    print("="*70)
    print("\nTekan Ctrl+C untuk berhenti lebih awal\n")
    
    try:
        ser = serial.Serial(
            port=port,
            baudrate=baudrate,
            bytesize=databits,
            parity=parity,
            stopbits=stopbits,
            timeout=1
        )
        
        ser.reset_input_buffer()
        
        count = 0
        clean_count = 0
        corrupt_count = 0
        start = time.time()
        
        while (time.time() - start) < duration:
            if ser.in_waiting > 0:
                raw_bytes = ser.readline()
                count += 1
                
                # Cek byte
                is_clean = all(b < 128 for b in raw_bytes)
                
                # HEX representation
                hex_str = ' '.join(f'{b:02X}' for b in raw_bytes)
                
                try:
                    decoded = raw_bytes.decode('ascii', errors='strict').strip()
                    
                    if is_clean:
                        clean_count += 1
                        print(f"[{count}] ✓ CLEAN: '{decoded}'")
                        print(f"     HEX: {hex_str}")
                    else:
                        corrupt_count += 1
                        print(f"[{count}] ✗ CORRUPT: '{decoded}'")
                        print(f"     HEX: {hex_str}")
                        
                except UnicodeDecodeError:
                    corrupt_count += 1
                    print(f"[{count}] ✗ DECODE ERROR")
                    print(f"     HEX: {hex_str}")
                
                print()
            
            time.sleep(0.01)
        
        ser.close()
        
        print("="*70)
        print("HASIL:")
        print(f"Total data  : {count}")
        print(f"Clean       : {clean_count}")
        print(f"Corrupt     : {corrupt_count}")
        print(f"Success Rate: {(clean_count/max(1,count))*100:.1f}%")
        print("="*70)
        
    except KeyboardInterrupt:
        print("\n\n✓ Test dihentikan oleh user")
        if 'ser' in locals():
            ser.close()
    except Exception as e:
        print(f"\n✗ Error: {e}")

def main():
    print("="*70)
    print("    FIX KOMUNIKASI TIMBANGAN AND GX-4000")
    print("="*70)
    print("\nProgram ini akan mencoba SEMUA kombinasi setting")
    print("untuk menemukan konfigurasi yang BENAR")
    print("="*70)
    
    # List ports
    ports = list_ports()
    if not ports:
        print("\n✗ Tidak ada COM port tersedia")
        return
    
    # Pilih port
    port_input = input("\nPilih nomor port atau ketik nama (contoh: COM3): ").strip()
    
    if port_input.isdigit():
        idx = int(port_input) - 1
        if 0 <= idx < len(ports):
            port = ports[idx]
        else:
            print("✗ Nomor tidak valid")
            return
    else:
        port = port_input
    
    # Menu
    while True:
        print("\n" + "="*70)
        print("MENU:")
        print("="*70)
        print("1. AUTO TEST - Test semua kombinasi (REKOMENDASI)")
        print("2. MANUAL TEST - Test satu konfigurasi spesifik")
        print("3. QUICK TEST - Test kombinasi umum saja")
        print("4. Keluar")
        print("="*70)
        
        choice = input("Pilih (1-4): ").strip()
        
        if choice == '1':
            duration = input("Durasi test per kombinasi (detik, default=5): ").strip()
            dur = int(duration) if duration else 5
            test_all_combinations(port, dur)
            
        elif choice == '2':
            print("\nMasukkan konfigurasi:")
            baud = int(input("  Baud rate (9600): ") or "9600")
            databits = int(input("  Data bits (7/8): ") or "8")
            parity = input("  Parity (None/Even/Odd): ") or "None"
            stopbits = int(input("  Stop bits (1/2): ") or "1")
            duration = int(input("  Durasi test (detik): ") or "30")
            
            test_single_config(port, baud, databits, parity, stopbits, duration)
            
        elif choice == '3':
            # Quick test kombinasi umum
            print("\nTesting kombinasi umum...")
            common_configs = [
                (9600, 8, 'None', 1),
                (9600, 7, 'Even', 1),
                (9600, 8, 'Even', 1),
                (2400, 7, 'Even', 1),
                (2400, 8, 'None', 1),
            ]
            
            for baud, databits, parity, stopbits in common_configs:
                print(f"\n→ Testing: {baud} bps, {databits} bit, {parity}, {stopbits} stop")
                test_single_config(port, baud, databits, parity, stopbits, duration=5)
                
        elif choice == '4':
            break
        else:
            print("✗ Pilihan tidak valid")
    
    print("\n✓ Program selesai")

if __name__ == "__main__":
    main()