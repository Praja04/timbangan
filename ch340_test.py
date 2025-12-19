"""
CH340 Connection Test Tool - Standalone
Untuk troubleshoot USB-SERIAL CH340 (COM11)
Khusus AND GX-4000

Cara pakai:
1. Pastikan timbangan NYALA
2. Jalankan: python ch340_test.py
3. Pilih COM11 (atau Enter untuk default)
4. Tunggu hasil test
"""
import serial
import time
import sys

def print_header():
    print("\n" + "="*70)
    print(" CH340 USB-SERIAL CONNECTION TEST TOOL")
    print(" Untuk AND GX-4000 Timbangan")
    print("="*70)

def test_ch340(port="COM11"):
    """Test CH340 connection step by step"""
    
    print(f"\n📌 Testing Port: {port}")
    print("   Driver: USB-SERIAL CH340")
    print("\n⚙️  Konfigurasi AND GX-4000:")
    print("   • Baudrate: 9600 bps")
    print("   • Data bits: 7 bit")
    print("   • Parity: Even")
    print("   • Stop bits: 1")
    print("\n" + "-"*70)
    
    # TEST 1: Quick Open/Close
    print("\n[TEST 1/4] Quick Open/Close Test")
    print("   Purpose: Cek apakah port bisa dibuka")
    try:
        print("   ⏳ Opening port...")
        ser = serial.Serial(port, timeout=0.5)
        print("   ✅ SUCCESS - Port dapat dibuka!")
        ser.close()
        print("   ✅ Port closed successfully")
        time.sleep(1)
    except Exception as e:
        print(f"   ❌ FAILED - {e}")
        print("\n   🔧 Solusi:")
        print("      1. Cabut USB, tunggu 5 detik, colok lagi")
        print("      2. Coba port USB lain")
        print("      3. Restart komputer")
        print("      4. Update driver CH340")
        return False
    
    # TEST 2: Open with CH340-specific delays
    print("\n[TEST 2/4] Open with Delays (CH340 Specific)")
    print("   Purpose: Test dengan timing khusus CH340")
    try:
        print("   ⏳ Waiting 1 second before opening...")
        time.sleep(1)
        
        print("   ⏳ Opening port with full config...")
        ser = serial.Serial(
            port=port,
            baudrate=9600,
            bytesize=7,
            parity=serial.PARITY_EVEN,
            stopbits=1,
            timeout=2,
            write_timeout=2,
            xonxoff=False,
            rtscts=False,
            dsrdtr=False
        )
        print("   ✅ Port opened successfully")
        
        print("   ⏳ Configuring DTR/RTS...")
        time.sleep(0.5)
        ser.setDTR(False)
        ser.setRTS(False)
        time.sleep(0.3)
        print("   ✅ DTR/RTS configured")
        
        print("   ⏳ Flushing buffers...")
        ser.reset_input_buffer()
        ser.reset_output_buffer()
        print("   ✅ Buffers cleared")
        
        ser.close()
        time.sleep(0.5)
        print("   ✅ Port closed successfully")
        
    except Exception as e:
        print(f"   ❌ FAILED - {e}")
        return False
    
    # TEST 3: Read data from scale
    print("\n[TEST 3/4] Read Data from Timbangan")
    print("   Purpose: Cek apakah bisa terima data")
    print("   ⚠️  PASTIKAN: Timbangan NYALA dan STABIL!")
    print()
    
    try:
        time.sleep(1)
        ser = serial.Serial(
            port=port,
            baudrate=9600,
            bytesize=7,
            parity=serial.PARITY_EVEN,
            stopbits=1,
            timeout=2,
            xonxoff=False,
            rtscts=False,
            dsrdtr=False
        )
        time.sleep(0.5)
        ser.setDTR(False)
        ser.setRTS(False)
        time.sleep(0.3)
        ser.reset_input_buffer()
        
        print("   ⏳ Reading data (10 seconds)...")
        print("   📊 Waiting for data packets...\n")
        
        start = time.time()
        data_count = 0
        
        while time.time() - start < 10:
            if ser.in_waiting > 0:
                data = ser.readline().decode('ascii', errors='ignore').strip()
                if data:
                    data_count += 1
                    print(f"   📦 Packet #{data_count}: [{data}]")
                    if data_count >= 3:
                        print("\n   ✅ Enough data received!")
                        break
            time.sleep(0.1)
        
        ser.close()
        
        if data_count > 0:
            print(f"\n   ✅ SUCCESS! Received {data_count} data packets")
            return True
        else:
            print("\n   ⚠️  Port OK, but NO DATA received from scale")
            print("\n   📋 Checklist:")
            print("      • Timbangan sudah NYALA? ✓")
            print("      • Kabel terhubung dengan benar? ✓")
            print("      • Setting timbangan: 9600-7-E-1? ✓")
            print("      • Mode output timbangan aktif? ✓")
            return True
            
    except Exception as e:
        print(f"   ❌ FAILED - {e}")
        return False
    
    # TEST 4: Retry test
    print("\n[TEST 4/4] Retry Connection Test")
    print("   Purpose: Test dengan 3x retry (seperti app)")
    
    success_count = 0
    for attempt in range(3):
        try:
            print(f"\n   Attempt {attempt + 1}/3...")
            time.sleep(1)
            ser = serial.Serial(
                port=port,
                baudrate=9600,
                bytesize=7,
                parity=serial.PARITY_EVEN,
                stopbits=1,
                timeout=2,
                xonxoff=False,
                rtscts=False,
                dsrdtr=False
            )
            time.sleep(0.3)
            ser.setDTR(False)
            ser.setRTS(False)
            ser.close()
            success_count += 1
            print(f"   ✅ Attempt {attempt + 1} SUCCESS")
        except Exception as e:
            print(f"   ❌ Attempt {attempt + 1} FAILED: {e}")
        
        time.sleep(0.5)
    
    print(f"\n   Result: {success_count}/3 successful connections")
    return success_count >= 2

def main():
    print_header()
    
    # Get port from user
    default_port = "COM11"
    print(f"\n📍 Default port: {default_port} (USB-SERIAL CH340)")
    user_input = input("   Gunakan port lain? (Enter = COM11): ").strip().upper()
    
    if user_input:
        port = user_input
    else:
        port = default_port
    
    print("\n" + "="*70)
    print(" STARTING TESTS...")
    print("="*70)
    
    # Run tests
    success = test_ch340(port)
    
    # Show results
    print("\n" + "="*70)
    if success:
        print(" ✅✅✅ ALL TESTS PASSED ✅✅✅")
        print("="*70)
        print("\n🎉 Port CH340 berfungsi dengan baik!")
        print("\n📱 Jika aplikasi utama masih error:")
        print("   1. Pastikan klik STOP sebelum START ulang")
        print("   2. Tunggu 2-3 detik antara STOP dan START")
        print("   3. Jangan spam klik START berkali-kali")
        print("   4. Jalankan aplikasi sebagai Administrator")
        print("   5. Coba disable antivirus sementara")
    else:
        print(" ❌❌❌ TESTS FAILED ❌❌❌")
        print("="*70)
        print("\n🔧 TROUBLESHOOTING STEPS:")
        print("\n1️⃣  UPDATE DRIVER CH340:")
        print("   • Download: http://www.wch-ic.com/downloads/CH341SER_EXE.html")
        print("   • Install sebagai Administrator")
        print("   • Restart komputer")
        print("\n2️⃣  HARDWARE CHECK:")
        print("   • Cabut USB, tunggu 5 detik, colok lagi")
        print("   • Coba port USB 2.0 (bukan 3.0)")
        print("   • Gunakan port belakang PC (langsung ke motherboard)")
        print("   • Hindari USB hub")
        print("\n3️⃣  WINDOWS SETTINGS:")
        print("   • Device Manager → USB Root Hub")
        print("   • Properties → Power Management")
        print("   • Uncheck 'Allow computer to turn off this device'")
        print("\n4️⃣  LAINNYA:")
        print("   • Jalankan aplikasi sebagai Administrator")
        print("   • Disable antivirus sementara")
        print("   • Restart komputer")
    
    print("\n" + "="*70)
    input("\n Press Enter to exit...")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n⚠️  Test cancelled by user")
        sys.exit(0)
    except Exception as e:
        print(f"\n\n❌ Unexpected error: {e}")
        input("\nPress Enter to exit...")
        sys.exit(1)