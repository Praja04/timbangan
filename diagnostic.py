"""
COM Port Diagnostic Tool
Untuk troubleshooting masalah port COM
"""
import serial
import serial.tools.list_ports
import sys
import time

def test_port_basic(port_name):
    """Test basic port opening"""
    print(f"\n{'='*60}")
    print(f"Testing: {port_name}")
    print(f"{'='*60}")
    
    try:
        print("1. Trying basic open...")
        ser = serial.Serial(port_name, timeout=1)
        print("   ✓ Basic open SUCCESS")
        ser.close()
        time.sleep(0.5)
    except Exception as e:
        print(f"   ✗ Basic open FAILED: {e}")
        return False
    
    try:
        print("\n2. Trying with 9600 baud...")
        ser = serial.Serial(port_name, 9600, timeout=1)
        print("   ✓ 9600 baud SUCCESS")
        ser.close()
        time.sleep(0.5)
    except Exception as e:
        print(f"   ✗ 9600 baud FAILED: {e}")
        return False
    
    try:
        print("\n3. Trying with full config (7-E-1)...")
        ser = serial.Serial(
            port=port_name,
            baudrate=9600,
            bytesize=7,
            parity=serial.PARITY_EVEN,
            stopbits=1,
            timeout=1,
            xonxoff=False,
            rtscts=False,
            dsrdtr=False
        )
        print("   ✓ Full config SUCCESS")
        
        print("\n4. Setting DTR/RTS...")
        ser.setDTR(False)
        ser.setRTS(False)
        print("   ✓ DTR/RTS set SUCCESS")
        
        print("\n5. Testing read (5 seconds)...")
        ser.reset_input_buffer()
        start = time.time()
        data_received = False
        
        while time.time() - start < 5:
            if ser.in_waiting > 0:
                data = ser.readline().decode('ascii', errors='ignore').strip()
                if data:
                    print(f"   ✓ Data received: {data}")
                    data_received = True
                    break
            time.sleep(0.1)
        
        if not data_received:
            print("   ⚠ No data received (pastikan timbangan aktif)")
        
        ser.close()
        print("\n6. Closing port...")
        print("   ✓ Port closed SUCCESS")
        
        return True
        
    except Exception as e:
        print(f"   ✗ Full config FAILED: {e}")
        return False

def main():
    print("="*60)
    print("COM PORT DIAGNOSTIC TOOL")
    print("="*60)
    
    # List all ports
    print("\nAvailable COM Ports:")
    ports = list(serial.tools.list_ports.comports())
    
    if not ports:
        print("  ⚠ No COM ports found!")
        print("\nPossible causes:")
        print("  - USB cable not connected")
        print("  - Driver not installed")
        print("  - Device not powered on")
        input("\nPress Enter to exit...")
        return
    
    for i, port in enumerate(ports, 1):
        print(f"\n  {i}. {port.device}")
        print(f"     Description: {port.description}")
        print(f"     Hardware ID: {port.hwid}")
        
        # Check if port is in use
        try:
            ser = serial.Serial(port.device, timeout=0.1)
            ser.close()
            print(f"     Status: ✓ Available")
        except serial.SerialException as e:
            if "PermissionError" in str(e) or "Access is denied" in str(e):
                print(f"     Status: ✗ IN USE by another application")
            else:
                print(f"     Status: ✗ Error: {e}")
    
    # Ask user to select port for detailed test
    print("\n" + "="*60)
    try:
        choice = input("\nSelect port number to test (or Enter to skip): ").strip()
        if choice:
            idx = int(choice) - 1
            if 0 <= idx < len(ports):
                selected_port = ports[idx].device
                success = test_port_basic(selected_port)
                
                if success:
                    print("\n" + "="*60)
                    print("✓✓✓ PORT TEST PASSED ✓✓✓")
                    print("="*60)
                    print("\nPort berfungsi dengan baik!")
                    print("Jika aplikasi masih error, coba:")
                    print("  1. Jalankan aplikasi sebagai Administrator")
                    print("  2. Disable antivirus sementara")
                    print("  3. Restart komputer")
                else:
                    print("\n" + "="*60)
                    print("✗✗✗ PORT TEST FAILED ✗✗✗")
                    print("="*60)
                    print("\nSaran:")
                    print("  1. Cabut dan colok ulang kabel USB")
                    print("  2. Coba port USB lain")
                    print("  3. Install/update driver")
                    print("  4. Check Device Manager untuk errors")
            else:
                print("Invalid selection")
    except ValueError:
        print("Invalid input")
    except KeyboardInterrupt:
        print("\n\nTest cancelled by user")
    
    input("\n\nPress Enter to exit...")

if __name__ == "__main__":
    main()