#delete this file

import time
import serial

PORT = 'COM3'
BAUD = 9600

def raw_map(data):
    i = 0
    length = len(data)
    
    # Store found headers to print them cleanly
    found = {}
    
    while i < length:
        byte = data[i]
        
        # Check for Header (C0-DF)
        if 0xC0 <= byte <= 0xDF:
            header = f"{byte:02X}"
            
            # Extract raw digits following the header
            digits = ""
            j = i + 1
            while j < length and j < i + 16:
                b = data[j]
                if 0xF0 <= b <= 0xF9:
                    digits += str(b & 0x0F) # Just the number (0-9)
                elif 0xC0 <= b <= 0xDF:
                    break # Stop if we hit the next header
                j += 1
            
            if digits:
                found[header] = digits
            i = j
        else:
            i += 1

    # Print results sorted
    if found:
        print("-" * 40)
        for h, d in found.items():
            print(f"Header [{h}] -> Raw: {d}")

def main():
    try:
        ser = serial.Serial(PORT, BAUD, timeout=1)
        ser.dtr = True; ser.rts = True
        time.sleep(1)
        
        print(f"Scanning {PORT}... Match these numbers to your screen:")
        print(f"Look for: 1095 (Vin), 1106 (Vout), 5012 (Freq)")

        while True:
            ser.reset_input_buffer()
            ser.write(b'\x01')
            time.sleep(0.5)
            
            raw = ser.read(250)
            if raw:
                raw_map(raw)
            time.sleep(2)
            
    except Exception as e:
        print(e)

if __name__ == "__main__":
    main()