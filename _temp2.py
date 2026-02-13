#delete this file

import time
import serial

PORT = "COM3"
BAUD = 9600
TIMEOUT = 1.0


def connect():
    ser = serial.Serial(
        port=PORT,
        baudrate=BAUD,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        timeout=TIMEOUT
    )

    ser.dtr = True
    ser.rts = True
    time.sleep(1.0)
    return ser


def parse_stream(data: bytes):
    i = 0
    length = len(data)

    while i < length:
        byte = data[i]

        if 0xC0 <= byte <= 0xDF:
            meter_id = byte
            val_str = ""
            j = i + 1
            digit_count = 0

            while j < length and digit_count < 8:
                b = data[j]

                if 0xF0 <= b <= 0xF9:
                    val_str += str(b & 0x0F)
                    digit_count += 1

                elif 0xE0 <= b <= 0xEF:
                    pass  # ignore dot segments

                elif 0xC0 <= b <= 0xDF:
                    break

                j += 1

            meter_hex = format(meter_id, "X")

            if meter_hex == "D3" and len(val_str) >= 5:
                yield val_str

            i = j
        else:
            i += 1


def decode_vin(raw_digits: str):
    if len(raw_digits) < 4:
        return 0.0

    # Remove first and last two digits
    core = raw_digits[1:-2]

    # Reverse digits
    rev = core[::-1]

    # Convert to voltage
    vin = float(rev) / 100.0
    return vin


def main():
    print("Connecting to VIN meter...")
    ser = connect()
    print("Connected. Decoding VIN...\n")

    try:
        while True:
            raw_data = ser.read(ser.in_waiting or 100)

            if raw_data:
                for digits in parse_stream(raw_data):
                    vin = decode_vin(digits)

                    print(
                        f"[{time.strftime('%H:%M:%S')}] "
                        f"Raw: {digits}  →  VIN: {vin:.2f} V"
                    )

            time.sleep(0.3)

    except KeyboardInterrupt:
        print("\nStopping...")
        ser.close()


if __name__ == "__main__":
    main()
