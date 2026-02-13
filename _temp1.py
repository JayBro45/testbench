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


def decode_value(raw_digits: str):
    if len(raw_digits) < 3:
        return None

    core = raw_digits[:-2]
    rev = core[::-1]

    try:
        return float(rev) / 1000.0
    except ValueError:
        return None


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
                    pass

                elif 0xC0 <= b <= 0xDF:
                    break

                j += 1

            if val_str:
                yield meter_id, val_str

            i = j
        else:
            i += 1


def main():
    print("Connecting...")
    ser = connect()
    print("Connected.\n")

    try:
        while True:
            raw_data = ser.read(ser.in_waiting or 100)

            if raw_data:
                for meter_id, digits in parse_stream(raw_data):
                    meter_hex = format(meter_id, "X")
                    value = decode_value(digits)

                    print(
                        f"[{time.strftime('%H:%M:%S')}] "
                        f"Meter {meter_hex} | Raw: {digits} | Value: {value}"
                    )

            time.sleep(0.3)

    except KeyboardInterrupt:
        print("\nStopping...")
        ser.close()


if __name__ == "__main__":
    main()
