import serial
import serial.tools.list_ports
import time

# Function to send a message and receive the response
def send_and_receive_message(port, baud_rate=9600, timeout=1):
    try:
        # Open the serial port
        with serial.Serial(port, baud_rate, timeout=timeout) as ser:
            # Send the message
            ser.write(b"Hello\n")
            # Wait for a response
            time.sleep(5)
            # Read response
            response = ser.read_all().decode('utf-8').strip()
            return response
    except serial.SerialException as e:
        print(f"Error with port {port}: {e}")
        return None

# Function to find available serial ports
def find_serial_ports():
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]

# Function to check each port
def check_ports():
    ports = find_serial_ports()
    print(f"Found ports: {ports}")
    for port in ports:
        print(f"Checking port {port}...")
        response = send_and_receive_message(port)
        if response == "Salam Reza":
            print(f"Port {port} responded with '{response}'")
        else:
            print(f"Port {port} did not respond with 'Salam Reza'.")

if __name__ == "__main__":
    check_ports()
