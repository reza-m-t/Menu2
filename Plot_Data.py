import serial
import serial.tools.list_ports
import time
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from threading import Thread, Event
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import openpyxl
from collections import deque

class SerialReader(Thread):
    def __init__(self, port, baud_rate, callback):
        super().__init__()
        self.ser = None
        self.port = port
        self.baud_rate = baud_rate
        self.callback = callback
        self.running = True
        self.connection_event = Event()
    
    def run(self):
        while self.running:
            try:
                self.ser = connect_to_serial_port(self.port, self.baud_rate)
                if self.ser:
                    self.connection_event.set()
                    self.callback("Connected")
                    while self.running:
                        if self.ser.in_waiting > 0:
                            try:
                                data = self.ser.readline()
                                decoded_data = data.decode('utf-8', errors='ignore').strip()
                                self.callback(decoded_data)
                            except Exception as e:
                                self.callback(f"Error reading data: {e}")
                else:
                    self.callback("Failed to connect")
                    self.connection_event.clear()
            except serial.SerialException as e:
                self.callback(f"Serial exception: {e}")
            except Exception as e:
                self.callback(f"Unexpected error: {e}")
            
            self.callback("Reconnecting...")
            self.connection_event.clear()
            time.sleep(5)

    def stop(self):
        self.running = False
        if self.ser:
            self.ser.close()

def connect_to_serial_port(port, baud_rate=9600):
    try:
        ser = serial.Serial(port, baud_rate, timeout=1)
        print(f"Serial port {port} opened with baud rate {baud_rate}.")
        return ser
    except serial.SerialException as e:
        print(f"Error opening serial port {port}: {e}")
        return None

def find_serial_ports():
    ports = serial.tools.list_ports.comports()
    return [port.device for port in ports]

class SerialApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Serial Data Display")

        # Set global font and background color
        root.configure(bg="#f0f0f0")
        root.option_add("*Font", "Verdana 10")
        root.option_add("*Background", "#f0f0f0")
        root.option_add("*Button.Background", "#ffffff")
        root.option_add("*Button.Foreground", "#333333")

        # Create a notebook (tabs)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill=tk.BOTH, expand=True)


        # Set global font and background color
        root.configure(bg="#f0f0f0")
        root.option_add("*Font", "Verdana 10")
        root.option_add("*Background", "#f0f0f0")
        root.option_add("*Button.Background", "#ffffff")
        root.option_add("*Button.Foreground", "#333333")

        # Create file menu at the top of the window
        self.file_menu = tk.Menu(root)
        root.config(menu=self.file_menu)
        self.file_menu.add_command(label="Select Save Path", command=self.select_save_path)

        # Create frames for tabs
        self.tab1 = tk.Frame(self.notebook, bg="#f0f0f0")
        self.tab2 = tk.Frame(self.notebook, bg="#f0f0f0")

        self.notebook.add(self.tab1, text="Tab 1")
        self.notebook.add(self.tab2, text="Plot Data")

        # Plot Data tab (tab2) - Moving all settings here
        # Frame for Checkboxes
        self.checkbox_frame = tk.Frame(self.tab2, bg="#e0e0e0", bd=2, relief=tk.RIDGE)
        self.checkbox_frame.pack(pady=10, padx=10, fill=tk.X)

        self.selected_plots = {
            "Temperature": tk.BooleanVar(value=False),
            "Voltage": tk.BooleanVar(value=False),
            "Current": tk.BooleanVar(value=False),
            "Power": tk.BooleanVar(value=False)
        }

        for plot_name in self.selected_plots:
            cb = tk.Checkbutton(self.checkbox_frame, text=plot_name, variable=self.selected_plots[plot_name], 
                                bg="#e0e0e0", fg="#333333", selectcolor="#d0d0d0", padx=10, pady=5)
            cb.pack(side=tk.LEFT, anchor="w", padx=5)

        self.figure = Figure(figsize=(10, 8), dpi=100)
        self.canvas = FigureCanvasTkAgg(self.figure, master=self.tab2)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        self.message_label = tk.Label(self.tab2, text="Select Parameter To Plot", font=("Verdana", 14, "bold"), fg="#777777", bg="#f0f0f0")
        self.message_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

        self.status_label = tk.Label(self.tab2, text="Status: Disconnected", fg="red", bg="#f0f0f0", font=("Verdana", 10, "bold"))
        self.status_label.pack(pady=5, padx=10)

        # Data queues
        self.times = deque(maxlen=100)
        self.temperatures = deque(maxlen=100)
        self.voltages = deque(maxlen=100)
        self.currents = deque(maxlen=100)
        self.powers = deque(maxlen=100)

        self.start_time = time.time()

        self.serial_thread = None
        self.start_serial()

        self.excel_file_path = None

    def select_save_path(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Select Save Path"
        )
        if file_path:
            self.excel_file_path = file_path
            self.init_excel_file()
            self.update_status(f"Save path selected: {file_path}")

    def init_excel_file(self):
        if self.excel_file_path:
            self.workbook = openpyxl.Workbook()
            self.sheet = self.workbook.active
            self.sheet.title = "Serial Data"
            self.sheet.append(["Time (s)", "Temperature", "Voltage", "Current", "Power"])
            self.workbook.save(self.excel_file_path)

    def log_data_to_excel(self, time_val, temp, volt, curr, power):
        if self.excel_file_path:
            self.sheet.append([time_val, temp, volt, curr, power])
            self.workbook.save(self.excel_file_path)

    def start_serial(self):
        self.update_status("Connecting...")
        
        ports = find_serial_ports()
        if not ports:
            self.update_status("No serial ports found. Retrying in 5 seconds...")
            self.root.after(5000, self.start_serial)
            return

        port = ports[0]
        baud_rate = 9600

        if self.serial_thread and self.serial_thread.is_alive():
            self.update_status("Already connected.")
            return

        self.serial_thread = SerialReader(port, baud_rate, self.update_plot)
        self.serial_thread.start()
        self.root.after(1000, self.check_connection)
    
    def check_connection(self):
        if self.serial_thread and not self.serial_thread.connection_event.is_set():
            self.update_status("Connection lost. Retrying...")
            self.serial_thread.stop()
            self.root.after(5000, self.start_serial)
        else:
            self.root.after(1000, self.check_connection)

    def update_plot(self, data):
        if "Error reading data" in data or "Serial exception" in data:
            self.update_status(data)
        elif data == "Failed to connect":
            self.update_status("Connection failed. Retrying...")
        else:
            try:
                temp, volt, curr, power = data.split('-')
                current_time = time.time() - self.start_time
                
                self.times.append(current_time)
                self.temperatures.append(float(temp))
                self.voltages.append(float(volt))
                self.currents.append(float(curr))
                self.powers.append(float(power))

                self.log_data_to_excel(current_time, temp, volt, curr, power)

                self.update_plots()

            except ValueError:
                self.update_status("Connected")

    def update_plots(self):
        self.figure.clear()

        plot_count = sum(var.get() for var in self.selected_plots.values())
        if plot_count == 0:
            self.canvas.get_tk_widget().pack_forget()
            self.message_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            return

        self.message_label.place_forget()
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        rows, cols = (plot_count + 1) // 2, 2
        plot_index = 1

        if self.selected_plots["Temperature"].get():
            ax = self.figure.add_subplot(rows, cols, plot_index)
            ax.plot(self.times, self.temperatures, label='Temperature', color='r')
            ax.set_title("Temperature", fontsize=12)
            ax.set_xlabel("Time (s)", fontsize=10)
            ax.set_ylabel("Temperature", fontsize=10)
            ax.legend()
            plot_index += 1

        if self.selected_plots["Voltage"].get():
            ax = self.figure.add_subplot(rows, cols, plot_index)
            ax.plot(self.times, self.voltages, label='Voltage', color='b')
            ax.set_title("Voltage", fontsize=12)
            ax.set_xlabel("Time (s)", fontsize=10)
            ax.set_ylabel("Voltage", fontsize=10)
            ax.legend()
            plot_index += 1

        if self.selected_plots["Current"].get():
            ax = self.figure.add_subplot(rows, cols, plot_index)
            ax.plot(self.times, self.currents, label='Current', color='g')
            ax.set_title("Current", fontsize=12)
            ax.set_xlabel("Time (s)", fontsize=10)
            ax.set_ylabel("Current", fontsize=10)
            ax.legend()
            plot_index += 1

        if self.selected_plots["Power"].get():
            ax = self.figure.add_subplot(rows, cols, plot_index)
            ax.plot(self.times, self.powers, label='Power', color='m')
            ax.set_title("Power", fontsize=12)
            ax.set_xlabel("Time (s)", fontsize=10)
            ax.set_ylabel("Power", fontsize=10)
            ax.legend()
            plot_index += 1

        self.figure.tight_layout()
        self.canvas.draw()

    def update_status(self, message):
        self.status_label.config(text=f"Status: {message}")
        if "Disconnected" in message or "Connection lost" in message or "Connection failed" in message:
            self.status_label.config(fg="red")
        else:
            self.status_label.config(fg="green")

    def on_close(self):
        if self.serial_thread and self.serial_thread.is_alive():
            self.serial_thread.stop()
        if hasattr(self, 'workbook') and self.excel_file_path:
            self.workbook.close()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = SerialApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()
