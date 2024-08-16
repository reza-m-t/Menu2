import tkinter as tk
from tkinter import ttk
import serial
import serial.tools.list_ports

# Global variable to keep track of the step number
step_number = 1

def add_selection():
    global step_number
    selected_mode = mode_var.get()
    selected_submode = submode_var.get()
    main_parameter_value = main_parameter_entry.get()

    # Check if inputs are valid
    if selected_mode and (selected_submode or selected_mode == "Rest"):
        if selected_mode != "Rest" and not main_parameter_value:
            return  # Do not add if there's no input for main parameter when required
        
        # Gather termination parameters
        termination_params = []
        if termination1_var.get() and termination1_entry.get():
            termination_params.append(f"{termination1_var.get()} = {termination1_entry.get()}")
        if termination2_var.get() and termination2_entry.get():
            termination_params.append(f"{termination2_var.get()} = {termination2_entry.get()}")
        if termination3_var.get() and termination3_entry.get():
            termination_params.append(f"{termination3_var.get()} = {termination3_entry.get()}")

        # Add a default value of 0 if no termination parameters were entered
        if not termination_params:
            termination_params.append("No Termination Parameters")

        termination_values = ", ".join(termination_params)
        main_parameter_unit = main_parameter_label.cget("text").split("(")[-1][:-1]  # Extract the unit

        if edit_mode.get():
            # Edit existing row
            selected_item = table.selection()[0]
            table.item(selected_item, values=(f"Step {table.index(selected_item) + 1}", selected_mode, selected_submode, f"{main_parameter_value} {main_parameter_unit}", termination_values))
            edit_mode.set(False)
            add_button.config(text="Add")
        else:
            # Insert into table
            table.insert('', tk.END, values=(f"Step {step_number}", selected_mode, selected_submode, f"{main_parameter_value} {main_parameter_unit}", termination_values))
            step_number += 1  # Increment the step number

        # Clear input fields after adding
        main_parameter_entry.delete(0, tk.END)
        termination1_entry.delete(0, tk.END)
        termination2_entry.delete(0, tk.END)
        termination3_entry.delete(0, tk.END)

    check_add_button_state()

def update_submenu(mode):
    submodes = []
    if mode == "Charge":
        submodes = ["CC", "CV"]
    elif mode == "DisCharge":
        submodes = ["CC", "CV", "CL", "CP"]
    elif mode == "Rest":
        submodes = ["Rest"]
    
    submode_menu['menu'].delete(0, 'end')  # Clear previous submenu items
    
    for submode in submodes:
        submode_menu['menu'].add_command(label=submode, command=tk._setit(submode_var, submode, update_main_parameter_entry))
    
    submode_var.set(submodes[0])  # Set default selection
    update_main_parameter_entry()
    check_add_button_state()

def update_main_parameter_entry(event=None):
    selected_mode = mode_var.get()
    selected_submode = submode_var.get()
    
    if selected_mode == "Rest":
        main_parameter_label.config(text="No input required")
        main_parameter_entry.config(state=tk.DISABLED)
    elif selected_mode == "Charge" and selected_submode == "CC":
        main_parameter_label.config(text="Current (mA):")
        main_parameter_entry.config(state=tk.NORMAL)
    elif selected_mode == "Charge" and selected_submode == "CV":
        main_parameter_label.config(text="Voltage (mV):")
        main_parameter_entry.config(state=tk.NORMAL)
    elif selected_mode == "DisCharge" and selected_submode == "CC":
        main_parameter_label.config(text="Current (mA):")
        main_parameter_entry.config(state=tk.NORMAL)
    elif selected_mode == "DisCharge" and selected_submode == "CV":
        main_parameter_label.config(text="Voltage (mV):")
        main_parameter_entry.config(state=tk.NORMAL)
    elif selected_mode == "DisCharge" and selected_submode == "CL":
        main_parameter_label.config(text="Resistance (ohm):")
        main_parameter_entry.config(state=tk.NORMAL)
    elif selected_mode == "DisCharge" and selected_submode == "CP":
        main_parameter_label.config(text="Power (Watt):")
        main_parameter_entry.config(state=tk.NORMAL)
    
    update_termination_parameters(selected_mode, selected_submode)
    termination_frame.grid(row=3, column=0, columnspan=4, padx=5, pady=5)
    check_add_button_state()

def update_termination_parameters(mode, submode):
    # Clear all termination options
    termination1_var.set("")
    termination2_var.set("")
    termination3_var.set("")
    termination1_entry.delete(0, tk.END)
    termination2_entry.delete(0, tk.END)
    termination3_entry.delete(0, tk.END)

    termination_options = []
    if mode == "Charge":
        if submode == "CC":
            termination_options = [("Time (min)", "min"), ("Temp (C)", "C"), ("Current (mA)", "mA")]
        elif submode == "CV":
            termination_options = [("Time (min)", "min"), ("Temp (C)", "C"), ("Voltage (mV)", "mV")]
    elif mode == "DisCharge":
        if submode == "CC":
            termination_options = [("Time (min)", "min"), ("Temp (C)", "C"), ("Current (mA)", "mA")]
        elif submode == "CV":
            termination_options = [("Time (min)", "min"), ("Temp (C)", "C"), ("Voltage (mV)", "mV")]
        elif submode == "CL" or submode == "CP":
            termination_options = [("Time (min)", "min"), ("Temp (C)", "C"), ("Current (mA)", "mA"), ("Voltage (mV)", "mV")]
    elif mode == "Rest":
        termination_options = [("Time (min)", "min"), ("Temp (C)", "C"), ("Current (mA)", "mA")]

    if termination_options:
        termination1_label.config(text=termination_options[0][0])
        termination1_unit.config(text=termination_options[0][1])
        termination1_var.set(termination_options[0][0])
        termination1_entry.config(state=tk.NORMAL)
        
        if len(termination_options) > 1:
            termination2_label.config(text=termination_options[1][0])
            termination2_unit.config(text=termination_options[1][1])
            termination2_var.set(termination_options[1][0])
            termination2_entry.config(state=tk.NORMAL)
        else:
            termination2_entry.config(state=tk.DISABLED)
        
        if len(termination_options) > 2:
            termination3_label.config(text=termination_options[2][0])
            termination3_unit.config(text=termination_options[2][1])
            termination3_var.set(termination_options[2][0])
            termination3_entry.config(state=tk.NORMAL)
        else:
            termination3_entry.config(state=tk.DISABLED)

def on_mode_select(mode):
    mode_var.set(mode)
    update_submenu(mode)
    check_add_button_state()

def check_add_button_state():
    selected_mode = mode_var.get()
    selected_submode = submode_var.get()
    main_parameter_value = main_parameter_entry.get()

    # بررسی اینکه آیا حالت انتخاب شده یا نه
    if selected_mode and (selected_submode or selected_mode == "Rest"):
        # اگر حالت Rest است، نیازی به بررسی مقدار ورودی اصلی نیست
        if selected_mode == "Rest":
            add_button.config(state=tk.NORMAL)
        else:
            # اگر حالت Rest نیست، مطمئن می‌شویم که مقدار ورودی اصلی تکمیل شده باشد
            if main_parameter_value:
                add_button.config(state=tk.NORMAL)
            else:
                add_button.config(state=tk.DISABLED)
    else:
        add_button.config(state=tk.DISABLED)


def remove_selection():
    global step_number
    selected_item = table.selection()
    if selected_item:
        # Delete the selected item
        table.delete(selected_item)
        
        # Re-number the steps
        for index, item in enumerate(table.get_children(), start=1):
            current_values = table.item(item, 'values')
            table.item(item, values=(f"Step {index}", current_values[1], current_values[2], current_values[3], current_values[4]))
        
        # Adjust the step number for new additions
        step_number = len(table.get_children()) + 1

def edit_selection():
    selected_item = table.selection()
    if selected_item:
        # Highlight the selected row
        for row in table.get_children():
            table.item(row, tags=())  # Clear tags for all rows
        table.item(selected_item, tags=('selected',))  # Tag the selected row with 'selected'

        # Retrieve current values
        current_values = table.item(selected_item, 'values')

        # Extract values
        current_mode = current_values[1]
        current_submode = current_values[2]
        main_param_with_unit = current_values[3]
        main_param_value = main_param_with_unit.split()[0] if main_param_with_unit != "No input required" else ""

        # Set the form fields with current values
        mode_var.set(current_mode)
        update_submenu(current_mode)
        submode_var.set(current_submode)
        main_parameter_entry.delete(0, tk.END)
        main_parameter_entry.insert(0, main_param_value)

        # Extract termination values (if any)
        termination_params = current_values[4].split(", ")

        termination_entries = [termination1_entry, termination2_entry, termination3_entry]
        termination_vars = [termination1_var, termination2_var, termination3_var]

        for i, param in enumerate(termination_params):
            param_parts = param.split(" = ")
            if len(param_parts) == 2 and i < 3:
                termination_vars[i].set(param_parts[0])
                termination_entries[i].delete(0, tk.END)
                termination_entries[i].insert(0, param_parts[1])

        # Set edit mode
        edit_mode.set(True)
        add_button.config(text="Save")

def clear_table():
    global step_number
    table.delete(*table.get_children())
    step_number = 1
    check_add_button_state()

def connect_serial():
    ports = serial.tools.list_ports.comports()
    serial_port = None
    for port in ports:
        if "USB" in port.description or "UART" in port.description:
            serial_port = port.device
            break
    
    if serial_port:
        try:
            ser = serial.Serial(serial_port, 9600)
            return ser
        except serial.SerialException:
            return None

def send_data_to_serial(data):
    ser = connect_serial()
    if ser:
        ser.write(data.encode('utf-8'))
        ser.close()
    else:
        print("Serial connection failed")

def send_table_data():
    table_data = []
    for row in table.get_children():
        row_values = table.item(row, 'values')
        step, mode, submode, main_param, terminations = row_values
        table_data.append(f"{step}: {mode} - {submode} - {main_param} - {terminations}")

    data_string = "\n".join(table_data)
    send_data_to_serial(data_string)
    print(f"Data sent to serial: \n{data_string}")

root = tk.Tk()
root.title("Step Selector")

# Variables
mode_var = tk.StringVar()
submode_var = tk.StringVar()
edit_mode = tk.BooleanVar(value=False)

# Modes
modes = ["Charge", "DisCharge", "Rest"]
mode_frame = ttk.LabelFrame(root, text="Mode")
mode_frame.grid(row=0, column=0, padx=5, pady=5)
for mode in modes:
    ttk.Radiobutton(mode_frame, text=mode, variable=mode_var, value=mode, command=lambda m=mode: on_mode_select(m)).pack(anchor=tk.W)

# Submode
submode_label = ttk.Label(root, text="SubMode:")
submode_label.grid(row=1, column=0, padx=5, pady=5)
submode_menu = ttk.OptionMenu(root, submode_var, "", "")
submode_menu.grid(row=1, column=1, padx=5, pady=5)

# Main parameter entry
main_parameter_label = ttk.Label(root, text="Main Parameter:")
main_parameter_label.grid(row=2, column=0, padx=5, pady=5)
main_parameter_entry = ttk.Entry(root)
main_parameter_entry.grid(row=2, column=1, padx=5, pady=5)
main_parameter_entry.bind("<KeyRelease>", check_add_button_state)

# Termination parameters
termination_frame = ttk.LabelFrame(root, text="Termination Parameters")

termination1_var = tk.StringVar()
termination1_label = ttk.Label(termination_frame, text="Termination 1:")
termination1_label.grid(row=0, column=0, padx=5, pady=5)
termination1_entry = ttk.Entry(termination_frame)
termination1_entry.grid(row=0, column=1, padx=5, pady=5)
termination1_unit = ttk.Label(termination_frame, text="")
termination1_unit.grid(row=0, column=2, padx=5, pady=5)

termination2_var = tk.StringVar()
termination2_label = ttk.Label(termination_frame, text="Termination 2:")
termination2_label.grid(row=1, column=0, padx=5, pady=5)
termination2_entry = ttk.Entry(termination_frame)
termination2_entry.grid(row=1, column=1, padx=5, pady=5)
termination2_unit = ttk.Label(termination_frame, text="")
termination2_unit.grid(row=1, column=2, padx=5, pady=5)

termination3_var = tk.StringVar()
termination3_label = ttk.Label(termination_frame, text="Termination 3:")
termination3_label.grid(row=2, column=0, padx=5, pady=5)
termination3_entry = ttk.Entry(termination_frame)
termination3_entry.grid(row=2, column=1, padx=5, pady=5)
termination3_unit = ttk.Label(termination_frame, text="")
termination3_unit.grid(row=2, column=2, padx=5, pady=5)

# Table
columns = ("Step", "Mode", "SubMode", "Main Parameter", "Termination Parameters")
table = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    table.heading(col, text=col)
table.grid(row=4, column=0, columnspan=4, padx=5, pady=5)

# Styling for selected row
table.tag_configure('selected', background='lightblue')

# Add Button
add_button = ttk.Button(root, text="Add", command=add_selection, state=tk.DISABLED)
add_button.grid(row=5, column=0, padx=5, pady=5)

# Edit Button
edit_button = ttk.Button(root, text="Edit", command=edit_selection)
edit_button.grid(row=5, column=1, padx=5, pady=5)

# Remove Button
remove_button = ttk.Button(root, text="Remove", command=remove_selection)
remove_button.grid(row=5, column=2, padx=5, pady=5)

# Reset Button
reset_button = ttk.Button(root, text="Reset", command=clear_table)
reset_button.grid(row=5, column=3, padx=5, pady=5)

# Send to Serial Button
send_button = ttk.Button(root, text="Send to Serial", command=send_table_data)
send_button.grid(row=6, column=1, columnspan=2, padx=5, pady=5)

root.mainloop()
