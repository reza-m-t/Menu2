import tkinter as tk
from tkinter import ttk

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

        if edit_mode.get():
            # Edit existing row
            selected_item = table.selection()[0]
            table.item(selected_item, values=(f"Step {table.index(selected_item) + 1}", selected_mode, selected_submode, main_parameter_value, termination_values))
            edit_mode.set(False)
            add_button.config(text="Add")
        else:
            # Insert into table
            table.insert('', tk.END, values=(f"Step {step_number}", selected_mode, selected_submode, main_parameter_value, termination_values))
            step_number += 1  # Increment the step number

        # Clear input fields after adding
        main_parameter_entry.delete(0, tk.END)
        termination1_entry.delete(0, tk.END)
        termination2_entry.delete(0, tk.END)
        termination3_entry.delete(0, tk.END)

    # Reset the form
    mode_var.set("Select Mode")
    submode_var.set("Select Submode")
    main_parameter_label.config(text="Parameter:")
    main_parameter_entry.config(state=tk.DISABLED)
    termination_frame.grid_remove()  # Hide termination frame after adding
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
        # Get current values from the selected row
        current_values = table.item(selected_item, 'values')
        
        # Set the mode and submode
        mode_var.set(current_values[1])
        update_submenu(current_values[1])  # Update the submenu based on the mode
        submode_var.set(current_values[2])
        
        # Set the main parameter if not in Rest mode
        if current_values[1] != "Rest":
            main_parameter_entry.delete(0, tk.END)
            main_parameter_entry.insert(0, current_values[3])
            update_main_parameter_entry()  # Update the main parameter entry
        
        # Set the termination parameters
        termination_params = current_values[4].split(", ")
        termination1_var.set("")
        termination2_var.set("")
        termination3_var.set("")
        termination1_entry.delete(0, tk.END)
        termination2_entry.delete(0, tk.END)
        termination3_entry.delete(0, tk.END)
        
        if termination_params:
            termination1_entry.insert(0, termination_params[0].split(" = ")[1])
            if len(termination_params) > 1:
                termination2_entry.insert(0, termination_params[1].split(" = ")[1])
            if len(termination_params) > 2:
                termination3_entry.insert(0, termination_params[2].split(" = ")[1])

        # Enable edit mode
        edit_mode.set(True)
        add_button.config(text="Save Changes")

def reset_form():
    mode_var.set("Select Mode")
    submode_var.set("Select Submode")
    main_parameter_label.config(text="Parameter:")
    main_parameter_entry.config(state=tk.DISABLED)
    main_parameter_entry.delete(0, tk.END)
    termination1_var.set("")
    termination2_var.set("")
    termination3_var.set("")
    termination1_entry.delete(0, tk.END)
    termination2_entry.delete(0, tk.END)
    termination3_entry.delete(0, tk.END)
    termination_frame.grid_remove()
    edit_mode.set(False)
    add_button.config(text="Add")
    check_add_button_state()

# Create the main application window
root = tk.Tk()
root.title("Battery Test Step Manager")

# Mode selection
mode_var = tk.StringVar(value="Select Mode")
submode_var = tk.StringVar(value="Select Submode")
edit_mode = tk.BooleanVar(value=False)  # Track if we are in edit mode

modes = ["Charge", "DisCharge", "Rest"]

# Replace Menubutton with OptionMenu for better compatibility and functionality
mode_menu = tk.OptionMenu(root, mode_var, *modes, command=on_mode_select)
mode_menu.grid(row=0, column=0, padx=5, pady=5)

# Submode menu, updated dynamically based on mode selection
submode_menu = tk.OptionMenu(root, submode_var, "")
submode_menu.grid(row=0, column=1, padx=5, pady=5)

# Main parameter input
main_parameter_label = tk.Label(root, text="Parameter:")
main_parameter_label.grid(row=1, column=0, padx=5, pady=5)
main_parameter_entry = tk.Entry(root, state=tk.DISABLED)
main_parameter_entry.grid(row=1, column=1, padx=5, pady=5)

# Termination parameters
termination_frame = tk.Frame(root)
termination1_var = tk.StringVar()
termination2_var = tk.StringVar()
termination3_var = tk.StringVar()

termination1_label = tk.Label(termination_frame, text="Termination 1:")
termination1_label.grid(row=0, column=0)
termination1_entry = tk.Entry(termination_frame, state=tk.DISABLED)
termination1_entry.grid(row=0, column=1)
termination1_unit = tk.Label(termination_frame, text="Unit")
termination1_unit.grid(row=0, column=2)

termination2_label = tk.Label(termination_frame, text="Termination 2:")
termination2_label.grid(row=1, column=0)
termination2_entry = tk.Entry(termination_frame, state=tk.DISABLED)
termination2_entry.grid(row=1, column=1)
termination2_unit = tk.Label(termination_frame, text="Unit")
termination2_unit.grid(row=1, column=2)

termination3_label = tk.Label(termination_frame, text="Termination 3:")
termination3_label.grid(row=2, column=0)
termination3_entry = tk.Entry(termination_frame, state=tk.DISABLED)
termination3_entry.grid(row=2, column=1)
termination3_unit = tk.Label(termination_frame, text="Unit")
termination3_unit.grid(row=2, column=2)

# Table to display added selections
columns = ("Step", "Mode", "Submode", "Parameter", "Termination")
table = ttk.Treeview(root, columns=columns, show="headings", height=10)
for col in columns:
    table.heading(col, text=col)
table.grid(row=4, column=0, columnspan=4, padx=5, pady=5)

# Scrollbar for the table
scrollbar = ttk.Scrollbar(root, orient="vertical", command=table.yview)
table.configure(yscrollcommand=scrollbar.set)
scrollbar.grid(row=4, column=4, sticky="ns")

# Buttons
add_button = tk.Button(root, text="Add", state=tk.DISABLED, command=add_selection)
add_button.grid(row=5, column=0, padx=5, pady=5)

edit_button = tk.Button(root, text="Edit", command=edit_selection)
edit_button.grid(row=5, column=1, padx=5, pady=5)

remove_button = tk.Button(root, text="Remove", command=remove_selection)
remove_button.grid(row=5, column=2, padx=5, pady=5)

reset_button = tk.Button(root, text="Reset", command=reset_form)
reset_button.grid(row=5, column=3, padx=5, pady=5)

# Start the application
root.mainloop()
