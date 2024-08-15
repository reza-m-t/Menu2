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
        
        # Insert into table
        table.insert('', tk.END, values=(f"Step {step_number}", selected_mode, selected_submode, main_parameter_value, termination_values))
        step_number += 1  # Increment the step number
        main_parameter_entry.delete(0, tk.END)  # Clear the entry after adding
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
    submenu.delete(0, tk.END)  # Clear previous submenu items
    if mode == "Charge":
        submenu.add_radiobutton(label="CC", variable=submode_var, value="CC", command=update_main_parameter_entry)
        submenu.add_radiobutton(label="CV", variable=submode_var, value="CV", command=update_main_parameter_entry)
    elif mode == "DisCharge":
        submenu.add_radiobutton(label="CC", variable=submode_var, value="CC", command=update_main_parameter_entry)
        submenu.add_radiobutton(label="CV", variable=submode_var, value="CV", command=update_main_parameter_entry)
        submenu.add_radiobutton(label="CL", variable=submode_var, value="CL", command=update_main_parameter_entry)
        submenu.add_radiobutton(label="CP", variable=submode_var, value="CP", command=update_main_parameter_entry)
    elif mode == "Rest":
        submenu.add_radiobutton(label="Rest", variable=submode_var, value="Rest", command=update_main_parameter_entry)
    
    # Ensure the add button is checked every time submenu is updated
    check_add_button_state()

def update_main_parameter_entry():
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
    
    # Update termination parameters based on submode
    update_termination_parameters(selected_mode, selected_submode)
    termination_frame.grid(row=3, column=0, columnspan=4, padx=5, pady=5)  # Show termination frame
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
    # Enable Add button only if both mode and submode are selected and main parameter is valid
    if mode_var.get() and (submode_var.get() or mode_var.get() == "Rest"):
        if mode_var.get() == "Rest" or main_parameter_entry.get():
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

# Create the main application window
root = tk.Tk()
root.title("Menu Example")

mode_var = tk.StringVar()
mode_var.set("Select Mode")
mode_var.trace("w", lambda *args: check_add_button_state())

submode_var = tk.StringVar()
submode_var.set("Select Submode")
submode_var.trace("w", lambda *args: check_add_button_state())

# Create main menu
main_menu = tk.Menu(root, tearoff=0)
main_menu.add_radiobutton(label="Charge", variable=mode_var, value="Charge", command=lambda: on_mode_select("Charge"))
main_menu.add_radiobutton(label="DisCharge", variable=mode_var, value="DisCharge", command=lambda: on_mode_select("DisCharge"))
main_menu.add_radiobutton(label="Rest", variable=mode_var, value="Rest", command=lambda: on_mode_select("Rest"))

# Create submenu
submenu = tk.Menu(root, tearoff=0)

# Create a menu bar and attach the main menu and submenu to it
menubar = tk.Menu(root)
menubar.add_cascade(label="Select Mode", menu=main_menu)
menubar.add_cascade(label="Submode", menu=submenu)  # Add submenu to the menu bar
root.config(menu=menubar)

# Create a table (Treeview) to display the step, mode, submode, main parameter, and termination parameters
columns = ('Step', 'Mode', 'Submode', 'Main Parameter', 'Termination Parameters')
table = ttk.Treeview(root, columns=columns, show='headings')
table.heading('Step', text='Step')
table.heading('Mode', text='Mode')
table.heading('Submode', text='Submode')
table.heading('Main Parameter', text='Main Parameter')
table.heading('Termination Parameters', text='Termination Parameters')
table.grid(row=0, column=0, columnspan=4, padx=5, pady=5)

# Create a label and entry for the main parameter
main_parameter_label = tk.Label(root, text="Parameter:")
main_parameter_label.grid(row=1, column=0, padx=5, pady=5)

main_parameter_entry = tk.Entry(root, state=tk.DISABLED)
main_parameter_entry.grid(row=1, column=1, padx=5, pady=5)
main_parameter_entry.bind("<KeyRelease>", lambda event: check_add_button_state())  # Check button state on entry change

# Termination Parameters Section
termination_frame = tk.Frame(root)
termination_frame.grid(row=3, column=0, columnspan=4, padx=5, pady=5)
termination_frame.grid_remove()  # Initially hidden

termination1_var = tk.StringVar()
termination2_var = tk.StringVar()
termination3_var = tk.StringVar()

termination1_label = tk.Label(termination_frame, text="Termination 1")
termination1_label.grid(row=0, column=0, padx=5, pady=5)
termination1_entry = tk.Entry(termination_frame)
termination1_entry.grid(row=0, column=1, padx=5, pady=5)
termination1_unit = tk.Label(termination_frame, text="")
termination1_unit.grid(row=0, column=2, padx=5, pady=5)

termination2_label = tk.Label(termination_frame, text="Termination 2")
termination2_label.grid(row=1, column=0, padx=5, pady=5)
termination2_entry = tk.Entry(termination_frame)
termination2_entry.grid(row=1, column=1, padx=5, pady=5)
termination2_unit = tk.Label(termination_frame, text="")
termination2_unit.grid(row=1, column=2, padx=5, pady=5)

termination3_label = tk.Label(termination_frame, text="Termination 3")
termination3_label.grid(row=2, column=0, padx=5, pady=5)
termination3_entry = tk.Entry(termination_frame)
termination3_entry.grid(row=2, column=1, padx=5, pady=5)
termination3_unit = tk.Label(termination_frame, text="")
termination3_unit.grid(row=2, column=2, padx=5, pady=5)

# Create "Remove" button
remove_button = tk.Button(root, text="Remove", command=remove_selection)
remove_button.grid(row=2, column=2, padx=5, pady=5)

# Create "Add" button and set it to be initially disabled
add_button = tk.Button(root, text="Add", command=add_selection, state=tk.DISABLED)
add_button.grid(row=2, column=1, padx=5, pady=5)

# Run the application
root.mainloop()
