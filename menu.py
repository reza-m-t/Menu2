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
        
        # Insert into table
        table.insert('', tk.END, values=(f"Step {step_number}", selected_mode, selected_submode, main_parameter_value))
        step_number += 1  # Increment the step number
        main_parameter_entry.delete(0, tk.END)  # Clear the entry after adding

    # Reset the form
    mode_var.set("Select Mode")
    submode_var.set("Select Submode")
    main_parameter_label.config(text="Parameter:")
    main_parameter_entry.config(state=tk.DISABLED)
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
    
    # Re-check the add button state when the main parameter entry is updated
    check_add_button_state()

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
            table.item(item, values=(f"Step {index}", current_values[1], current_values[2], current_values[3]))
        
        # Adjust the step number for new additions
        step_number = len(table.get_children()) + 1

def edit_selection():
    selected_item = table.selection()
    if selected_item:
        # Retrieve the current values
        item = table.item(selected_item)
        current_step, current_mode, current_submode, current_main_parameter = item['values']
        
        # Create a new top-level window for editing
        edit_window = tk.Toplevel(root)
        edit_window.title("Edit Selection")

        # Mode variable for editing
        edit_mode_var = tk.StringVar(value=current_mode)
        edit_submode_var = tk.StringVar(value=current_submode)
        edit_main_parameter_var = tk.StringVar(value=current_main_parameter)

        # Create a menu for modes
        mode_menu = tk.OptionMenu(edit_window, edit_mode_var, "Charge", "DisCharge", "Rest")
        mode_menu.pack(padx=10, pady=5)

        # Create a dropdown menu for submodes
        submode_menu = tk.OptionMenu(edit_window, edit_submode_var, "")
        submode_menu.pack(padx=10, pady=5)

        # Create an entry for main parameter
        main_parameter_label_edit = tk.Label(edit_window, text="Parameter:")
        main_parameter_label_edit.pack(padx=10, pady=5)
        main_parameter_entry_edit = tk.Entry(edit_window, textvariable=edit_main_parameter_var)
        main_parameter_entry_edit.pack(padx=10, pady=5)

        # Update the submenu based on selected mode
        def update_edit_submenu(mode):
            menu = submode_menu["menu"]
            menu.delete(0, "end")
            options = []
            if mode == "Charge":
                options = ["CC", "CV"]
            elif mode == "DisCharge":
                options = ["CC", "CV", "CL", "CP"]
            elif mode == "Rest":
                options = ["Rest"]
            
            for option in options:
                menu.add_command(label=option, command=tk._setit(edit_submode_var, option))

        # Trigger submenu update when mode is changed
        edit_mode_var.trace("w", lambda *args: update_edit_submenu(edit_mode_var.get()))

        # Initialize submenu based on current mode
        update_edit_submenu(current_mode)

        # Create "Save" button to apply changes
        def save_changes():
            new_mode = edit_mode_var.get()
            new_submode = edit_submode_var.get()
            new_main_parameter = edit_main_parameter_var.get()
            if new_mode and new_submode:
                table.item(selected_item, values=(current_step, new_mode, new_submode, new_main_parameter))
            edit_window.destroy()  # Close the edit window

        save_button = tk.Button(edit_window, text="Save", command=save_changes)
        save_button.pack(padx=10, pady=10)

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

# Create a table (Treeview) to display the step, mode, submode, and main parameter
columns = ('Step', 'Mode', 'Submode', 'Main Parameter')
table = ttk.Treeview(root, columns=columns, show='headings')
table.heading('Step', text='Step')
table.heading('Mode', text='Mode')
table.heading('Submode', text='Submode')
table.heading('Main Parameter', text='Main Parameter')
table.grid(row=0, column=0, columnspan=4, padx=5, pady=5)

# Create a label and entry for the main parameter
main_parameter_label = tk.Label(root, text="Parameter:")
main_parameter_label.grid(row=1, column=0, padx=5, pady=5)

main_parameter_entry = tk.Entry(root, state=tk.DISABLED)
main_parameter_entry.grid(row=1, column=1, padx=5, pady=5)
main_parameter_entry.bind("<KeyRelease>", lambda event: check_add_button_state())  # Check button state on entry change

# Create "Remove" button
remove_button = tk.Button(root, text="Remove", command=remove_selection)
remove_button.grid(row=2, column=2, padx=5, pady=5)

# Create "Edit" button
edit_button = tk.Button(root, text="Edit", command=edit_selection)
edit_button.grid(row=2, column=3, padx=5, pady=5)

# Create "Add" button and set it to be initially disabled
add_button = tk.Button(root, text="Add", command=add_selection, state=tk.DISABLED)
add_button.grid(row=2, column=1, padx=5, pady=5)

# Run the application
root.mainloop()
