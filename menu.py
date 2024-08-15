import tkinter as tk
from tkinter import ttk

# Global variable to keep track of the step number
step_number = 1

def add_selection():
    global step_number
    selected_mode = mode_var.get()
    selected_submode = submode_var.get()
    if selected_mode and selected_submode:
        selection = f"{selected_mode} - {selected_submode}"
    elif selected_mode:
        selection = selected_mode
    else:
        selection = "No selection"
    
    # Insert into table
    table.insert('', tk.END, values=(f"Step {step_number}", selected_mode, selected_submode))
    step_number += 1  # Increment the step number

def update_submenu(mode):
    submenu.delete(0, tk.END)  # Clear previous submenu items
    if mode == "Charge":
        submenu.add_radiobutton(label="CC", variable=submode_var, value="CC")
        submenu.add_radiobutton(label="CV", variable=submode_var, value="CV")
    elif mode == "DisCharge":
        submenu.add_radiobutton(label="CC", variable=submode_var, value="CC")
        submenu.add_radiobutton(label="CV", variable=submode_var, value="CV")
        submenu.add_radiobutton(label="CL", variable=submode_var, value="CL")
        submenu.add_radiobutton(label="CP", variable=submode_var, value="CP")
    elif mode == "Rest":
        submenu.add_radiobutton(label="Rest", variable=submode_var, value="Rest")

def on_mode_select(mode):
    mode_var.set(mode)
    update_submenu(mode)

def remove_selection():
    global step_number
    selected_item = table.selection()
    if selected_item:
        # Delete the selected item
        table.delete(selected_item)
        
        # Re-number the steps
        for index, item in enumerate(table.get_children(), start=1):
            current_values = table.item(item, 'values')
            table.item(item, values=(f"Step {index}", current_values[1], current_values[2]))
        
        # Adjust the step number for new additions
        step_number = len(table.get_children()) + 1

def edit_selection():
    selected_item = table.selection()
    if selected_item:
        # Retrieve the current values
        item = table.item(selected_item)
        current_step, current_mode, current_submode = item['values']
        
        # Create a new top-level window for editing
        edit_window = tk.Toplevel(root)
        edit_window.title("Edit Selection")

        # Mode variable for editing
        edit_mode_var = tk.StringVar(value=current_mode)
        edit_submode_var = tk.StringVar(value=current_submode)

        # Create a menu for modes
        mode_menu = tk.OptionMenu(edit_window, edit_mode_var, "Charge", "DisCharge", "Rest")
        mode_menu.pack(padx=10, pady=5)

        # Create a dropdown menu for submodes
        submode_menu = tk.OptionMenu(edit_window, edit_submode_var, "")
        submode_menu.pack(padx=10, pady=5)

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
            if new_mode and new_submode:
                table.item(selected_item, values=(current_step, new_mode, new_submode))
            edit_window.destroy()  # Close the edit window

        save_button = tk.Button(edit_window, text="Save", command=save_changes)
        save_button.pack(padx=10, pady=10)

# Create the main application window
root = tk.Tk()
root.title("Menu Example")

mode_var = tk.StringVar()
mode_var.set("Select Mode")

submode_var = tk.StringVar()
submode_var.set("Select Submode")

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

# Create a table (Treeview) to display the step, mode, and submode
columns = ('Step', 'Mode', 'Submode')
table = ttk.Treeview(root, columns=columns, show='headings')
table.heading('Step', text='Step')
table.heading('Mode', text='Mode')
table.heading('Submode', text='Submode')
table.grid(row=0, column=0, columnspan=4, padx=5, pady=5)

# Create "Remove" button
remove_button = tk.Button(root, text="Remove", command=remove_selection)
remove_button.grid(row=1, column=2, padx=5, pady=5)

# Create "Edit" button
edit_button = tk.Button(root, text="Edit", command=edit_selection)
edit_button.grid(row=1, column=3, padx=5, pady=5)

# Create "Add" button
add_button = tk.Button(root, text="Add", command=add_selection)
add_button.grid(row=1, column=1, padx=5, pady=5)

# Run the application
root.mainloop()
