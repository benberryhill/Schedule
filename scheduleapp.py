import pandas as pd
import os
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox, scrolledtext

def load_employees_from_excel(file_path):
    """Load employees from an Excel file."""
    if not os.path.exists(file_path):
        return []  # Return an empty list if the file does not exist

    # Load the Excel file
    df = pd.read_excel(file_path)

    # Convert the DataFrame to a list of Employee objects
    employees = []
    for _, row in df.iterrows():
        name = row['Name']
        if not isinstance(name, str):
            print(f"Warning: Employee name is not a string. Converting '{name}' to string.")
            name = str(name)  # Ensure name is a string

        availability = {
            'Sun': row.get('Sun') == "Yes",  # Convert "Yes" to True
            'Mon': row.get('Mon') == "Yes",
            'Tue': row.get('Tue') == "Yes",
            'Wed': row.get('Wed') == "Yes",
            'Thu': row.get('Thu') == "Yes",
            'Fri': row.get('Fri') == "Yes",
            'Sat': row.get('Sat') == "Yes"
        }
        print(f"Loaded Employee: {name}, Availability: {availability}")  # Debug output
        employees.append(Employee(name, availability))

    return employees

class Employee:
    def __init__(self, name, availability):
        self.name = name
        self.availability = availability  # availability as a dictionary {'Sun': True, 'Mon': False, ...}

    def __str__(self):
        return f"Employee(name={self.name}, availability={self.availability})"

class Schedule:
    def __init__(self, days, employees):
        self.days = days
        self.employees = employees  # Store employees in the class
        self.employees_needed = {day: [] for day in days}  # Initialize employees_needed
        self.schedule = {day: [] for day in days}
        self.unassigned_employees = {day: [] for day in days}  # Track unassigned employees per day

    def set_employees_needed(self, employees_needed):
        """Set the number of employees needed for each day."""
        self.employees_needed = employees_needed  # Update the employees_needed attribute

    def get_max_employees_for_day(self, day):
        """Get the maximum number of employees needed for a specific day."""
        return self.employees_needed.get(day, 0)

    def add_employee_to_day(self, day, employee, force=False):
        """Assign an employee to a day, considering availability and employee limits."""
        if day not in self.schedule:
            print(f"Invalid day: {day}")
            return

        # Check availability unless force is True
        if not employee.availability.get(day, False) and not force:
            print(f"{employee.name} is not available on {day}.")
            return

        # Only check the max employee limit if not forcing a manual assignment
        if len(self.schedule[day]) < self.get_max_employees_for_day(day) or force:
            self.schedule[day].append(employee)
            self.schedule[day] = sorted(self.schedule[day], key=lambda emp: emp.name)
            print(f"Assigned {employee.name} to {day} and sorted alphabetically.")
            
            # Remove the employee from the unassigned list if they were unassigned
            if employee in self.unassigned_employees[day]:
                self.unassigned_employees[day].remove(employee)
                print(f"Removed {employee.name} from unassigned employees for {day}.")
        else:
            # This will now only affect auto-generated assignments
            if not force:
                self.unassigned_employees[day].append(employee)
                print(f"Could not assign {employee.name} to {day}, max employees reached.")

    def generate_schedule(self):
        self.unassigned_employees = {day: [] for day in self.days}  # Reset unassigned employees
        self.schedule = {day: [] for day in self.days}  # Clear the existing schedule

        for day in self.days:
            needed = self.get_max_employees_for_day(day)  # Get the number of needed employees for the day

            # Filter available employees for the current day
            available_employees = [emp for emp in self.employees if emp.availability.get(day, False)]

            # Ensure employee names are valid strings for sorting, handle potential errors
            available_employees = [emp for emp in available_employees if isinstance(emp.name, str)]

            # Reverse order for Saturday and Sunday
            if day in ['Sun', 'Sat']:
                available_employees.reverse()  # Reverse the list of employees for these days

            print(f"Available employees for {day}: {[emp.name for emp in available_employees]}")  # Debug output

            assigned_count = len(self.schedule[day])  # Start with already assigned count
            
            for employee in available_employees:
                if assigned_count < needed:
                    print(f"Trying to assign {employee.name} to {day}.")  # Debug output
                    self.add_employee_to_day(day, employee)
                    assigned_count += 1  # Increment assigned count
                else:
                    break  # Exit the loop if the required number has been met

            # Collect unassigned employees
            for employee in available_employees:
                if employee not in self.schedule[day]:
                    self.unassigned_employees[day].append(employee)

            print(f"Assigned employees for {day}: {[emp.name for emp in self.schedule[day]]}")  # Debug output

        self.refresh_unassigned_employees()
        return self.unassigned_employees

    def print_schedule(self):
        output = "Final Schedule:\n"
        for day, employees in self.schedule.items():
            employee_names = sorted([emp.name for emp in employees])
            output += f"{day}: {', '.join(employee_names) if employee_names else 'No employees assigned'}\n"
        
        # Add unassigned employees
        output += "\nUnassigned Employees:\n"
        for day, unassigned in self.unassigned_employees.items():
            # Extract names of unassigned employees and sort alphabetically
            unassigned_names = sorted([emp.name for emp in unassigned])
            output += f"{day}: {', '.join(unassigned_names) if unassigned_names else 'All employees assigned'}\n"

        return output

    def manually_add_employee(self, day, employee, force=True):
        """Manually add an employee to a day, considering availability."""
        if day not in self.schedule:
            return
        if not employee.availability.get(day, False) and not force:
            return
        self.add_employee_to_day(day, employee, force)

    def refresh_unassigned_employees(self):
        """Refresh the unassigned employees list and sort them alphabetically."""
        self.unassigned_employees = {
            day: sorted(
                [emp for emp in self.employees if emp not in self.schedule[day]], 
                key=lambda emp: str(emp.name)  # Ensure name is treated as a string
            )
            for day in self.days
        }
        print("Unassigned employees refreshed.")

class EmployeesNeededWindow:
    def __init__(self, master, schedule, app):
        self.master = master
        self.master.title("Set Employees Needed")
        self.schedule = schedule
        self.app = app  # Store the reference to the main App instance

        self.entries = {}
        for idx, day in enumerate(self.schedule.days):
            tk.Label(master, text=f"{day}:").grid(row=idx, column=0, padx=10, pady=5)
            entry = tk.Entry(master)
            entry.grid(row=idx, column=1, padx=10, pady=5)
            self.entries[day] = entry

        self.submit_button = tk.Button(master, text="Submit", command=self.submit)
        self.submit_button.grid(row=len(self.schedule.days), columnspan=2, pady=10)

    def submit(self):
        employees_needed = {}
        for day, entry in self.entries.items():
            try:
                num_needed = int(entry.get())
                employees_needed[day] = num_needed
                entry.config(state='normal', disabledforeground='gray', bg='lightgray')
            except ValueError:
                messagebox.showerror("Invalid Input", f"Please enter a valid number for {day}.")
                return

        self.schedule.set_employees_needed(employees_needed)
        messagebox.showinfo("Success", "Employees needed updated successfully!")
        
        # Generate the schedule based on the new employee needs
        self.schedule.generate_schedule()  # Generate schedule after updating needs
        
        # Refresh the schedule preview in the main app
        self.app.refresh_schedule_preview()  # Call the refresh method of the main App

        self.master.destroy()  # Close the window after submission

    def load_existing_data(self, employees_needed):
        """Populate the entries if data is already set."""
        for day, entry in self.entries.items():
            if day in employees_needed:
                entry.insert(0, employees_needed[day])
                entry.config(state='normal', disabledforeground='gray', bg='lightgray')

class ManualAssignmentWindow:
    def __init__(self, master, app, days, employees):
        self.app = app
        self.master = master
        self.days = days
        self.employees = employees  # Pass the global employees list

        tk.Label(master, text="Select Employee:").grid(row=0, column=0, padx=10, pady=5)
        self.employee_var = tk.StringVar()
        self.employee_dropdown = tk.OptionMenu(master, self.employee_var, *[emp.name for emp in employees], command=self.update_availability)
        self.employee_dropdown.grid(row=0, column=1, padx=10, pady=5)

        self.availability_label = tk.Label(master, text="Availability: ")
        self.availability_label.grid(row=1, columnspan=7, padx=10, pady=5)

        self.availability_vars = {}
        for idx, day in enumerate(days):
            self.availability_vars[day] = tk.BooleanVar()
            checkbutton = tk.Checkbutton(master, text=day, variable=self.availability_vars[day])
            checkbutton.grid(row=2, column=idx, padx=10, pady=5)

        self.submit_button = tk.Button(master, text="Assign Employee", command=self.assign_employee)
        self.submit_button.grid(row=3, columnspan=7, pady=10)

    def update_availability(self, selected_employee_name):
        """Update the availability label and checkboxes based on the selected employee."""
        employee = next((emp for emp in self.employees if emp.name == selected_employee_name), None)
        
        if employee:
            # Update the availability label to only show the available days
            available_days = [day for day, available in employee.availability.items() if available]
            availability_text = ", ".join(available_days) if available_days else "Not available"
            self.availability_label.config(text=f"Availability: {availability_text}")

            # Update the checkboxes for each day
            for day in self.days:
                self.availability_vars[day].set(employee.availability.get(day, False))

    def assign_employee(self):
        # Get all the selected days from the checkboxes
        selected_days = [day for day, var in self.availability_vars.items() if var.get()]
        employee_name = self.employee_var.get()

        if selected_days and employee_name:
            employee = next((emp for emp in self.employees if emp.name == employee_name), None)
            
            if employee:
                for day in selected_days:
                    is_available = employee.availability.get(day, False)
                    current_assigned_count = len(self.app.schedule.schedule.get(day, []))

                    # Use the method to get the max employees needed for the day
                    max_employees_needed = self.app.schedule.get_max_employees_for_day(day)

                    if is_available or current_assigned_count < max_employees_needed:
                        # Confirm assignment if employee is available
                        commit = messagebox.askyesno("Confirm Assignment", 
                                                    f"{employee.name} is available on {day}. Do you want to assign them?")
                        if commit:
                            self.app.schedule.manually_add_employee(day, employee)  # Correct reference to schedule
                            messagebox.showinfo("Success", f"{employee.name} assigned to {day}.")
                        self.app.refresh_schedule_preview()  # Refresh schedule after assignment

                    elif current_assigned_count >= max_employees_needed:
                        # If the maximum has been reached, ask if they want to assign anyway
                        force = messagebox.askyesno("Force Assignment", 
                                                    f"Maximum employees for {day} reached. Do you want to assign {employee.name} anyway?")
                        if force:
                            self.app.schedule.manually_add_employee(day, employee, True)  # True for force assignment
                            messagebox.showinfo("Success", f"{employee.name} assigned to {day} even though the max is reached.")
                        self.app.refresh_schedule_preview()  # Refresh schedule after assignment

                    else:
                        # If the employee is not available
                        force = messagebox.askyesno("Force Assignment", 
                                                    f"{employee.name} is not available on {day}. Do you want to assign them anyway?")
                        if force:
                            self.app.schedule.manually_add_employee(day, employee, True)  # Force assignment
                            messagebox.showinfo("Success", f"{employee.name} assigned to {day} even though they are unavailable.")
                        self.app.refresh_schedule_preview()  # Refresh schedule after assignment

                self.master.destroy()  # Close the window after all assignments

            else:
                messagebox.showerror("Error", "Employee not found.")
        else:
            messagebox.showerror("Error", "Please select both days and an employee.")

class ScheduleWindow:
    def __init__(self, master):
        self.master = master
        self.master.title("Schedule Management")
        self.schedule_window = self

        self.days = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat']
        self.employees = []
        self.schedule = Schedule(days=self.days, employees=[])
        
        # Top frame for buttons
        top_frame = ttk.Frame(master)
        top_frame.pack(side=tk.TOP, fill=tk.X)

        # Dropdown for selecting Excel file
        self.file_selection = ttk.Combobox(top_frame, values=self.get_excel_files(), state='readonly')
        self.file_selection.set("Select Employee Excel Sheet")
        self.file_selection.pack(side=tk.LEFT, padx=10)

        # Bind selection event
        self.file_selection.bind("<<ComboboxSelected>>", self.on_file_selected)

        # Generate Schedule Button
        self.generate_schedule_button = tk.Button(top_frame, text="Generate Schedule", command=self.generate_schedule)
        self.generate_schedule_button.pack(side=tk.LEFT, padx=5, pady=5)

        # Main frame for schedules
        main_frame = ttk.Frame(master)
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Left Preview Frame for Assigned and Unassigned
        preview_frame = tk.Frame(main_frame)
        preview_frame.grid(row=0, column=0, sticky="nsew")

        ### Schedule Treeview with Scrollbar ###
        schedule_frame = tk.Frame(preview_frame)
        schedule_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        schedule_scroll = tk.Scrollbar(schedule_frame)
        schedule_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.schedule_tree = ttk.Treeview(schedule_frame, columns=["Row"] + self.days, show="headings", height=15, yscrollcommand=schedule_scroll.set)
        self.schedule_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        schedule_scroll.config(command=self.schedule_tree.yview)

        self.schedule_tree.heading("Row", text="#")
        self.schedule_tree.column("Row", width=30, anchor="center", stretch=True)
        for day in self.days:
            self.schedule_tree.heading(day, text=day)
            self.schedule_tree.column(day, width=80, anchor="center", stretch=True)
        self.schedule_tree.tag_configure('oddrow', background='#f0f0ff')
        self.schedule_tree.tag_configure('evenrow', background='#ffffff')
        self.schedule_tree.bind("<Double-1>", self.on_schedule_double_click)

        ### Unassigned Treeview with Scrollbar ###
        unassigned_frame = tk.Frame(preview_frame)
        unassigned_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        unassigned_scroll = tk.Scrollbar(unassigned_frame)
        unassigned_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.unassigned_tree = ttk.Treeview(unassigned_frame, columns=["Row"] + self.days, show="headings", height=15, yscrollcommand=unassigned_scroll.set)
        self.unassigned_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        unassigned_scroll.config(command=self.unassigned_tree.yview)

        self.unassigned_tree.heading("Row", text="#")
        self.unassigned_tree.column("Row", width=30, anchor="center", stretch=True)
        for day in self.days:
            self.unassigned_tree.heading(day, text=day)
            self.unassigned_tree.column(day, width=80, anchor="center", stretch=True)
        self.unassigned_tree.tag_configure('oddrow', background='#f0f0ff')
        self.unassigned_tree.tag_configure('evenrow', background='#ffffff')
        self.unassigned_tree.bind("<Double-1>", self.on_unassigned_double_click)

        ### Finalized Treeview with Scrollbar ###
        finalized_frame = ttk.Frame(main_frame)
        finalized_frame.grid(row=0, column=1, sticky="nsew")

        finalized_tree_frame = tk.Frame(finalized_frame)
        finalized_tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        finalized_scroll = tk.Scrollbar(finalized_tree_frame)
        finalized_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.finalized_tree = ttk.Treeview(finalized_tree_frame, columns=["Row"] + self.days, show="headings", height=31, yscrollcommand=finalized_scroll.set)
        self.finalized_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        finalized_scroll.config(command=self.finalized_tree.yview)

        self.finalized_tree.heading("Row", text="#")
        self.finalized_tree.column("Row", width=30, anchor="center", stretch=True)
        for day in self.days:
            self.finalized_tree.heading(day, text=day)
            self.finalized_tree.column(day, width=80, anchor="center", stretch=True)
        self.finalized_tree.tag_configure('oddrow', background='#f0f0ff')
        self.finalized_tree.tag_configure('evenrow', background='#ffffff')

        # Set Employees Needed Frame
        set_employees_needed_frame = ttk.Frame(preview_frame)
        set_employees_needed_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)

        # Set Employees Needed Label
        set_employees_needed_label = tk.Label(set_employees_needed_frame, text="Set Employees Needed", font=("Arial", 12, "bold"))
        set_employees_needed_label.pack(anchor=tk.W)

        # Example Entries for Days
        self.employees_needed_entries = {}
        for day in self.days:
            label = tk.Label(set_employees_needed_frame, text=f"{day}:")
            label.pack(side=tk.LEFT, padx=5, pady=5)
            entry = tk.Entry(set_employees_needed_frame, width=5)
            entry.pack(side=tk.LEFT, padx=5, pady=5)
            self.employees_needed_entries[day] = entry  # Store entry for later use

        # Submit Button for Setting Employees Needed
        self.submit_needed_button = tk.Button(set_employees_needed_frame, text="Submit", command=self.submit_employees_needed)
        self.submit_needed_button.pack(side=tk.LEFT, padx=10, pady=5)

    def on_schedule_double_click(self, event):
        """Handle double-clicking on any assigned employee to remove them from the schedule."""
        # Get the row that was clicked
        item_id = self.schedule_tree.identify_row(event.y)
        if not item_id:
            return  # No row selected

        # Get the column that was clicked
        column = self.schedule_tree.identify_column(event.x)
        column_index = int(column.replace('#', '')) - 1  # Columns are 1-based, so subtract 1

        # Get the values for the selected row
        item_values = self.schedule_tree.item(item_id, 'values')

        if column_index == 0:
            return  # Clicked on the row number column — ignore

        # Get the employee name from the clicked column
        selected_employee = item_values[column_index]

        # If no employee found in the clicked cell, exit
        if not selected_employee:
            return

        # Determine the day based on the clicked column index
        selected_day = self.days[column_index - 1]

        # Prompt user to confirm removal
        confirm = messagebox.askyesno("Remove Employee", f"Do you want to remove {selected_employee} from {selected_day}?")

        if confirm:
            self.remove_employee_from_schedule(selected_day, selected_employee)
            messagebox.showinfo("Success", f"{selected_employee} was removed from {selected_day} and added back to unassigned employees.")
            self.refresh_schedule_preview()
            self.refresh_unassigned_employees()
        
    def on_unassigned_double_click(self, event):
        """Handle double-clicking on any unassigned employee to manually assign them."""
        item_id = self.unassigned_tree.identify_row(event.y)
        column_id = self.unassigned_tree.identify_column(event.x)  # Example: "#1" for the first column

        if not item_id or not column_id:
            return

        column_index = int(column_id.replace('#', '')) - 1  # Convert column ID to index
        if column_index == 0:
            return  # Clicked on the row number column — ignore
        
        selected_day = self.days[column_index - 1]  # Get the correct day from the column including num row

        item_values = self.unassigned_tree.item(item_id, 'values')
        if not item_values or column_index >= len(item_values):
            return

        selected_employee = item_values[column_index]
        if not selected_employee:
            return

        confirm = messagebox.askyesno("Manual Assignment", f"Do you want to manually add {selected_employee} to {selected_day}?")

        if confirm:
            self.add_employee_to_schedule([selected_day], selected_employee)
            messagebox.showinfo("Success", f"{selected_employee} has been assigned to {selected_day}.")
            self.refresh_schedule_preview()
            self.refresh_unassigned_employees()

    def get_excel_files(self):
        """Retrieve all available Excel files for selection."""
        excel_dir = os.path.join(os.getcwd(), "excel_sheets")  # Assuming excel files are in an 'excel_sheets' folder
        return [f for f in os.listdir(excel_dir) if f.endswith('.xlsx')]

    def on_file_selected(self, event):
        """Load employees from the selected Excel file."""
        selected_file = self.file_selection.get()
        if selected_file and selected_file != "Select Employee Excel Sheet":
            # Create full path to the selected file
            excel_dir = os.path.join(os.getcwd(), "excel_sheets")
            file_path = os.path.join(excel_dir, selected_file)
            
            self.employees = load_employees_from_excel(file_path)  # Pass selected file to the function
            if self.employees:  # Check if any employees were loaded
                self.schedule.employees = self.employees  # Update the schedule's employee list
                print(f"Loaded employees from {file_path}: {[emp.name for emp in self.employees]}.")  # Debug output
                self.refresh_employee_selection_menu()
            else:
                print(f"No employees loaded from {file_path}.")  # Debug if no employees are found
            
            self.refresh_unassigned_employees()  # Refresh the unassigned employees display

    def add_employee_to_schedule(self, selected_days, employee_name):
        """Add an employee to the schedule for the selected days."""
        # Look up the actual employee object based on the employee name
        employee = next((emp for emp in self.schedule.employees if emp.name == employee_name), None)

        if not employee:
            messagebox.showerror("Error", f"Employee '{employee_name}' not found.")
            return  # Exit if the employee is not found

        for day in selected_days:
            # Check if the day exists in the schedule
            if day not in self.schedule.schedule:
                messagebox.showerror("Error", f"{day} is not a valid day in the schedule.")
                continue  # Skip this day if it is invalid

            # Check if the employee is already assigned to this day
            if employee not in self.schedule.schedule[day]:
                self.schedule.schedule[day].append(employee)  # Add the employee object

                # Sort employees alphabetically by name after adding
                self.schedule.schedule[day] = sorted(self.schedule.schedule[day], key=lambda emp: emp.name.lower())

                # Remove the employee from the unassigned list for the selected days
                if employee in self.schedule.unassigned_employees[day]:
                    self.schedule.unassigned_employees[day].remove(employee)

        # Refresh the schedule preview after updating the schedule
        self.refresh_schedule_preview()
        self.refresh_unassigned_employees()

    def remove_employee_from_schedule(self, day, employee_name):
        """Remove an employee from the schedule and add them to the unassigned list."""
        # Find the employee in the schedule for the selected day
        employee = next((e for e in self.schedule.schedule[day] if e.name == employee_name), None)

        if employee:
            # Remove the employee from the schedule for that day
            self.schedule.schedule[day] = [e for e in self.schedule.schedule[day] if e.name != employee_name]

            # Add the employee back to the unassigned list for that day
            self.schedule.unassigned_employees[day].append(employee) 

    def refresh_schedule_preview(self):
        """Refresh the schedule preview treeview with current assignments."""
        self.schedule_tree.delete(*self.schedule_tree.get_children())  # Clear existing rows

        max_employees = max(len(self.schedule.schedule[day]) for day in self.days)  # Find the maximum number of employees assigned to a single day

        # Add rows for the maximum number of assigned employees
        for i in range(max_employees):
            row_values = []
            for day in self.days:
                employees = self.schedule.schedule[day]
                if i < len(employees):
                    row_values.append(employees[i].name)  # Add employee name if it exists
                else:
                    row_values.append('')  # Leave empty if no employee is assigned

            # Insert the row into the schedule treeview
            tag = 'oddrow' if i % 2 == 0 else 'evenrow'
            self.schedule_tree.insert('', 'end', values=[i + 1] + row_values, tags=(tag,)) #create each row with a num, names, and tags(colored lines)

        self.refresh_unassigned_employees()

    def refresh_unassigned_employees(self):
        """Refresh the unassigned employees treeview with current unassigned employees."""
        print("Refreshing unassigned employees...")  # Debug output
        self.unassigned_tree.delete(*self.unassigned_tree.get_children())  # Clear existing rows

        # Prepare a list of rows for each day
        row_values = []
        for day in self.days:
            # Get the list of unassigned employees for the day
            unassigned_employees = self.schedule.unassigned_employees[day]

            # Filter to show only available unassigned employees
            available_unassigned_employees = [emp for emp in unassigned_employees if emp.availability.get(day, False)]

            # Sort unassigned employees alphabetically by name
            available_unassigned_employees.sort(key=lambda emp: emp.name.lower())

            # Store employee names or leave it empty if no available unassigned employees
            if available_unassigned_employees:
                row_values.append([emp.name for emp in available_unassigned_employees])
            else:
                row_values.append([''])  # Leave empty if no available unassigned employees

        # Insert the row values into the unassigned treeview
        for i in range(max(len(values) for values in row_values) if row_values else 0):  # Check for non-empty row_values
            row = []
            for day in self.days:
                day_index = self.days.index(day)
                if i < len(row_values[day_index]):
                    row.append(row_values[day_index][i])  # Add employee name if it exists
                else:
                    row.append('')  # Leave empty if no unassigned available employee
            tag = 'oddrow' if i % 2 == 0 else 'evenrow'
            self.unassigned_tree.insert('', 'end', values=[i + 1] + row, tags=(tag,)) #create each row with a num, names, and tags(colored lines)

        print("Unassigned employees treeview refreshed.")  # Debug output

    def submit_employees_needed(self):
        """Submit the employees needed for each day."""
        success_messages = []
        error_messages = []

        for day, entry in self.employees_needed_entries.items():
            try:
                employees_needed = int(entry.get())
                self.schedule.employees_needed[day] = employees_needed
                success_messages.append(f"Set employees needed for {day} to {employees_needed}.")
            except ValueError:
                error_messages.append(f"Invalid number for {day}. Please enter a valid integer.")

        # Display a single message for all success messages
        if success_messages:
            success_message = "\n".join(success_messages)
            messagebox.showinfo("Success", success_message)

        # Display a single message for all error messages
        if error_messages:
            error_message = "\n".join(error_messages)
            messagebox.showerror("Error", error_message)

    def generate_schedule(self):
        """Generate the schedule based on current parameters."""
        self.schedule.generate_schedule()
        self.refresh_schedule_preview()

    """def create_employee_form(self):
        # Create the EmployeeCreationWindow instance and pack it at the bottom
        self.employee_creation_frame = EmployeeCreationWindow(self, self.schedule)
        self.employee_creation_frame.pack(side='bottom', fill='x', padx=10, pady=10)
"""
if __name__ == "__main__":
    root = tk.Tk()
    app = ScheduleWindow(root)
    root.mainloop()