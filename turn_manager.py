import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import json
import datetime

# Try to import tksheet for a spreadsheet-like widget; if unavailable, we’ll fall back to Treeview.
try:
    from tksheet import Sheet
except ImportError:
    print("Sheet not found")
    Sheet = None

# =============================================================================
# Database Manager
# =============================================================================
class DatabaseManager:
    def __init__(self, db_file="shift_scheduler.db"):
        self.conn = sqlite3.connect(db_file)
        self.create_tables()

    def create_tables(self):
        cursor = self.conn.cursor()
        
        # Employees table.
        cursor.execute('''CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            target_hours INTEGER NOT NULL,
            accumulated_hours INTEGER NOT NULL,
            preferences TEXT NOT NULL
        )''')
        
        # Schedules table (to save a schedule snapshot)
        cursor.execute('''CREATE TABLE IF NOT EXISTS schedules (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            schedule TEXT NOT NULL,
            UNIQUE(year, month)
        )''')

        # Shifts table.
        cursor.execute('''CREATE TABLE IF NOT EXISTS shifts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            shift_date TEXT NOT NULL,
            shift_type TEXT NOT NULL,
            employee_id INTEGER NOT NULL,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )''')
        
        # Absences table.
        cursor.execute('''CREATE TABLE IF NOT EXISTS absences (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            employee_id INTEGER NOT NULL,
            start_date TEXT NOT NULL,
            end_date TEXT NOT NULL,
            absence_type TEXT NOT NULL,
            FOREIGN KEY(employee_id) REFERENCES employees(id)
        )''')
        
        # Festivities table
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS festivities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                is_working_day INTEGER NOT NULL
            )
        ''')
        
        # Settings table.
        cursor.execute('''CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )''')
        
        # Insert default settings if they do not already exist.
        defaults = {
            'default_target_hours': '160',
            'duration_morning': '7',
            'duration_afternoon': '8',
            'duration_night': '8',
            'staffing_morning': '2',
            'staffing_afternoon': '2',
            'staffing_night': '1'
        }
        for key, value in defaults.items():
            cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?)", (key, value))
        self.conn.commit()

    # ----- Employee Operations -----
    def add_employee(self, name, target_hours, accumulated_hours, preferences):
        cursor = self.conn.cursor()
        pref_json = json.dumps(preferences)
        cursor.execute('''INSERT INTO employees (name, target_hours, accumulated_hours, preferences)
                          VALUES (?, ?, ?, ?)''', (name, target_hours, accumulated_hours, pref_json))
        self.conn.commit()

    def update_employee(self, emp_id, name, target_hours, accumulated_hours, preferences):
        cursor = self.conn.cursor()
        pref_json = json.dumps(preferences)
        cursor.execute('''UPDATE employees 
                          SET name=?, target_hours=?, accumulated_hours=?, preferences=? 
                          WHERE id=?''',
                          (name, target_hours, accumulated_hours, pref_json, emp_id))
        self.conn.commit()

    def delete_employee(self, emp_id):
        cursor = self.conn.cursor()
        cursor.execute('''DELETE FROM employees WHERE id=?''', (emp_id,))
        self.conn.commit()

    def get_employees(self):
        cursor = self.conn.cursor()
        cursor.execute('''SELECT id, name, target_hours, accumulated_hours, preferences FROM employees''')
        rows = cursor.fetchall()
        employees = []
        for row in rows:
            emp = {
                "id": row[0],
                "name": row[1],
                "target_hours": row[2],
                "accumulated_hours": row[3],
                "preferences": json.loads(row[4])
            }
            employees.append(emp)
        return employees

    # ----- Shift Operations -----
    def add_shift(self, shift_date, shift_type, employee_id):
        cursor = self.conn.cursor()
        cursor.execute('''INSERT INTO shifts (shift_date, shift_type, employee_id)
                          VALUES (?, ?, ?)''', (shift_date, shift_type, employee_id))
        self.conn.commit()

    def get_shifts_for_month(self, year, month):
        cursor = self.conn.cursor()
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year+1}-01-01"
        else:
            end_date = f"{year}-{month+1:02d}-01"
        cursor.execute('''SELECT s.id, s.shift_date, s.shift_type, s.employee_id, e.name 
                          FROM shifts s 
                          JOIN employees e ON s.employee_id = e.id
                          WHERE s.shift_date >= ? AND s.shift_date < ?
                          ORDER BY s.shift_date, s.shift_type''', (start_date, end_date))
        rows = cursor.fetchall()
        return rows

    def clear_shifts_for_month(self, year, month):
        cursor = self.conn.cursor()
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year+1}-01-01"
        else:
            end_date = f"{year}-{month+1:02d}-01"
        cursor.execute('''DELETE FROM shifts WHERE shift_date >= ? AND shift_date < ?''', (start_date, end_date))
        self.conn.commit()
    
    def update_shift_assignment(self, shift_id, employee_id):
        cursor = self.conn.cursor()
        cursor.execute("UPDATE shifts SET employee_id = ? WHERE id = ?", (employee_id, shift_id))
        self.conn.commit()


    # ----- Absence Operations -----
    def add_absence(self, employee_id, start_date, end_date, absence_type):
        cursor = self.conn.cursor()
        cursor.execute('''INSERT INTO absences (employee_id, start_date, end_date, absence_type)
                          VALUES (?, ?, ?, ?)''', (employee_id, start_date, end_date, absence_type))
        self.conn.commit()

    def get_absences_for_month(self, year, month):
        first_day_str = f"{year}-{month:02d}-01"
        if month == 12:
            next_month_str = f"{year+1}-01-01"
        else:
            next_month_str = f"{year}-{month+1:02d}-01"
        cursor = self.conn.cursor()
        query = '''SELECT employee_id, start_date, end_date, absence_type FROM absences
                   WHERE start_date < ? AND end_date >= ?'''
        cursor.execute(query, (next_month_str, first_day_str))
        rows = cursor.fetchall()
        absences = []
        for row in rows:
            absence = {
                "employee_id": row[0],
                "start_date": datetime.datetime.strptime(row[1], "%Y-%m-%d").date(),
                "end_date": datetime.datetime.strptime(row[2], "%Y-%m-%d").date(),
                "absence_type": row[3]
            }
            absences.append(absence)
        return absences
    
    
    # ----- Absence Operations -----
    def add_festivity(self, date_str, is_working_day=True):
        """
        Insert a single festivity date, marking whether it's a working day or not.
        is_working_day is a boolean; store as 1 or 0.
        """
        cursor = self.conn.cursor()
        cursor.execute('INSERT INTO festivities (date, is_working_day) VALUES (?, ?)', 
                    (date_str, 1 if is_working_day else 0))
        self.conn.commit()

    def get_festivities_for_month(self, year, month):
        """
        Return a dict where keys are 'YYYY-MM-DD' and values are True/False for is_working_day.
        """
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year+1}-01-01"
        else:
            end_date = f"{year}-{month+1:02d}-01"
        cursor = self.conn.cursor()
        cursor.execute('SELECT date, is_working_day FROM festivities WHERE date >= ? AND date < ?',
                    (start_date, end_date))
        rows = cursor.fetchall()
        # Convert to a dict: {'2025-12-25': False, ...}
        return {row[0]: bool(row[1]) for row in rows}


    # ----- Settings Operations -----
    def get_setting(self, key):
        cursor = self.conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key=?", (key,))
        row = cursor.fetchone()
        return row[0] if row else None

    def set_setting(self, key, value):
        cursor = self.conn.cursor()
        cursor.execute("UPDATE settings SET value=? WHERE key=?", (str(value), key))
        self.conn.commit()

    def get_all_settings(self):
        cursor = self.conn.cursor()
        cursor.execute("SELECT key, value FROM settings")
        rows = cursor.fetchall()
        return {key: value for key, value in rows}
    
    # ----- Schedule Operations -----
    def save_schedule(self, year, month, schedule_json):
        cursor = self.conn.cursor()
        cursor.execute("INSERT OR REPLACE INTO schedules (year, month, schedule) VALUES (?, ?, ?)",
                       (year, month, schedule_json))
        self.conn.commit()
        
    def get_schedule(self, year, month):
        cursor = self.conn.cursor()
        cursor.execute("SELECT schedule FROM schedules WHERE year=? AND month=?", (year, month))
        row = cursor.fetchone()
        return json.loads(row[0]) if row else None
    
    def delete_schedule_snapshot(self, year, month):
        cursor = self.conn.cursor()
        cursor.execute("DELETE FROM schedules WHERE year=? AND month=?", (year, month))
        self.conn.commit()


    # ----- Statistics Operations -----
    def update_employee_statistics(self):
        """
        For all shifts that occurred on or before today, group them by employee and by month.
        For every month that is complete (the last day of that month is before today),
        calculate the extra hours worked (total_hours - target_hours).
        Then update the employee's accumulated_hours to reflect the extra hours.
        (This method recalculates from scratch for all completed months.)
        """
        cursor = self.conn.cursor()
        today = datetime.date.today()

        # Get all shifts with a date on or before today.
        cursor.execute('''
            SELECT employee_id, shift_date, shift_type FROM shifts 
            WHERE shift_date <= ?
        ''', (today.strftime("%Y-%m-%d"),))
        shifts = cursor.fetchall()

        # Get shift durations from settings.
        duration_morning = int(self.get_setting("duration_morning"))
        duration_afternoon = int(self.get_setting("duration_afternoon"))
        duration_night = int(self.get_setting("duration_night"))
        shift_duration_map = {
            "Morning": duration_morning,
            "Afternoon": duration_afternoon,
            "Night": duration_night
        }

        # Group shifts by employee and by month.
        # Use a dict with keys (employee_id, year, month) and sum hours.
        worked = {}
        for employee_id, shift_date, shift_type in shifts:
            dt = datetime.datetime.strptime(shift_date, "%Y-%m-%d").date()
            key = (employee_id, dt.year, dt.month)
            worked.setdefault(key, 0)
            worked[key] += shift_duration_map.get(shift_type, 8)

        # For each employee, sum the extra hours from every complete month.
        # A month is considered complete if its last day is before today.
        extra_hours_by_emp = {}
        for (employee_id, year, month), total_hours in worked.items():
            # Compute last day of the month.
            if month == 12:
                next_month = datetime.date(year + 1, 1, 1)
            else:
                next_month = datetime.date(year, month + 1, 1)
            last_day_of_month = next_month - datetime.timedelta(days=1)

            if last_day_of_month < today:
                # Get the target hours for the employee.
                cursor.execute("SELECT target_hours FROM employees WHERE id=?", (employee_id,))
                result = cursor.fetchone()
                if result:
                    target_hours = result[0]
                    extra = total_hours - target_hours
                    if extra > 0:
                        extra_hours_by_emp.setdefault(employee_id, 0)
                        extra_hours_by_emp[employee_id] += extra

        # Now update each employee's accumulated_hours.
        # For simplicity, we set accumulated_hours equal to the computed extra hours.
        for employee_id, extra in extra_hours_by_emp.items():
            cursor.execute("UPDATE employees SET accumulated_hours=? WHERE id=?", (extra, employee_id))
        self.conn.commit()


# =============================================================================
# Employee Dialog (for Adding/Editing)
# =============================================================================
class EmployeeDialog(tk.Toplevel):
    def __init__(self, master, title="Add Employee", employee=None, default_target=None):
        super().__init__(master)
        self.title(title)
        self.geometry("300x350")
        self.employee = employee
        self.result = None
        self.grab_set()
        self.lift()
        self.focus_force()

        tk.Label(self, text="Name:").pack(pady=5)
        self.name_entry = tk.Entry(self)
        self.name_entry.pack()

        tk.Label(self, text="Target Hours (per month):").pack(pady=5)
        self.target_entry = tk.Entry(self)
        self.target_entry.pack()
        if default_target:
            self.target_entry.insert(0, str(default_target))

        tk.Label(self, text="Accumulated Hours:").pack(pady=5)
        self.accum_entry = tk.Entry(self)
        self.accum_entry.pack()
        self.accum_entry.insert(0, "0")

        self.shift_types = ["Morning", "Afternoon", "Night"]
        self.pref_vars = {}
        for shift in self.shift_types:
            tk.Label(self, text=f"{shift} Preference:").pack(pady=5)
            var = tk.StringVar()
            self.pref_vars[shift] = var
            combo = ttk.Combobox(self, textvariable=var, 
                                 values=["Avoid (0)", "Neutral (1)", "Prefer (2)"],
                                 state="readonly")
            combo.pack()
            combo.current(1)

        if self.employee:
            self.name_entry.insert(0, self.employee["name"])
            self.target_entry.delete(0, tk.END)
            self.target_entry.insert(0, str(self.employee["target_hours"]))
            self.accum_entry.delete(0, tk.END)
            self.accum_entry.insert(0, str(self.employee["accumulated_hours"]))
            for shift in self.shift_types:
                pref_value = self.employee["preferences"].get(shift, 1)
                if pref_value == 0:
                    self.pref_vars[shift].set("Avoid (0)")
                elif pref_value == 2:
                    self.pref_vars[shift].set("Prefer (2)")
                else:
                    self.pref_vars[shift].set("Neutral (1)")

        tk.Button(self, text="OK", command=self.on_ok).pack(pady=10)
        tk.Button(self, text="Cancel", command=self.destroy).pack()

    def on_ok(self):
        try:
            name = self.name_entry.get().strip()
            if not name:
                raise ValueError("Name cannot be empty.")
            target_hours_str = self.target_entry.get().strip()
            if not target_hours_str:
                raise ValueError("Target hours cannot be empty.")
            target_hours = int(target_hours_str)
            accum_str = self.accum_entry.get().strip()
            accumulated_hours = int(accum_str) if accum_str else 0
            preferences = {}
            for shift in self.shift_types:
                text = self.pref_vars[shift].get()
                if "Avoid" in text:
                    preferences[shift] = 0
                elif "Prefer" in text:
                    preferences[shift] = 2
                else:
                    preferences[shift] = 1
            self.result = {
                "name": name,
                "target_hours": target_hours,
                "accumulated_hours": accumulated_hours,
                "preferences": preferences
            }
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {e}")

# =============================================================================
# Employees Tab
# =============================================================================
class EmployeeTab(tk.Frame):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.db_manager = db_manager

        columns = ("id", "name", "target_hours", "accumulated_hours", "preferences")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col.title())
        self.tree.column("id", width=30)
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Add Employee", command=self.add_employee).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Edit Employee", command=self.edit_employee).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Delete Employee", command=self.delete_employee).pack(side=tk.LEFT, padx=5)

        self.refresh_tree()

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        employees = self.db_manager.get_employees()
        for emp in employees:
            pref_str = ", ".join([f"{k}:{v}" for k, v in emp["preferences"].items()])
            self.tree.insert("", "end", values=(emp["id"], emp["name"], emp["target_hours"],
                                                 emp["accumulated_hours"], pref_str))

    def add_employee(self):
        default_target = self.db_manager.get_setting("default_target_hours")
        dialog = EmployeeDialog(self, title="Add Employee", default_target=default_target)
        self.wait_window(dialog)
        if dialog.result:
            self.db_manager.add_employee(dialog.result["name"],
                                         dialog.result["target_hours"],
                                         dialog.result["accumulated_hours"],
                                         dialog.result["preferences"])
            self.refresh_tree()

    def edit_employee(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select Employee", "Please select an employee to edit.")
            return
        item = self.tree.item(selected[0])
        emp_id = item["values"][0]
        employees = self.db_manager.get_employees()
        employee = next((e for e in employees if e["id"] == emp_id), None)
        if employee:
            dialog = EmployeeDialog(self, title="Edit Employee", employee=employee)
            self.wait_window(dialog)
            if dialog.result:
                self.db_manager.update_employee(emp_id,
                                                dialog.result["name"],
                                                dialog.result["target_hours"],
                                                dialog.result["accumulated_hours"],
                                                dialog.result["preferences"])
                self.refresh_tree()

    def delete_employee(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select Employee", "Please select an employee to delete.")
            return
        item = self.tree.item(selected[0])
        emp_id = item["values"][0]
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this employee?"):
            self.db_manager.delete_employee(emp_id)
            self.refresh_tree()

# =============================================================================
# Schedule Tab
# =============================================================================
class ScheduleTab(tk.Frame):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.db_manager = db_manager
        self.shift_types = ["Morning", "Afternoon", "Night"]

        # --- Control Frame (Year, Month, Buttons, Filter) ---
        control_frame = tk.Frame(self)
        control_frame.pack(pady=10)
        
        # Create a navigation area to select the month
        self.current_date = datetime.date.today().replace(day=1)
        
        # Button to go to the previous month
        self.prev_button = tk.Button(control_frame, text="<", command=self.prev_month)
        self.prev_button.pack(side=tk.LEFT, padx=5)
        
        # Label to display the current month and year
        self.date_label = tk.Label(control_frame, text=self.current_date.strftime("%B %Y"), font=("Arial", 12))
        self.date_label.pack(side=tk.LEFT, padx=5)
        
        # Button to go to the next month
        self.next_button = tk.Button(control_frame, text=">", command=self.next_month)
        self.next_button.pack(side=tk.LEFT, padx=5)
        
        tk.Button(control_frame, text="Generate Schedule", command=self.generate_schedule).pack(side=tk.LEFT, padx=10)
        tk.Button(control_frame, text="Clear Schedule", command=self.clear_schedule).pack(side=tk.LEFT, padx=5)
        tk.Button(control_frame, text="Update Schedule", command=self.update_schedule).pack(side=tk.LEFT, padx=5)
        tk.Label(control_frame, text="Filter by Employee:").pack(side=tk.LEFT, padx=5)
        self.employee_filter_var = tk.StringVar()
        
        # Build a list with an "All" option plus every employee.
        employees = self.db_manager.get_employees()
        employee_options = ["All"] + [f"{emp['id']}: {emp['name']}" for emp in employees]
        self.employee_filter_combo = ttk.Combobox(control_frame, textvariable=self.employee_filter_var,
                                                  values=employee_options, state="readonly")
        self.employee_filter_combo.pack(side=tk.LEFT, padx=5)
        self.employee_filter_combo.current(0)
        tk.Button(control_frame, text="Filter Schedule", command=self.filter_schedule).pack(side=tk.LEFT, padx=5)

        # --- Create the Schedule Display Widget ---
        self.schedule_frame = tk.Frame(self)
        self.schedule_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        if Sheet:
            self.sheet = Sheet(self.schedule_frame, headers=["Date"] + self.shift_types, height=300)
            self.sheet.pack(fill=tk.BOTH, expand=True)
            self.sheet.enable_bindings()
            self.sheet.extra_begin_edit_cell_command = self.edit_cell_command
        else:
            columns = ("date",) + tuple(self.shift_types)
            self.tree = ttk.Treeview(self.schedule_frame, columns=columns, show="headings")
            for col in columns:
                self.tree.heading(col, text=col.title())
            self.tree.pack(fill=tk.BOTH, expand=True)
            self.tree.bind("<Double-1>", self.edit_treeview_cell)

        # Load the saved schedule (if any) for the current month.
        self.load_saved_schedule()

    def clear_schedule(self):
        # Use the current_date attribute from the arrow menu navigation.
        year = self.current_date.year
        month = self.current_date.month

        # Delete all shifts for the selected month.
        self.db_manager.clear_shifts_for_month(year, month)

        # Remove the saved schedule snapshot from the database.
        self.db_manager.delete_schedule_snapshot(year, month)

        # Update employee statistics to reflect the changes.
        # (Make sure your update_employee_statistics method recalculates totals based on existing shifts.)
        self.db_manager.update_employee_statistics()

        # Clear the schedule display.
        if Sheet:
            self.sheet.set_sheet_data([])
        else:
            for row in self.tree.get_children():
                self.tree.delete(row)

        messagebox.showinfo("Success", "Schedule for the selected month has been cleared and statistics updated.")

    def generate_schedule(self):
        # Use self.current_date (the date shown on the label)
        year = self.current_date.year
        month = self.current_date.month

        try:
            # Clear existing SHIFT records in DB for this month
            self.db_manager.clear_shifts_for_month(year, month)

            # Grab employees
            employees_data = self.db_manager.get_employees()
            if not employees_data:
                messagebox.showwarning("No Employees", "No employees available for scheduling.")
                return

            # Grab shift "staffing" (i.e. how many employees needed per shift)
            staffing = {
                "Morning": int(self.db_manager.get_setting("staffing_morning")),
                "Afternoon": int(self.db_manager.get_setting("staffing_afternoon")),
                "Night": int(self.db_manager.get_setting("staffing_night"))
            }

            # Grab shift durations
            shift_durations = {
                "Morning": int(self.db_manager.get_setting("duration_morning")),
                "Afternoon": int(self.db_manager.get_setting("duration_afternoon")),
                "Night": int(self.db_manager.get_setting("duration_night"))
            }

            # Load absences for the month, build a helper to check if absent
            absences_list = self.db_manager.get_absences_for_month(year, month)
            absences_by_employee = {}
            for a in absences_list:
                absences_by_employee.setdefault(a["employee_id"], []).append(a)

            def is_employee_absent(emp_id, date_obj):
                if emp_id not in absences_by_employee:
                    return False
                for absence in absences_by_employee[emp_id]:
                    if absence["start_date"] <= date_obj <= absence["end_date"]:
                        return True
                return False

            # Load festivities for the month => {date_str: True/False} => True = working day, False = non-working
            festivities = self.db_manager.get_festivities_for_month(year, month)

            # Build a lightweight employee class to track assigned_hours
            class Emp:
                def __init__(self, data):
                    self.id = data["id"]
                    self.name = data["name"]
                    self.target_hours = data["target_hours"]
                    self.accumulated_hours = data["accumulated_hours"]
                    self.preferences = data["preferences"]
                    self.assigned_hours = 0  # For newly assigned shifts this month

                def remaining_hours(self):
                    return (self.target_hours - self.accumulated_hours) - self.assigned_hours

            employees = [Emp(e) for e in employees_data]

            # Compute how many days in the selected month
            first_day = datetime.date(year, month, 1)
            if month == 12:
                next_month = datetime.date(year + 1, 1, 1)
            else:
                next_month = datetime.date(year, month + 1, 1)
            days = (next_month - first_day).days

            # Prepare a dictionary for the schedule => { date_str: {shift_type: [employee_names...] } }
            schedule = {}
            warnings_list = []

            for day in range(1, days + 1):
                current_date = datetime.date(year, month, day)
                date_str = current_date.strftime("%Y-%m-%d")

                # Check if it's a festivity *and* non-working day => skip assignment
                if date_str in festivities and festivities[date_str] == False:
                    # Non-working festivity => store empty day, no shifts assigned
                    schedule[date_str] = {}
                    continue

                # Otherwise, proceed with normal assignment
                schedule[date_str] = {}
                for shift in self.shift_types:
                    needed = staffing.get(shift, 0)
                    # Filter employees who are NOT absent
                    eligible = [e for e in employees if not is_employee_absent(e.id, current_date)]
                    if not eligible:
                        warnings_list.append(f"No eligible employees for {shift} on {date_str}")
                        schedule[date_str][shift] = []
                        continue

                    # Sort by (preference, remaining_hours). Higher preference => 2, normal => 1, avoid => 0
                    # We also want employees with more 'remaining_hours' up front
                    # so (pref, remaining_hours) in descending order
                    def sort_key(e):
                        pref = e.preferences.get(shift, 1)  # default is 1 if not found
                        return (pref, e.remaining_hours())

                    sorted_candidates = sorted(eligible, key=sort_key, reverse=True)

                    # If fewer unique candidates than needed, we fill with top candidate again => under-staffed warning
                    if len(sorted_candidates) < needed:
                        warnings_list.append(f"Not enough employees for {shift} on {date_str}; re-using top candidate")
                        # We can do something like:
                        extra_needed = needed - len(sorted_candidates)
                        assigned = sorted_candidates + [sorted_candidates[0]] * extra_needed
                    else:
                        assigned = sorted_candidates[:needed]

                    # Now record their assignment in DB and local schedule
                    assigned_names = []
                    for e in assigned:
                        e.assigned_hours += shift_durations.get(shift, 8)
                        # Insert shift record in DB
                        self.db_manager.add_shift(date_str, shift, e.id)
                        assigned_names.append(e.name)

                    # Put the assigned names in the schedule dictionary
                    schedule[date_str][shift] = assigned_names

            # Build the data for displaying in tksheet or treeview
            sheet_data = []
            for date_str in sorted(schedule.keys()):
                row = [date_str]
                for shift in self.shift_types:
                    emp_list = schedule[date_str].get(shift, [])
                    row.append(", ".join(emp_list))
                sheet_data.append(row)

            # Update the UI
            if Sheet:
                self.sheet.set_sheet_data(sheet_data)
                # Optionally highlight working vs. non-working festivity rows
                self.highlight_festivity_rows(sheet_data, festivities)
            else:
                # Treeview fallback
                for row in self.tree.get_children():
                    self.tree.delete(row)
                for row in sheet_data:
                    self.insert_festivity_treeview_row(row, festivities)

            # Show any warnings or success
            if warnings_list:
                msg = "\n".join(warnings_list)
                messagebox.showwarning("Warning", f"Some issues occurred:\n{msg}")
            else:
                messagebox.showinfo("Success", "Schedule generated successfully.")

            # Save the schedule to the 'schedules' table
            self.db_manager.save_schedule(year, month, json.dumps(schedule))

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate schedule: {e}")

    def update_schedule(self):
        year = self.year_var.get()
        month = self.month_var.get()
        try:
            # Load current schedule records for the month.
            current_shifts = self.db_manager.get_shifts_for_month(year, month)
            # Organize schedule: {date_str: {shift_type: [(shift_id, emp_id, emp_name), ...]}}
            schedule = {}
            for record in current_shifts:
                shift_id, shift_date, shift_type, emp_id, emp_name = record
                schedule.setdefault(shift_date, {}).setdefault(shift_type, []).append((shift_id, emp_id, emp_name))
            
            # Load employee records and build a map.
            employees_data = self.db_manager.get_employees()
            emp_map = {}
            for e in employees_data:
                emp_map[e["id"]] = {
                    "name": e["name"],
                    "target_hours": e["target_hours"],
                    "accumulated_hours": e["accumulated_hours"],
                    "preferences": e["preferences"],
                    "assigned_hours": 0  # Will compute next.
                }
            
            # Load shift durations from settings.
            shift_durations = {
                "Morning": int(self.db_manager.get_setting("duration_morning")),
                "Afternoon": int(self.db_manager.get_setting("duration_afternoon")),
                "Night": int(self.db_manager.get_setting("duration_night"))
            }
            
            # Compute current assigned hours per employee.
            for date_str, shifts in schedule.items():
                for shift, assignments in shifts.items():
                    for (shift_id, emp_id, emp_name) in assignments:
                        if emp_id in emp_map:
                            emp_map[emp_id]["assigned_hours"] += shift_durations.get(shift, 8)
            
            # Load absences for the month.
            absences_list = self.db_manager.get_absences_for_month(year, month)
            absences_by_emp = {}
            for a in absences_list:
                absences_by_emp.setdefault(a["employee_id"], []).append((a["start_date"], a["end_date"]))
            def is_absent(emp_id, date_obj):
                if emp_id not in absences_by_emp:
                    return False
                for (start, end) in absences_by_emp[emp_id]:
                    if start <= date_obj <= end:
                        return True
                return False

            changes = 0
            # Determine the days in the month.
            first_day = datetime.date(year, month, 1)
            if month == 12:
                next_month = datetime.date(year + 1, 1, 1)
            else:
                next_month = datetime.date(year, month + 1, 1)
            days = (next_month - first_day).days

            # --- Pass 1: Fix assignments where the employee is absent ---
            for day in range(1, days + 1):
                current_date = datetime.date(year, month, day)
                date_str = current_date.strftime("%Y-%m-%d")
                if date_str not in schedule:
                    continue
                for shift in self.shift_types:
                    if shift not in schedule[date_str]:
                        continue
                    new_assignments = []
                    # Get the list of employee IDs already assigned for this shift on this day.
                    assigned_ids = [rec[1] for rec in schedule[date_str][shift]]
                    for (shift_id, emp_id, emp_name) in schedule[date_str][shift]:
                        if is_absent(emp_id, current_date):
                            # Look for a replacement candidate.
                            candidate = None
                            best_score = -1e9
                            for cand_id, cand in emp_map.items():
                                if is_absent(cand_id, current_date):
                                    continue
                                # Skip if already assigned on that shift.
                                if cand_id in assigned_ids:
                                    continue
                                # Compute candidate's "remaining" work need.
                                remaining = (cand["target_hours"] - cand["accumulated_hours"]) - cand["assigned_hours"]
                                # Weight by preference (assume preference: 0=avoid, 1=neutral, 2=prefer).
                                preference = cand["preferences"].get(shift, 1)
                                score = remaining + (preference * 10)
                                if score > best_score:
                                    best_score = score
                                    candidate = cand_id
                            if candidate is not None:
                                # Update this shift in the database.
                                cursor = self.db_manager.conn.cursor()
                                cursor.execute("UPDATE shifts SET employee_id = ? WHERE id = ?", (candidate, shift_id))
                                self.db_manager.conn.commit()
                                # Update assigned hours.
                                emp_map[emp_id]["assigned_hours"] -= shift_durations.get(shift, 8)
                                emp_map[candidate]["assigned_hours"] += shift_durations.get(shift, 8)
                                # Replace the assignment.
                                new_assignments.append((shift_id, candidate, emp_map[candidate]["name"]))
                                changes += 1
                            else:
                                new_assignments.append((shift_id, emp_id, emp_name))
                        else:
                            new_assignments.append((shift_id, emp_id, emp_name))
                    schedule[date_str][shift] = new_assignments
            # --- Pass 2: Rebalance shifts (if an employee is over-assigned) ---
            # (Threshold here is set to –5 hours; adjust as needed.)
            for day in range(1, days + 1):
                current_date = datetime.date(year, month, day)
                date_str = current_date.strftime("%Y-%m-%d")
                if date_str not in schedule:
                    continue
                for shift in self.shift_types:
                    if shift not in schedule[date_str]:
                        continue
                    new_assignments = []
                    assigned_ids = [rec[1] for rec in schedule[date_str][shift]]
                    for (shift_id, emp_id, emp_name) in schedule[date_str][shift]:
                        # Compute remaining hours for this employee.
                        remaining = (emp_map[emp_id]["target_hours"] - emp_map[emp_id]["accumulated_hours"]) - emp_map[emp_id]["assigned_hours"]
                        if remaining < -5:  # if over-assigned by more than 5 hours
                            # Look for an alternative candidate not assigned on this shift.
                            candidate = None
                            best_remaining = -1e9
                            for cand_id, cand in emp_map.items():
                                if is_absent(cand_id, current_date):
                                    continue
                                if cand_id in assigned_ids:
                                    continue
                                cand_remaining = (cand["target_hours"] - cand["accumulated_hours"]) - cand["assigned_hours"]
                                if cand_remaining > best_remaining and cand_remaining > 0:
                                    best_remaining = cand_remaining
                                    candidate = cand_id
                            if candidate is not None:
                                cursor = self.db_manager.conn.cursor()
                                cursor.execute("UPDATE shifts SET employee_id = ? WHERE id = ?", (candidate, shift_id))
                                self.db_manager.conn.commit()
                                emp_map[emp_id]["assigned_hours"] -= shift_durations.get(shift, 8)
                                emp_map[candidate]["assigned_hours"] += shift_durations.get(shift, 8)
                                new_assignments.append((shift_id, candidate, emp_map[candidate]["name"]))
                                changes += 1
                            else:
                                new_assignments.append((shift_id, emp_id, emp_name))
                        else:
                            new_assignments.append((shift_id, emp_id, emp_name))
                    schedule[date_str][shift] = new_assignments

            # Rebuild the display data.
            sheet_data = []
            for date_str, shifts in sorted(schedule.items()):
                row = [date_str]
                for shift in self.shift_types:
                    names = [rec[2] for rec in shifts.get(shift, [])]
                    row.append(", ".join(names))
                sheet_data.append(row)
            if Sheet:
                self.sheet.set_sheet_data(sheet_data)
            else:
                for row in self.tree.get_children():
                    self.tree.delete(row)
                for row in sheet_data:
                    self.tree.insert("", "end", values=row)
            messagebox.showinfo("Update Complete", f"Schedule updated. {changes} shift(s) changed.")
        
            # Save the schedule (as JSON) into the database
            self.db_manager.save_schedule(year, month, json.dumps(schedule))
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update schedule: {e}")

    def edit_cell_command(self, event):
        # event is expected to be a tuple like (widget, row, col, ...)
        row, col = event[1], event[2]
        if col == 0:  # Skip editing the Date column
            return
        shift_type = self.shift_types[col - 1]
        date_str = self.sheet.get_cell_data(row, 0)
        # Get the corresponding shift record from the database
        cursor = self.db_manager.conn.cursor()
        cursor.execute("SELECT id, employee_id FROM shifts WHERE shift_date=? AND shift_type=?", (date_str, shift_type))
        record = cursor.fetchone()
        if not record:
            return
        shift_id, current_emp_id = record
        # Open the drop-down selection dialog:
        dialog = EmployeeSelectionDialog(self, self.db_manager, current_emp_id)
        self.wait_window(dialog)
        if dialog.result:
            new_emp_id = dialog.result
            self.db_manager.update_shift_assignment(shift_id, new_emp_id)
            cursor.execute("SELECT name FROM employees WHERE id=?", (new_emp_id,))
            new_emp_name = cursor.fetchone()[0]
            self.sheet.set_cell_data(row, col, new_emp_name)

    def filter_schedule(self):
        selected = self.employee_filter_var.get()
        year = self.year_var.get()
        month = self.month_var.get()
        if selected == "All":
            # Re–generate full schedule if "All" is selected
            self.generate_schedule()
            return
        emp_id = int(selected.split(":")[0])
        # Get all shifts for the month
        shifts = self.db_manager.get_shifts_for_month(year, month)
        # Build a filtered schedule dictionary with only the selected employee's shifts
        schedule = {}
        for shift in shifts:
            shift_id, shift_date, shift_type, employee_id, emp_name = shift
            if employee_id == emp_id:
                if shift_date not in schedule:
                    schedule[shift_date] = {}
                schedule[shift_date][shift_type] = emp_name
        # Build the sheet (or treeview) data accordingly:
        sheet_data = []
        first_day = datetime.date(year, month, 1)
        if month == 12:
            next_month = datetime.date(year+1, 1, 1)
        else:
            next_month = datetime.date(year, month+1, 1)
        days = (next_month - first_day).days
        for day in range(1, days + 1):
            date_str = datetime.date(year, month, day).strftime("%Y-%m-%d")
            row = [date_str]
            for shift in self.shift_types:
                value = schedule.get(date_str, {}).get(shift, "")
                row.append(value)
            sheet_data.append(row)
        if Sheet:
            self.sheet.set_sheet_data(sheet_data)
        else:
            for row in self.tree.get_children():
                self.tree.delete(row)
            for row in sheet_data:
                self.tree.insert("", "end", values=row)

    def edit_treeview_cell(self, event):
        # Figure out which row and column are clicked.
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        selected_item = self.tree.selection()
        if not selected_item:
            return
        # The row in the Treeview
        item = self.tree.item(selected_item)
        row_values = item["values"]
        # Suppose row_values = [date_str, morning_assignment, afternoon_assignment, night_assignment]
        # The column index (0=Date, 1=Morning, 2=Afternoon, 3=Night)
        column = self.tree.identify_column(event.x)  # e.g. "#1", "#2", ...
        col_index = int(column.replace("#", "")) - 1
        if col_index == 0:
            # The Date column - skip
            return

        # Determine which shift was clicked
        shift_type = self.shift_types[col_index - 1]  # since we have "date" in col 0
        date_str = row_values[0]

        # Query the shift record
        cursor = self.db_manager.conn.cursor()
        cursor.execute("SELECT id, employee_id FROM shifts WHERE shift_date=? AND shift_type=?",
                    (date_str, shift_type))
        record = cursor.fetchone()
        if not record:
            return

        shift_id, current_emp_id = record

        # Launch the same EmployeeSelectionDialog you use for the tksheet approach
        dialog = EmployeeSelectionDialog(self, self.db_manager, current_emp_id)
        self.wait_window(dialog)
        if dialog.result:
            new_emp_id = dialog.result
            self.db_manager.update_shift_assignment(shift_id, new_emp_id)

            # Update the displayed Treeview text
            cursor.execute("SELECT name FROM employees WHERE id=?", (new_emp_id,))
            result = cursor.fetchone()
            new_emp_name = result[0] if result else ""

            # Build a new row_values list with the updated name
            new_values = list(row_values)
            new_values[col_index] = new_emp_name
            self.tree.item(selected_item, values=new_values)

    def load_saved_schedule(self):
        year = self.current_date.year
        month = self.current_date.month
        schedule = self.db_manager.get_schedule(year, month)
        if schedule:
            sheet_data = []
            # Convert the saved schedule dictionary into sheet data.
            for date_str, shifts in sorted(schedule.items()):
                row = [date_str]
                for shift in self.shift_types:
                    value = shifts.get(shift, "")
                    if isinstance(value, list):
                        row.append(", ".join(value))
                    else:
                        row.append(value)
                sheet_data.append(row)
            if Sheet:
                self.sheet.set_sheet_data(sheet_data)
            else:
                for row in self.tree.get_children():
                    self.tree.delete(row)
                for row in sheet_data:
                    self.tree.insert("", "end", values=row)
        else:
            # If no schedule is saved, clear the display.
            if Sheet:
                self.sheet.set_sheet_data([])
            else:
                for row in self.tree.get_children():
                    self.tree.delete(row)
            # Optionally, you can show a message that no schedule exists.
            # messagebox.showinfo("Info", f"No saved schedule for {self.current_date.strftime('%B %Y')}.")

    def prev_month(self):
        # Subtract one month from self.current_date
        year = self.current_date.year
        month = self.current_date.month
        if month == 1:
            new_year = year - 1
            new_month = 12
        else:
            new_year = year
            new_month = month - 1

        self.current_date = datetime.date(new_year, new_month, 1)
        self.date_label.config(text=self.current_date.strftime("%B %Y"))
        
        # Now load any saved schedule (if it exists)
        self.load_saved_schedule()

    def next_month(self):
        # Add one month to self.current_date
        year = self.current_date.year
        month = self.current_date.month
        if month == 12:
            new_year = year + 1
            new_month = 1
        else:
            new_year = year
            new_month = month + 1

        self.current_date = datetime.date(new_year, new_month, 1)
        self.date_label.config(text=self.current_date.strftime("%B %Y"))

        # Now load any saved schedule (if it exists)
        self.load_saved_schedule()

    def highlight_festivity_rows(self, sheet_data, festivities):
        """
        Color the rows in tksheet:
        - red background if the day is a non-working festivity
        - yellow background if the day is a working festivity
        """
        for row_index, row_values in enumerate(sheet_data):
            date_str = row_values[0]
            if date_str in festivities:
                # If festivities[date_str] == False => non-working => red
                # If festivities[date_str] == True  => working => yellow
                if festivities[date_str]:
                    bg_color = "yellow"
                else:
                    bg_color = "red"
                # tksheet has highlight_cells() for row or for individual cells
                self.sheet.highlight_cells(
                    row=row_index, 
                    column=0,       # date column only, or skip 'column' to highlight the entire row
                    bg=bg_color
                )
                # If you want to highlight the entire row: 
                # self.sheet.highlight_cells(row=row_index, bg=bg_color)

    def insert_festivity_treeview_row(self, row_values, festivities):
        """
        row_values is something like [date_str, assigned_morning, assigned_afternoon, assigned_night]
        festivities is dict => date_str -> bool
        """
        date_str = row_values[0]
        if date_str in festivities:
            is_working = festivities[date_str]
            if is_working:
                tags = ("working_fest",)
            else:
                tags = ("nonworking_fest",)
        else:
            tags = ()
        
        self.tree.insert("", "end", values=row_values, tags=tags)


# =============================================================================
# Absence Dialog (for adding an absence record)
# =============================================================================
class AbsenceDialog(tk.Toplevel):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.title("Add Absence")
        self.geometry("300x300")
        self.db_manager = db_manager
        self.result = None
        self.grab_set()
        self.lift()
        self.focus_force()

        tk.Label(self, text="Employee:").pack(pady=5)
        self.employee_var = tk.StringVar()
        employees = self.db_manager.get_employees()
        self.employee_options = {f"{emp['id']}: {emp['name']}": emp["id"] for emp in employees}
        self.employee_combo = ttk.Combobox(self, textvariable=self.employee_var,
                                           values=list(self.employee_options.keys()), state="readonly")
        self.employee_combo.pack(pady=5)
        if self.employee_options:
            self.employee_combo.current(0)

        tk.Label(self, text="Start Date (YYYY-MM-DD):").pack(pady=5)
        self.start_entry = tk.Entry(self)
        self.start_entry.pack(pady=5)

        tk.Label(self, text="End Date (YYYY-MM-DD):").pack(pady=5)
        self.end_entry = tk.Entry(self)
        self.end_entry.pack(pady=5)

        tk.Label(self, text="Absence Type (Sickness, Holiday, Maternity):").pack(pady=5)
        self.type_entry = tk.Entry(self)
        self.type_entry.pack(pady=5)

        tk.Button(self, text="OK", command=self.on_ok).pack(pady=10)
        tk.Button(self, text="Cancel", command=self.destroy).pack(pady=5)

    def on_ok(self):
        try:
            emp_key = self.employee_var.get()
            if not emp_key:
                raise ValueError("Please select an employee.")
            employee_id = self.employee_options[emp_key]
            start_date = self.start_entry.get().strip()
            end_date = self.end_entry.get().strip()
            absence_type = self.type_entry.get().strip()
            if not start_date or not end_date or not absence_type:
                raise ValueError("All fields are required.")
            self.result = {
                "employee_id": employee_id,
                "start_date": start_date,
                "end_date": end_date,
                "absence_type": absence_type
            }
            self.db_manager.add_absence(employee_id, start_date, end_date, absence_type)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {e}")


# =============================================================================
# Absences Tab
# =============================================================================
class AbsencesTab(tk.Frame):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.db_manager = db_manager

        columns = ("id", "employee", "start_date", "end_date", "absence_type")
        self.tree = ttk.Treeview(self, columns=columns, show="headings", selectmode="browse")
        for col in columns:
            self.tree.heading(col, text=col.title())
        self.tree.column("id", width=30)
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Add Absence", command=self.add_absence).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Delete Absence", command=self.delete_absence).pack(side=tk.LEFT, padx=5)

        self.refresh_tree()

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        cursor = self.db_manager.conn.cursor()
        cursor.execute('''SELECT a.id, e.name, a.start_date, a.end_date, a.absence_type
                          FROM absences a
                          JOIN employees e ON a.employee_id = e.id''')
        for row in cursor.fetchall():
            self.tree.insert("", "end", values=row)

    def add_absence(self):
        dialog = AbsenceDialog(self, self.db_manager)
        self.wait_window(dialog)
        self.refresh_tree()

    def delete_absence(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select Absence", "Please select an absence to delete.")
            return
        item = self.tree.item(selected[0])
        absence_id = item["values"][0]
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this absence?"):
            cursor = self.db_manager.conn.cursor()
            cursor.execute("DELETE FROM absences WHERE id=?", (absence_id,))
            self.db_manager.conn.commit()
            self.refresh_tree()


# =============================================================================
# Festivity Dialog
# =============================================================================
class FestivityDialog(tk.Toplevel):
    def __init__(self, master, db_manager, festivity=None):
        """
        If 'festivity' is provided, it's a dict or row to edit.
        Otherwise, we'll treat this as 'add new festivity.'
        """
        super().__init__(master)
        self.db_manager = db_manager
        self.title("Add/Edit Festivity")
        self.geometry("300x200")
        self.grab_set()     # Make this dialog modal
        self.lift()
        self.focus_force()

        self.result = None  # We'll store the result here when OK is pressed

        # Festivity Date
        tk.Label(self, text="Date (YYYY-MM-DD):").pack(pady=5)
        self.date_entry = tk.Entry(self)
        self.date_entry.pack(pady=5)

        # Is Working Day?
        tk.Label(self, text="Is Working Day? (Yes=1, No=0):").pack(pady=5)
        self.is_working_var = tk.BooleanVar()
        self.is_working_var.set(True)  # default
        # If you prefer a Checkbutton:
        self.working_chk = ttk.Checkbutton(
            self, text="Check if working day",
            variable=self.is_working_var
        )
        self.working_chk.pack(pady=5)

        # If editing an existing record, populate fields
        if festivity:
            self.date_entry.insert(0, festivity["date"])
            self.is_working_var.set(festivity["is_working_day"])

        # Button row
        tk.Button(self, text="OK", command=self.on_ok).pack(side=tk.LEFT, padx=10, pady=10)
        tk.Button(self, text="Cancel", command=self.destroy).pack(side=tk.LEFT, padx=10, pady=10)

    def on_ok(self):
        try:
            date_str = self.date_entry.get().strip()
            if not date_str:
                raise ValueError("Date cannot be empty.")
            is_working = self.is_working_var.get()  # Boolean
            self.result = {
                "date": date_str,
                "is_working_day": is_working
            }
            self.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Invalid input: {e}")


# =============================================================================
# Festivity Tab
# =============================================================================
class FestivitiesTab(tk.Frame):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.db_manager = db_manager

        columns = ("id", "date", "is_working_day")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col.title())
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="Add Festivity", command=self.add_festivity).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Delete Festivity", command=self.delete_festivity).pack(side=tk.LEFT, padx=5)
        # If you want an Edit button
        tk.Button(btn_frame, text="Edit Festivity", command=self.edit_festivity).pack(side=tk.LEFT, padx=5)

        self.refresh_tree()

    def refresh_tree(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        cursor = self.db_manager.conn.cursor()
        cursor.execute("SELECT id, date, is_working_day FROM festivities ORDER BY date")
        for fest_id, date_str, is_working_day in cursor.fetchall():
            self.tree.insert("", "end", values=(fest_id, date_str, is_working_day))

    def add_festivity(self):
        dialog = FestivityDialog(self, self.db_manager)
        self.wait_window(dialog)
        if dialog.result:
            date_str = dialog.result["date"]
            is_working_day_bool = dialog.result["is_working_day"]
            self.db_manager.add_festivity(date_str, is_working_day_bool)
            self.refresh_tree()

    def edit_festivity(self):
        # First, find the currently selected row
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("No selection", "Please select a festivity to edit.")
            return
        item = self.tree.item(selected[0])
        fest_id, date_str, is_working_day_int = item["values"]

        # Build the 'festivity' dict
        fest_data = {
            "id": fest_id,
            "date": date_str,
            "is_working_day": bool(is_working_day_int)
        }

        dialog = FestivityDialog(self, self.db_manager, festivity=fest_data)
        self.wait_window(dialog)
        if dialog.result:
            new_date = dialog.result["date"]
            new_working_bool = dialog.result["is_working_day"]
            # Just do an UPDATE in the DB
            cursor = self.db_manager.conn.cursor()
            cursor.execute("""
                UPDATE festivities
                SET date=?, is_working_day=?
                WHERE id=?
            """, (new_date, 1 if new_working_bool else 0, fest_id))
            self.db_manager.conn.commit()
            self.refresh_tree()

    def delete_festivity(self):
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Select Festivity", "Please select a festivity to delete.")
            return
        item = self.tree.item(selected[0])
        fest_id = item["values"][0]
        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this festivity?"):
            cursor = self.db_manager.conn.cursor()
            cursor.execute("DELETE FROM festivities WHERE id=?", (fest_id,))
            self.db_manager.conn.commit()
            self.refresh_tree()


# =============================================================================
# Settings Tab
# =============================================================================
class SettingsTab(tk.Frame):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.db_manager = db_manager

        self.settings_keys = [
            ("default_target_hours", "Default Target Hours (per month):"),
            ("duration_morning", "Morning Shift Duration (hours):"),
            ("duration_afternoon", "Afternoon Shift Duration (hours):"),
            ("duration_night", "Night Shift Duration (hours):"),
            ("staffing_morning", "Staffing for Morning Shift:"),
            ("staffing_afternoon", "Staffing for Afternoon Shift:"),
            ("staffing_night", "Staffing for Night Shift:")
        ]
        self.entries = {}

        for key, label_text in self.settings_keys:
            frame = tk.Frame(self)
            frame.pack(pady=5, anchor="w", padx=10)
            tk.Label(frame, text=label_text, width=30, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar()
            var.set(self.db_manager.get_setting(key))
            entry = tk.Entry(frame, textvariable=var, width=10)
            entry.pack(side=tk.LEFT)
            self.entries[key] = var

        tk.Button(self, text="Save Settings", command=self.save_settings).pack(pady=20)

    def save_settings(self):
        try:
            for key, var in self.entries.items():
                value = var.get().strip()
                if not value:
                    raise ValueError(f"Value for {key} cannot be empty.")
                self.db_manager.set_setting(key, value)
            messagebox.showinfo("Settings Updated", "Settings have been updated successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update settings: {e}")


# =============================================================================
# Main Application
# =============================================================================
class ShiftSchedulerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Hospital Shift Scheduler")
        self.geometry("900x700")
        self.db_manager = DatabaseManager()

        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True)

        self.emp_tab = EmployeeTab(notebook, self.db_manager)
        notebook.add(self.emp_tab, text="Employees")

        self.schedule_tab = ScheduleTab(notebook, self.db_manager)
        notebook.add(self.schedule_tab, text="Schedule")

        self.stats_tab = StatsTab(notebook, self.db_manager)
        notebook.add(self.stats_tab, text="Statistics")

        self.settings_tab = SettingsTab(notebook, self.db_manager)
        notebook.add(self.settings_tab, text="Settings")

        self.absences_tab = AbsencesTab(notebook, self.db_manager)
        notebook.add(self.absences_tab, text="Absences")
        
        self.fest_tab = FestivitiesTab(notebook, self.db_manager)
        notebook.add(self.fest_tab, text="Festivities")


        self.lift()
        self.focus_force()
        self.update_idletasks()


# =============================================================================
# Statistics Tab
# =============================================================================
class StatsTab(tk.Frame):
    def __init__(self, master, db_manager):
        super().__init__(master)
        self.db_manager = db_manager
        self.shift_duration = 8  # Or use a setting if desired
        
        # Initialize current_date to the first day of the current month.
        self.current_date = datetime.date.today().replace(day=1)
        
        # Build the control frame with arrow buttons and a label.
        control_frame = tk.Frame(self)
        control_frame.pack(pady=10)
        
        self.prev_button = tk.Button(control_frame, text="<", command=self.prev_month)
        self.prev_button.pack(side=tk.LEFT, padx=5)
        
        self.date_label = tk.Label(control_frame, text=self.current_date.strftime("%B %Y"), font=("Arial", 12))
        self.date_label.pack(side=tk.LEFT, padx=5)
        
        self.next_button = tk.Button(control_frame, text=">", command=self.next_month)
        self.next_button.pack(side=tk.LEFT, padx=5)
        
        # Show/Update Statistics button.
        tk.Button(control_frame, text="Show Statistics", command=self.show_stats).pack(side=tk.LEFT, padx=10)
        
        # Create the treeview to display statistics.
        columns = ("Employee", 
           "Normal Shifts", "Normal Hours", 
           "Festive Shifts", "Festive Hours", 
           "Target Hours", "Accumulated Hours")
        self.tree = ttk.Treeview(self, columns=columns, show="headings")

        for col in columns:
            self.tree.heading(col, text=col)
        self.tree.pack(fill=tk.BOTH, expand=True, pady=10)
        
    def prev_month(self):
        # Subtract one month from self.current_date.
        year = self.current_date.year
        month = self.current_date.month
        if month == 1:
            new_year = year - 1
            new_month = 12
        else:
            new_year = year
            new_month = month - 1
        self.current_date = datetime.date(new_year, new_month, 1)
        self.date_label.config(text=self.current_date.strftime("%B %Y"))
        # Optionally, clear the statistics display or wait until "Show Statistics" is pressed.
    
    def next_month(self):
        # Add one month to self.current_date.
        year = self.current_date.year
        month = self.current_date.month
        if month == 12:
            new_year = year + 1
            new_month = 1
        else:
            new_year = year
            new_month = month + 1
        self.current_date = datetime.date(new_year, new_month, 1)
        self.date_label.config(text=self.current_date.strftime("%B %Y"))
        # Optionally, clear the statistics display or wait until "Show Statistics" is pressed.
    
    def show_stats(self):
        # 1. Update the accumulated_hours in the DB (optional, if desired).
        self.db_manager.update_employee_statistics()

        # 2. Gather current month's shifts.
        year = self.current_date.year
        month = self.current_date.month
        shifts = self.db_manager.get_shifts_for_month(year, month)

        # 3. Get festivity info for the month: { 'YYYY-MM-DD': True/False for is_working_day }
        festivity_map = self.db_manager.get_festivities_for_month(year, month)

        # 4. Prepare stats dict.
        #    We'll store per-employee:
        #        normal_shifts, normal_hours, fest_shifts, fest_hours, target_hours, accumulated_hours
        stats = {}
        
        # 5. Gather employees to fill in target/accumulated data even if they have 0 shifts
        all_employees = {e["id"]: e for e in self.db_manager.get_employees()}

        # 6. SHIFT loop: categorize each shift
        for shift_id, shift_date, shift_type, emp_id, emp_name in shifts:
            # If not in stats, initialize
            if emp_id not in stats:
                stats[emp_id] = {
                    "name": emp_name,
                    "normal_shifts": 0,
                    "normal_hours": 0,
                    "fest_shifts": 0,
                    "fest_hours": 0,
                    "target": all_employees[emp_id]["target_hours"],
                    "accumulated": all_employees[emp_id]["accumulated_hours"],
                }
            # Determine shift duration from settings (morning/afternoon/night).
            # You already have something like:
            duration_map = {
                "Morning": int(self.db_manager.get_setting("duration_morning")),
                "Afternoon": int(self.db_manager.get_setting("duration_afternoon")),
                "Night": int(self.db_manager.get_setting("duration_night"))
            }
            shift_hours = duration_map.get(shift_type, 8)

            # Check if date is festive
            # If it's in festivity_map but is_working_day=False => it's a "non-working" festivity.
            is_festive = False
            if shift_date in festivity_map:
                # If is_working_day == False, treat as festive
                is_festive = not festivity_map[shift_date]

            if is_festive:
                stats[emp_id]["fest_shifts"] += 1
                stats[emp_id]["fest_hours"] += shift_hours
            else:
                stats[emp_id]["normal_shifts"] += 1
                stats[emp_id]["normal_hours"] += shift_hours

        # 7. For employees who had no shifts in this month, still fill data
        for emp_id, emp_data in all_employees.items():
            if emp_id not in stats:
                stats[emp_id] = {
                    "name": emp_data["name"],
                    "normal_shifts": 0,
                    "normal_hours": 0,
                    "fest_shifts": 0,
                    "fest_hours": 0,
                    "target": emp_data["target_hours"],
                    "accumulated": emp_data["accumulated_hours"],
                }

        # 8. Clear existing rows from the Treeview
        for row in self.tree.get_children():
            self.tree.delete(row)

        # 9. Insert updated rows
        for emp_id, s in stats.items():
            self.tree.insert("", "end", values=(
                s["name"],
                s["normal_shifts"], s["normal_hours"],
                s["fest_shifts"],   s["fest_hours"],
                s["target"],        s["accumulated"]
            ))


# =============================================================================
# Employee Selection Dialog
# =============================================================================
class EmployeeSelectionDialog(tk.Toplevel):
    def __init__(self, master, db_manager, current_emp_id):
        super().__init__(master)
        self.title("Select Employee")
        self.db_manager = db_manager
        self.result = None
        tk.Label(self, text="Select Employee:").pack(pady=10)
        self.employee_var = tk.StringVar()
        employees = self.db_manager.get_employees()
        self.employee_options = {f"{emp['id']}: {emp['name']}": emp['id'] for emp in employees}
        self.combo = ttk.Combobox(self, textvariable=self.employee_var,
                                  values=list(self.employee_options.keys()), state="readonly")
        self.combo.pack(pady=10)
        # Set the current employee as the default selection
        for key, emp_id in self.employee_options.items():
            if emp_id == current_emp_id:
                self.combo.set(key)
                break
        tk.Button(self, text="OK", command=self.on_ok).pack(pady=10)

    def on_ok(self):
        emp_key = self.employee_var.get()
        if emp_key:
            self.result = self.employee_options[emp_key]
        self.destroy()


# =============================================================================
# Run the Application
# =============================================================================
if __name__ == "__main__":
    app = ShiftSchedulerApp()
    app.mainloop()