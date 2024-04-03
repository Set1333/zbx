import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from pyzabbix import ZabbixAPI
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import tkinter.messagebox
import json


class ZabbixExportGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Zabbix Triggers Export")
        self.root.geometry("600x400")

        # Load configuration
        self.load_config()

        # Zabbix connection variables
        self.zabbix_server = tk.StringVar(value=self.config.get('server', ''))
        self.zabbix_user = tk.StringVar(value=self.config.get('user', ''))
        self.zabbix_password = tk.StringVar(value=self.config.get('password', ''))

        # Trigger export variables
        self.group_id = tk.StringVar(value=self.config.get('group_id', ''))
        self.server_name = tk.StringVar(value=self.config.get('server_name', ''))
        self.due_date = tk.StringVar(value=self.config.get('due_date', ''))
        self.start_date = tk.StringVar(value=self.config.get('start_date', ''))
        self.end_date = tk.StringVar(value=self.config.get('end_date', ''))
        self.errors_only = tk.BooleanVar(value=self.config.get('errors_only', False))

        # User email variables
        self.user_ids = tk.StringVar(value=self.config.get('user_ids', ''))
        self.fetch_all_attributes = tk.BooleanVar(value=self.config.get('fetch_all_attributes', False))
        self.user_emails = {}

        # Zabbix connection frame
        connection_frame = ttk.LabelFrame(self.root, text="Zabbix Connection")
        connection_frame.pack(pady=10)

        ttk.Label(connection_frame, text="Server URL:").grid(row=0, column=0, sticky="w")
        ttk.Entry(connection_frame, textvariable=self.zabbix_server).grid(row=0, column=1)

        ttk.Label(connection_frame, text="Username:").grid(row=1, column=0, sticky="w")
        ttk.Entry(connection_frame, textvariable=self.zabbix_user).grid(row=1, column=1)

        ttk.Label(connection_frame, text="Password:").grid(row=2, column=0, sticky="w")
        ttk.Entry(connection_frame, textvariable=self.zabbix_password, show="*").grid(row=2, column=1)

        # Trigger export frame
        export_frame = ttk.LabelFrame(self.root, text="Trigger Export Parameters")
        export_frame.pack(pady=10)

        ttk.Label(export_frame, text="Group ID:").grid(row=0, column=0, sticky="w")
        ttk.Entry(export_frame, textvariable=self.group_id).grid(row=0, column=1)

        ttk.Label(export_frame, text="Server Name:").grid(row=1, column=0, sticky="w")
        ttk.Entry(export_frame, textvariable=self.server_name).grid(row=1, column=1)

        ttk.Label(export_frame, text="Due Date:").grid(row=2, column=0, sticky="w")
        self.due_date_entry = DateEntry(export_frame, textvariable=self.due_date, date_pattern="yyyy-mm-dd")
        self.due_date_entry.grid(row=2, column=1)

        ttk.Label(export_frame, text="Start Date:").grid(row=3, column=0, sticky="w")
        self.start_date_entry = DateEntry(export_frame, textvariable=self.start_date, date_pattern="yyyy-mm-dd")
        self.start_date_entry.grid(row=3, column=1)

        ttk.Label(export_frame, text="End Date:").grid(row=4, column=0, sticky="w")
        self.end_date_entry = DateEntry(export_frame, textvariable=self.end_date, date_pattern="yyyy-mm-dd")
        self.end_date_entry.grid(row=4, column=1)

        if self.start_date.get() or self.end_date.get():
            ttk.Checkbutton(export_frame, text="Fetch All Attributes", variable=self.fetch_all_attributes).grid(row=5,
                                                                                                                columnspan=2)

        # Checkbutton for fetching all attributes
        email_frame = ttk.LabelFrame(self.root, text="User Emails")
        email_frame.pack(pady=10)

        ttk.Label(email_frame, text="User IDs:").grid(row=0, column=0, sticky="w")
        ttk.Entry(email_frame, textvariable=self.user_ids).grid(row=0, column=1)

        ttk.Button(email_frame, text="Fetch Emails", command=self.fetch_emails).grid(row=0, column=2)

        # Export button
        ttk.Button(self.root, text="Export Triggers", command=self.export_triggers).pack(pady=10)

    def load_config(self):
        try:
            with open('config.json', 'r') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            self.config = {}
        except json.JSONDecodeError:
            self.config = {}  # Handle corrupted config file gracefully

    def save_config(self):
        with open('config.json', 'w') as f:
            json.dump(self.config, f)

    def fetch_emails(self):
        zabbix_server = self.zabbix_server.get()
        zabbix_user = self.zabbix_user.get()
        zabbix_password = self.zabbix_password.get()

        try:
            zapi = ZabbixAPI(zabbix_server)
            zapi.login(zabbix_user, zabbix_password)
        except Exception as e:
            tk.messagebox.showerror("Error", f"Failed to log in to Zabbix: {e}")
            return

        user_ids = self.user_ids.get().split(',')
        self.user_emails = {}

        fetch_all_attributes = self.fetch_all_attributes.get()

        for user_id in user_ids:
            user_filter = {'userid': user_id}

            if fetch_all_attributes:
                user = zapi.user.get(output='extend', filter=user_filter)
            else:
                user = zapi.user.get(output=['email'], filter=user_filter)

            if user and user[0].get('email'):
                self.user_emails[user_id] = user[0]['email']
            else:
                tk.messagebox.showwarning("Warning", f"Email not found for user ID: {user_id}")

        # Display fetched emails in a popup window
        self.show_emails_popup()

        self.save_config()

    def show_emails_popup(self):
        popup = tk.Toplevel(self.root)
        popup.title("Fetched Emails")

        listbox = tk.Listbox(popup)
        for user_id, email in self.user_emails.items():
            listbox.insert(tk.END, f"User ID: {user_id}, Email: {email}")

        listbox.pack(padx=10, pady=10)

    def export_triggers(self):
        zabbix_server = self.zabbix_server.get()
        zabbix_user = self.zabbix_user.get()
        zabbix_password = self.zabbix_password.get()

        group_id = self.group_id.get() if self.group_id.get() else None
        server_name = self.server_name.get() if self.server_name.get() else None
        due_date = self.due_date.get() if self.due_date.get() else None
        start_date = self.start_date.get() if self.start_date.get() else None
        end_date = self.end_date.get() if self.end_date.get() else None
        errors_only = self.errors_only.get()

        zapi = ZabbixAPI(zabbix_server)
        zapi.login(zabbix_user, zabbix_password)

        # Function to get triggers based on group ID or server name and due date
        def get_triggers(zapi, group_id=None, server_name=None, due_date=None, errors_only=False):
            filter_dict = {}
            if group_id is not None:
                filter_dict['group'] = group_id
            if server_name:
                filter_dict['host'] = server_name
            if due_date:
                filter_dict['value'] = due_date

            if errors_only:
                filter_dict['value'] = '1'

            return zapi.trigger.get(output=['triggerid', 'description', 'value'], filter=filter_dict)

        # Convert due date string to epoch time if provided
        due_date_epoch = None
        if due_date:
            due_date_dt = datetime.strptime(due_date, '%Y-%m-%d')
            due_date_epoch = int(due_date_dt.timestamp())

        # Get list of triggers based on due date
        if due_date:
            triggers = get_triggers(zapi, group_id=group_id, server_name=server_name, due_date=due_date_epoch,
                                    errors_only=errors_only)
        # Get list of triggers based on date range
        else:
            start_date_epoch = None
            end_date_epoch = None
            if start_date:
                start_date_dt = datetime.strptime(start_date, '%Y-%m-%d')
                start_date_epoch = int(start_date_dt.timestamp())
            if end_date:
                end_date_dt = datetime.strptime(end_date, '%Y-%m-%d')
                end_date_epoch = int((end_date_dt + timedelta(days=1)).timestamp())

            triggers = get_triggers(zapi, group_id=group_id, server_name=server_name, start_date=start_date_epoch,
                                    end_date=end_date_epoch, errors_only=errors_only)

        # Create an Excel workbook
        wb = Workbook()
        ws = wb.active

        # Write headers
        ws.append(['Trigger ID', 'Description', 'Value', 'User Email'])

        # Apply color formatting to triggers with value = 1
        red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        for trigger in triggers:
            trigger_id = trigger['triggerid']
            description = trigger['description']
            value = trigger['value']
            user_id = trigger.get('userid', '')
            user_email = self.user_emails.get(user_id, '') if user_id else ''
            ws.append([trigger_id, description, value, user_email])
            if value == '1':
                ws[f'C{ws.max_row}'].fill = red_fill

        # Save the workbook
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            wb.save(filename)

        # Log out from Zabbix API
        zapi.user.logout()

        # Save the configuration
        self.config['server'] = zabbix_server
        self.config['user'] = zabbix_user
        self.config['password'] = zabbix_password
        self.config['group_id'] = group_id
        self.config['server_name'] = server_name
        self.config['due_date'] = due_date
        self.config['start_date'] = start_date
        self.config['end_date'] = end_date
        self.config['errors_only'] = errors_only
        self.config['user_ids'] = self.user_ids.get()
        self.config['fetch_all_attributes'] = self.fetch_all_attributes.get()
        self.save_config()


if __name__ == "__main__":
    root = tk.Tk()
    app = ZabbixExportGUI(root)
    root.mainloop()
