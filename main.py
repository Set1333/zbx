import tkinter as tk
from tkinter import messagebox
from pyzabbix import ZabbixAPI, ZabbixAPIException
import datetime
import openpyxl
import ssl

# Отключение проверки сертификата SSL (нужно только для самоподписанных сертификатов)
ssl._create_default_https_context = ssl._create_unverified_context

def get_triggers(zabbix_api, group=None, host=None, period_start=None, period_end=None):
    # Формируем фильтры для запроса триггеров
    filter_params = {}
    if group:
        filter_params['group'] = group
    if host:
        filter_params['host'] = host

    # Получаем список триггеров с учетом фильтров
    try:
        triggers = zabbix_api.trigger.get(
            output=['description', 'lastchange', 'priority', 'value', 'hosts'],
            selectHosts=['host'],
            filter=filter_params,
            expandDescription=1,
            monitored=1
        )
    except ZabbixAPIException as e:
        messagebox.showerror('Error', f'Error fetching triggers: {e}')
        return []

    # Фильтрация триггеров по временному периоду
    filtered_triggers = []
    for trigger in triggers:
        last_change_timestamp = int(trigger['lastchange'])
        try:
            last_change_date = datetime.datetime.fromtimestamp(last_change_timestamp)
        except OSError as e:
            print(f"Error converting timestamp for trigger: {trigger['description']}, Error: {e}")
            continue

        if period_start and period_start > last_change_date:
            continue
        if period_end and period_end < last_change_date:
            continue

        filtered_triggers.append(trigger)

    return filtered_triggers

def fetch_triggers():
    url = url_entry.get()
    user = user_entry.get()
    password = password_entry.get()
    group = group_entry.get()
    host = host_entry.get()
    start_date = start_date_entry.get()
    end_date = end_date_entry.get()

    if not url or not user or not password:
        messagebox.showwarning('Warning', 'Please enter Zabbix URL, username, and password.')
        return

    zabbix_api = ZabbixAPI(url)
    zabbix_api.use_ssl = True  # Используем SSL поддержку

    try:
        zabbix_api.login(user, password)
    except ZabbixAPIException as e:
        messagebox.showerror('Error', f'Error logging in: {e}')
        return

    try:
        period_start = datetime.datetime.strptime(start_date, '%Y-%m-%d') if start_date else None
        period_end = datetime.datetime.strptime(end_date, '%Y-%m-%d') if end_date else None
    except ValueError:
        messagebox.showerror('Error', 'Invalid date format. Please use YYYY-MM-DD.')
        return

    triggers = get_triggers(zabbix_api, group=group, host=host, period_start=period_start, period_end=period_end)

    if triggers:
        save_to_excel(triggers)
        messagebox.showinfo('Information', 'Data saved to Excel file.')
    else:
        messagebox.showinfo('Information', 'No triggers found matching the criteria.')

    save_settings()


def save_to_excel(triggers):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Description', 'Last Change', 'Priority', 'Value', 'Hosts'])

    for trigger in triggers:
        # Преобразуем значение 'hosts' в строку для сохранения в Excel
        hosts_str = ', '.join([host['host'] for host in trigger['hosts']])

        # Преобразуем timestamp в формат даты и времени
        last_change_date = datetime.datetime.fromtimestamp(int(trigger['lastchange']))

        ws.append([trigger['description'], last_change_date, trigger['priority'], trigger['value'], hosts_str])

    wb.save('triggers.xlsx')

def save_settings():
    with open('settings.txt', 'w') as f:
        f.write(f"URL={url_entry.get()}\n")
        f.write(f"User={user_entry.get()}\n")
        f.write(f"Password={password_entry.get()}\n")
        f.write(f"Group={group_entry.get()}\n")
        f.write(f"Host={host_entry.get()}\n")
        f.write(f"Start Date={start_date_entry.get()}\n")
        f.write(f"End Date={end_date_entry.get()}\n")

def load_settings():
    try:
        with open('settings.txt', 'r') as f:
            settings = {}
            for line in f:
                key, value = line.strip().split('=')
                settings[key] = value

            url_entry.insert(0, settings.get('URL', ''))
            user_entry.insert(0, settings.get('User', ''))
            password_entry.insert(0, settings.get('Password', ''))  # Загружаем пароль
            group_entry.insert(0, settings.get('Group', ''))
            host_entry.insert(0, settings.get('Host', ''))
            start_date_entry.insert(0, settings.get('Start Date', ''))  # Загружаем дату начала
            end_date_entry.insert(0, settings.get('End Date', ''))  # Загружаем дату окончания
    except FileNotFoundError:
        pass

# Создание основного окна
root = tk.Tk()
root.title('Zabbix Trigger Fetcher')

# Создание и размещение элементов управления
tk.Label(root, text="Zabbix URL:").pack()
url_entry = tk.Entry(root)
url_entry.pack()

tk.Label(root, text="Username:").pack()
user_entry = tk.Entry(root)
user_entry.pack()

tk.Label(root, text="Password:").pack()
password_entry = tk.Entry(root, show="*")
password_entry.pack()

tk.Label(root, text="Group Name:").pack()
group_entry = tk.Entry(root)
group_entry.pack()

tk.Label(root, text="Host Name:").pack()
host_entry = tk.Entry(root)
host_entry.pack()

tk.Label(root, text="Start Date (YYYY-MM-DD):").pack()
start_date_entry = tk.Entry(root)
start_date_entry.pack()

tk.Label(root, text="End Date (YYYY-MM-DD):").pack()
end_date_entry = tk.Entry(root)
end_date_entry.pack()

fetch_button = tk.Button(root, text="Fetch Triggers", command=fetch_triggers)
fetch_button.pack()

load_settings()

root.mainloop()
