from tkinter import ttk
from tkinter.messagebox import showerror
from tkinter.filedialog import askdirectory
from PIL import Image, ImageTk, ImageEnhance
from ttkthemes import ThemedTk
import os
import socket
import platform
import subprocess
import tkinter as tk
from PIL import Image, ImageTk
from tkinter import ttk, filedialog
from PIL import ImageGrab
from PIL.ImageDraw import ImageDraw
from io import BytesIO
from pywifi import PyWiFi, const
from tkinter import simpledialog
from tkinter import Toplevel
from PIL import Image, ImageTk, ImageEnhance
from tkinter import Frame
from tkinter import PhotoImage
from tkinter import TOP, Entry, Label, Button, LEFT, RIGHT, X, YES
from ttkthemes import ThemedTk
import psutil
import ctypes
import winreg
import wmi
import datetime
import re
import os
import requests
import gspread
import win32print
import win32com.client
import subprocess
import win32api
from tkinter import *

# Inicializa el objeto WMI
c = wmi.WMI()

def is_laptop():
    battery = psutil.sensors_battery()
    return battery is not None

def get_laptop_info():
    for computer_system in c.Win32_ComputerSystem():
        manufacturer = computer_system.Manufacturer
        model = computer_system.Model

    for bios in c.Win32_BIOS():
        serial_number = getattr(bios, "SerialNumber", "No disponible")

    return f"Marca: {manufacturer} - Modelo: {model} - Número de serie: {serial_number}"

def get_public_ip():
    try:
        response = requests.get("https://api.ipify.org")
        ip_address = response.text
    except:
        ip_address = "No se pudo obtener la dirección IP pública"
    return ip_address

def get_system_info():
    for os_info in c.Win32_OperatingSystem():
        name = os_info.Caption
        version = os_info.Version
        build_number = os_info.BuildNumber
        install_date = os_info.InstallDate

        # Extraer la fecha y hora sin la zona horaria
        match = re.match(r"(\d{8}\d{6})\.\d{6}[+-]\d{3}", install_date)
        if match:
            stripped_install_date = match.group(1)
            formatted_install_date = datetime.datetime.strptime(stripped_install_date, "%Y%m%d%H%M%S").strftime("%d-%m-%Y")
        else:
            formatted_install_date = "No se pudo determinar la fecha de instalación"

        try:
            last_update = c.Win32_QuickFixEngineering()[-1]
            last_update_id = last_update.HotFixID
            last_update_date = last_update.InstalledOn
        except IndexError:
            last_update_id = "No se encontraron actualizaciones"
            last_update_date = "No se encontraron actualizaciones"

        return f"Sistema Operativo: {name} {version} Build {build_number} - Última actualización: {last_update_id} - Fecha de instalación: {formatted_install_date} - Fecha de la última actualización: {last_update_date}"

def get_teamviewer_id():
    try:
        registry_paths = [
            r"SOFTWARE\WOW6432Node\TeamViewer",
            r"SOFTWARE\TeamViewer"
        ]

        for reg_path in registry_paths:
            try:
                with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                    value, _ = winreg.QueryValueEx(key, "ClientID")
                    return value
            except FileNotFoundError:
                pass

        return "Teamviewer no instalado"

    except Exception as e:
        print(f"Error al obtener el ID de TeamViewer: {e}")
        return "Teamviewer no instalado"

def export_to_txt(info):
    file_name = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if file_name:
        with open(file_name, 'w') as f:
            f.write(info)

def get_processor_info():
    for cpu in c.Win32_Processor():
        return f"Marca y Modelo del Procesador: {cpu.Name.strip()}"

def get_ram_info():
    total_memory = int(psutil.virtual_memory().total / (1024 * 1024))
    return f"Memoria RAM total: {total_memory} MB"

import re

def print_debug_info(drive, drive_info, disk, model, manufacturer, serial_number, disk_type):
    print(f"Unidad: {drive}")
    print(f"Drive Info: {drive_info}")
    print(f"Disco: {disk}")
    print(f"Modelo: {model}")
    print(f"Fabricante: {manufacturer}")
    print(f"Número de serie: {serial_number}")
    print(f"Tipo de disco: {disk_type}")
    print("\n")


def get_disk_info():
    disk_info = ""

    def is_google_drive(drive_label):
        return "Google Drive" in drive_label or "Drive File Stream" in drive_label

    try:
        for disk in c.Win32_DiskDrive():
            capabilities = disk.Capabilities
            if capabilities and 4 in capabilities:
                disk_type = "SSD"
            else:
                disk_type = "HDD"
            manufacturer = disk.Manufacturer.strip()
            model = disk.Model.strip()
            serial_number = disk.SerialNumber.strip()

            for partition in disk.associators("Win32_DiskDriveToDiskPartition"):
                for logical_disk in partition.associators("Win32_LogicalDiskToPartition"):
                    drive = logical_disk.DeviceID
                    drive_label = logical_disk.VolumeName

                    if is_google_drive(drive_label):
                        disk_info += f"Google Drive for Desktop INSTALADO en la unidad {drive}\n"
                        continue

                    drive_info = psutil.disk_usage(drive)
                    total_gb = int(drive_info.total) // (1024**3)
                    free_gb = int(drive_info.free) // (1024**3)

                    disk_info += f"{drive}: {manufacturer} {model} ({disk_type}) - Número de serie: {serial_number} - Espacio total: {total_gb} GB - Espacio libre: {free_gb} GB\n"

    except Exception as e:
        print(f"Error al obtener información del disco: {e}")

    return disk_info.strip()

def get_wifi_info():
    try:
        wifi_info = [iface for iface in psutil.net_if_stats().keys() if "Wireless" in iface or "Wi-Fi" in iface]
        if wifi_info:
            wifi_adapter = wifi_info[0]
            mac_address = psutil.net_if_addrs()[wifi_adapter][0].address
            ip_address = ""
            for snic in psutil.net_if_addrs()[wifi_adapter]:
                if snic.family == socket.AF_INET:
                    ip_address = snic.address

            ssid = subprocess.check_output(["netsh", "wlan", "show", "interfaces"]).decode("utf-8", errors="ignore")
            ssid = ssid.partition("SSID")[2].partition(":")[2].strip().partition("\n")[0]
            return f"SSID: {ssid} - MAC Address: {mac_address} - IP Address: {ip_address}"
    except Exception as e:
        print(e)
        return "Información de la placa Wi-Fi no disponible"

def get_ethernet_info():
    try:
        ethernet_info = [iface for iface in psutil.net_if_stats().keys() if "Ethernet" in iface or "eth" in iface]
        if ethernet_info:
            ethernet_adapter = ethernet_info[0]
            mac_address = psutil.net_if_addrs()[ethernet_adapter][0].address
            ip_address = ""
            for snic in psutil.net_if_addrs()[ethernet_adapter]:
                if snic.family == socket.AF_INET:
                    ip_address = snic.address
            return f"MAC Address: {mac_address} - IP Address: {ip_address}"
    except Exception as e:
        print(e)
        return "Información de la placa Ethernet no disponible"

def ask_user_info(root):
    user_info = {}

    fields = ["Nombre", "Apellido", "Cargo", "Empresa para la cual trabaja", "Sede de la empresa"]
    entries = []

    def submit_info():
        for i, field in enumerate(fields):
            user_info[field] = entries[i].get()
        popup.destroy()

        # Muestra nuevamente la ventana principal cuando se cierra la ventana de información del usuario
        root.deiconify()

    # Oculta la ventana principal mientras se muestra la ventana de información del usuario
    root.withdraw()

    popup = Toplevel()
    popup.title("Información del usuario")

    for field in fields:
        row = Frame(popup)
        row.pack(side=TOP, fill=X, padx=5, pady=5)
        label = Label(row, text=field, width=25)
        label.pack(side=LEFT)
        entry = Entry(row)
        entry.pack(side=RIGHT, expand=YES, fill=X)
        entries.append(entry)

    submit_button = Button(popup, text="Enviar", command=submit_info)
    submit_button.pack(side=RIGHT, padx=5, pady=5)

    popup.grab_set()
    popup.wait_window()
    
    return "\n".join([f"{k}: {v}" for k, v in user_info.items()])

import win32print

def get_printers_and_ports():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)
    printer_info = []
    non_physical_printers = ["PDFCreator", "Fax", "OneNote", "Microsoft Print to PDF", "Microsoft XPS Document Writer"]

    for printer in printers:
        name = printer[2]
        is_physical = True

        # Verifica si el nombre de la impresora coincide con alguna de las impresoras no físicas
        for non_physical_printer in non_physical_printers:
            if non_physical_printer.lower() in name.lower():
                is_physical = False
                break

        if is_physical:
            hPrinter = win32print.OpenPrinter(name)
            printer_details = win32print.GetPrinter(hPrinter, 2)
            win32print.ClosePrinter(hPrinter)
            port = printer_details['pPortName']
            printer_info.append(f"Impresora: {name}, Puerto: {port}")

    return "\n".join(printer_info)

def export_to_txt(info, user_info):
    name_match = re.search(r'Nombre:\s+(\w+)\s+Apellido:\s+(\w+)', user_info)
    full_name = f"{name_match.group(1)}_{name_match.group(2)}" if name_match else "Info"
    today = datetime.datetime.today().strftime('%Y%m%d')
    suggested_filename = f"{full_name}_{today}"
    file_name = filedialog.asksaveasfilename(initialfile=suggested_filename, defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if file_name:
        with open(file_name, 'w') as f:
            f.write(info)

# Función principal que recopila la información y la muestra en una ventana
def main():
    root = ThemedTk(theme="arc")
    root.title("Recopilacion dato de sistema - Tiknology")
    frame = ttk.Frame(root, padding=10)
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    user_info = ask_user_info(root)
    if is_laptop():
        device_type = f"Notebook - {get_laptop_info()}"
    else:
        device_type = "No se detecto que este equipo sea una Notebook"    
    info = f"Tipo de equipo: {'Notebook' if is_laptop() else 'PC de escritorio'}\n"
    info += f"{get_system_info()}\n"
    info += f"{get_processor_info()}\n"
    info += f"{get_ram_info()}\n"
    info += f"{get_disk_info()}\n"
    info += f"Impresoras instaladas:\n{get_printers_and_ports()}\n"
    wifi_info = f"WiFi: {get_wifi_info()}\n"
    ethernet_info = f"Ethernet: {get_ethernet_info()}\n"
    teamviewer_info = f"Teamviewer ID: {get_teamviewer_id()}\n"
    info += f"Tipo de equipo: {device_type}\n" 
    publicIP_info = f"Dirección IP pública detectada: {get_public_ip()}\n"

    #resto codigo

    label = ttk.Label(frame, text=info, justify=tk.LEFT, font=("Arial", 10))
    label.grid(column=0, row=0, sticky=(tk.W, tk.N))

    label_wifi = ttk.Label(frame, text=wifi_info, justify=tk.LEFT, foreground="blue", font=("Arial", 10))
    label_wifi.grid(column=0, row=1, sticky=(tk.W, tk.N))

    label_ethernet = ttk.Label(frame, text=ethernet_info, justify=tk.LEFT, foreground="blue", font=("Arial", 10))
    label_ethernet.grid(column=0, row=2, sticky=(tk.W, tk.N))

    label_publicIP = ttk.Label(frame, text=publicIP_info, justify=tk.LEFT, foreground="blue", font=("Arial", 14))
    label_publicIP.grid(column=0, row=3, sticky=(tk.W, tk.N))

    label_teamviewer = ttk.Label(frame, text=teamviewer_info, justify=tk.LEFT, foreground="red", font=("Arial", 14))
    label_teamviewer.grid(column=0, row=4, sticky=(tk.W, tk.N))

    def save_to_txt():
        export_to_txt(info + wifi_info + ethernet_info + teamviewer_info, user_info)

    style = ttk.Style()
    style.map('Green.TButton',
            foreground=[('active', 'black'), ('disabled', 'gray')],
            background=[('active', 'green'), ('disabled', 'light gray')])
    style.configure('Green.TButton', foreground='black', background='green')
    button2 = ttk.Button(frame, text='Exportar TXT', command=save_to_txt, style='Green.TButton', foreground='black', background='green')
    button2.grid(column=3, row=3, pady=5, sticky=(tk.W, tk.N))


    root.mainloop()

if __name__ == '__main__':
    main()