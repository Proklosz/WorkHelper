import tkinter as tk
from tkinter import filedialog, messagebox
import os
import shutil
import win32com.client

def browse_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_folder_path.delete(0, tk.END)
        entry_folder_path.insert(0, folder_path)

def process_folder():
    folder_path = entry_folder_path.get()
    file_name = entry_file_name.get()

    if folder_path and file_name:
        # Creating the PowerShell script
        ps_script_content = f"""Set-Location "{folder_path}" 
Start-Process "{folder_path}" 
code -n .
Start-Process powershell.exe
"""
        ps_script_path = f".\\{file_name}.ps1"
        with open(ps_script_path, "w") as ps_script_file:
            ps_script_file.write(ps_script_content)

        print(f"PowerShell script created at: {ps_script_path}")

        # Creating a shortcut on the desktop
        desktop_path = os.path.expanduser("~\\Desktop")
        shortcut_path = os.path.join(desktop_path, f"{file_name}.lnk")
        target_path = os.path.abspath(ps_script_path)

        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(shortcut_path)
        shortcut.TargetPath = "powershell.exe"  # Set the target to PowerShell
        shortcut.Arguments = f"-ExecutionPolicy Bypass -File {target_path}"
        shortcut.WorkingDirectory = os.path.dirname(target_path)
        shortcut.save()

        print(f"Shortcut created at: {shortcut_path}")
    else:
        messagebox.showerror("Error", "Please provide a folder path and file name.")

# Create the main window
root = tk.Tk()
root.title("Folder Path Input Example")

# Create labels for input fields
label_folder_path = tk.Label(root, text="Folder Path:")
label_file_name = tk.Label(root, text="File Name:")  # New label

# Create input fields and buttons
entry_folder_path = tk.Entry(root, width=50)
entry_file_name = tk.Entry(root, width=50)  # New input field
button_browse = tk.Button(root, text="Browse", command=browse_folder)
button_process = tk.Button(root, text="Process Folder", command=process_folder)

# Arrange the UI components using grid layout
label_folder_path.grid(row=0, column=0, padx=10, pady=10, sticky="E")
entry_folder_path.grid(row=0, column=1, padx=10, pady=10)
button_browse.grid(row=0, column=2, padx=10, pady=10)

label_file_name.grid(row=1, column=0, padx=10, pady=10, sticky="E")  # New label position
entry_file_name.grid(row=1, column=1, padx=10, pady=10)  # New input field position

button_process.grid(row=2, columnspan=3, padx=10, pady=10)

# Start the GUI event loop
root.mainloop()