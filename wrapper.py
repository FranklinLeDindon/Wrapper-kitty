import shutil
import tempfile
import os
import ctypes
import subprocess
import customtkinter as ctk
import sys
import time


def is_admin():
    """Check if the script is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def run_as_admin():
    """Re-run the script with administrative privileges."""
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(sys.argv), None, 1)


# Function to import the PowerShell module and invoke the command
def import_and_invoke(output_box, only_filtered_csv=False):
    # Define commands to set execution policy and import module
    set_policy_cmd = "Set-ExecutionPolicy Bypass -Scope Process -Force"
    import_module_cmd = r"Import-Module 'C:\Users\f0r4\PycharmProjects\DefendeRsGUI\DefendeRs1\DefendeRs.psm1'"
    invoke_command = "Invoke-DefendeRs"

    # Move non-matching files if only filtered CSVs are requested
    temp_dir = None
    original_files = []
    if only_filtered_csv:
        temp_dir, original_files = move_non_matching_files()

    # Add a small delay to ensure PowerShell detects file changes
    time.sleep(1)

    # Complete the command string
    commands = f"{set_policy_cmd}; {import_module_cmd}; {invoke_command}"

    # Execute the command in PowerShell, ensuring we set the working directory
    process = subprocess.Popen(["powershell", "-Command", commands],
                               cwd=r"C:\Users\f0r4\PycharmProjects\DefendeRsGUI\DefendeRs1\lists",
                               # Set working directory
                               stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE)

    output, error = process.communicate()

    # Restore files after execution
    if temp_dir:
        restore_files(temp_dir, original_files)

    # Output results to GUI
    output_box.delete(1.0, ctk.END)
    if process.returncode == 0:
        output_box.insert(ctk.END, "Command executed successfully:\n")
        apply_colored_output(output.decode(), output_box)
    else:
        output_box.insert(ctk.END, "Error executing command:\n")
        output_box.insert(ctk.END, f"Return code: {process.returncode}\n")
        output_box.insert(ctk.END, "Output:\n")
        apply_colored_output(output.decode(), output_box)
        output_box.insert(ctk.END, "Error:\n")
        apply_colored_output(error.decode(), output_box)


def filter_csv(csv=0):
    if csv == 1:
        prefix = "finding_list_dod_microsoft_windows_10"
        csv_dir = r"C:\Users\f0r4\PycharmProjects\DefendeRsGUI\DefendeRs1\lists"
        allowed_files = {"finding_list_0x6d69636b_machine.csv", "finding_list_0x6d69636b_user.csv"}
        move_csv(prefix, csv_dir, allowed_files)


def move_csv(prefix, csv_dir, allowed_files):
    temp_dir = tempfile.mkdtemp()
    original_files = []

    for file in os.listdir(csv_dir):
        file_path = os.path.join(csv_dir, file)

        # Skip directories
        if not os.path.isfile(file_path):
            continue

        # Check if file should be kept
        if not (file.startswith(prefix) or file in allowed_files) or not file.endswith(".csv"):
            print(f"Moving file: '{file}'")  # Debug: Print file being moved
            shutil.move(file_path, temp_dir)
            original_files.append(file)

    return temp_dir, original_files


# Function to temporarily move non-matching CSV files out of the directory
def move_non_matching_files():
    csv_dir = r"C:\Users\f0r4\PycharmProjects\DefendeRsGUI\DefendeRs1\lists"  # Ensure this path is correct
    prefix = "finding_list_bsi_sisyphus_windows_10"
    # Specify the additional files to include in the allowed list
    allowed_files = {"finding_list_0x6d69636b_machine.csv", "finding_list_0x6d69636b_user.csv"}

    temp_dir = tempfile.mkdtemp()
    original_files = []

    print(temp_dir)  # Debug: List all files in the directory

    for file in os.listdir(csv_dir):
        file_path = os.path.join(csv_dir, file)

        # Skip directories
        if not os.path.isfile(file_path):
            continue

        # Check if file should be kept
        if not (file.startswith(prefix) or file in allowed_files) or not file.endswith(".csv"):
            print(f"Moving file: '{file}'")  # Debug: Print file being moved
            shutil.move(file_path, temp_dir)
            original_files.append(file)

    return temp_dir, original_files


# Function to restore files after running the command
def restore_files(temp_dir, original_files):
    csv_dir = r"C:\Users\f0r4\PycharmProjects\DefendeRsGUI\DefendeRs1\Lists"  # Ensure this path is correct

    for file in original_files:
        source_path = os.path.join(temp_dir, file)
        destination_path = os.path.join(csv_dir, file)

        # If a file with the same name already exists, remove it first
        if os.path.exists(destination_path):
            print(f"Overwriting existing file: {destination_path}")
            os.remove(destination_path)

        # Move the file back to the original directory
        shutil.move(source_path, csv_dir)


# Function to apply color coding based on keywords in the output
def apply_colored_output(text, output_box):
    # Define color tags for different risk levels
    output_box.tag_config("passed", foreground="gray")
    output_box.tag_config("low", foreground="cyan")
    output_box.tag_config("medium", foreground="yellow")
    output_box.tag_config("high", foreground="red")

    # Split the text into lines to check for keywords
    lines = text.split("\n")
    for line in lines:
        if "Passed" in line:
            output_box.insert(ctk.END, line + "\n", "passed")
        elif "Low" in line:
            output_box.insert(ctk.END, line + "\n", "low")
        elif "Medium" in line:
            output_box.insert(ctk.END, line + "\n", "medium")
        elif "High" in line:
            output_box.insert(ctk.END, line + "\n", "high")
        else:
            output_box.insert(ctk.END, line + "\n")  # Default text with no tag


# Create the GUI
def run_gui():
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")

    app = ctk.CTk()
    app.title("PowerShell Module Importer")
    app.geometry("600x400")  # Set a larger default size

    # Create a textbox for output
    output_box = ctk.CTkTextbox(app, wrap='word', width=580, height=300)
    output_box.pack(fill='both', expand=True, padx=20, pady=20)  # Make the textbox fill the window and resize

    # Create buttons to load with or without filtered CSV files
    import_button = ctk.CTkButton(app, text="Import and Invoke DefendeRs Module",
                                  command=lambda: import_and_invoke(output_box, only_filtered_csv=False))
    import_button.pack(pady=5)

    filtered_button = ctk.CTkButton(app, text="Load Only Filtered CSV Files",
                                    command=lambda: import_and_invoke(output_box, only_filtered_csv=True))
    filtered_button.pack(pady=5)

    filtered_button = ctk.CTkButton(app, text="Filter_CSV_finding_list_dod_microsoft_windows_10",
                                    command=lambda: filter_csv(1))
    filtered_button.pack(pady=5)

    # Create an option menu with different options
    options = ["CSV A", "CSV B"]
    selected_option = ctk.StringVar(value=options[0])  # Default selection

    def on_option_change():
        # Get the currently selected option
        value = selected_option.get()
        if value == "Filter_CSV_finding_list_dod_microsoft_windows_10":
            filter_csv(1)  # Call filter_csv(1) for CSV A
        elif value == "CSV B":
            filter_csv(2)  # Call filter_csv(2) for CSV B

    # Create the CTkOptionMenu
    option_menu = ctk.CTkOptionMenu(app, variable=selected_option, values=options)
    option_menu.pack(pady=5)

    # Create a button to trigger the action
    submit_button = ctk.CTkButton(app, text="Submit", command=on_option_change)
    submit_button.pack(pady=5)

    app.mainloop()


if __name__ == "__main__":
    if not is_admin():
        run_as_admin()  # Rerun as admin
    else:
        run_gui()
