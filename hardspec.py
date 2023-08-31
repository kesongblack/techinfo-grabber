# DISCLAIMER: 
# - Code obtained thru interactions with ChatGPT. 
# - I only compile them into a single file for this purpose only.
# - This program only works on Windows.

from datetime import datetime
from timeout_decorator import timeout
import cpuinfo
import humanize
import multiprocessing
import os
import psutil
import pyperclip
import re
import speedtest
import subprocess
import time
import winreg
import win32com.client
import wmi

def loading_screen():
    # Get the current date and time
    current_datetime = datetime.now()

    # Format the datetime as a readable string
    curr_time = current_datetime.strftime("%I:%M %p")
    
    animation = ["Loading program, please wait.   ",
                "Loading program, please wait..  ",
                "Loading program, please wait... ",
                "Loading program, please wait...."]
    for frame in animation:
        print(frame, end='\r')
        time.sleep(0.2)  # Adjust the delay as needed
    print('\n')
    print("Technical Information Grabber Version 1.0")
    print("by KesongBlack\n")
    print("PROGRAM RUNNING AT EXACTLY", curr_time)
    print("If program exceeds five minutes and doesn't show anything else")
    print("Please close this program and indicate 'PROGRAM ERROR' on your Excel sheet.\n")
        
# Shows current progress of scanning
#region output_progress
def output_progress(component):
    print("\r--- Checking " + component + " --------------", end='')
#endregion output_progress

# The magic starts here. Char
#region MAIN PROGRAM
def main_program():
    # COMPUTER INFORMATION - related functions
    # Determine information from computer as well as OS license status
    #region COMPUTER INFORMATION
    # Determine if desktop or laptop by checking existence of battery
    #region is_laptop()
    def is_laptop():
        output_progress("Device Type")
        battery = psutil.sensors_battery()

        if battery is None:
            return "DESKTOP"
        else:
            return "LAPTOP"
    #endregion    
        
    # Determine what version of Windows it is and its license status
    #region windows_info_and_license()

    # Get Windows Information
    #region get_windows_info()
    def get_systeminfo():
        output_progress("System Information")
        try:
            result = subprocess.run(['systeminfo'], capture_output=True, text=True, shell=True)
            output = result.stdout
            return output
        except FileNotFoundError:
            return "systeminfo command not found"

    def get_systeminfo_item(item, systeminfo_output):
        item = r'' + item + ':\s+([^\n]+)'
        match = re.search(item, systeminfo_output)
        if match:
            return match.group(1).strip()
        else:
            return "OS Version not found"
    #endregion get_windows_info()

    # Get Windows License Status and Partial Product Key
    #region windows_license()
    def get_system32_path():
        output_progress("Windows Version")
        system_root = os.environ.get('SystemRoot', 'C:\\Windows')
        return os.path.join(system_root, 'System32')

    def get_license_info():
        try:
            output_progress("License Information")
            # Determine the appropriate path to slmgr.vbs script
            script_path = os.path.join(get_system32_path(), 'slmgr.vbs')

            # Run slmgr.vbs script to retrieve license information
            result = subprocess.run(['cscript', '//Nologo', script_path, '-dli'], capture_output=True, text=True, stdin=subprocess.PIPE)

            if result.returncode == 0:
                output = result.stdout

                # Extract license status and product key
                license_status = re.search(r'License Status:\s*(.*)', output).group(1)
                product_key = re.search(r'Partial Product Key:\s*(\w{5})', output).group(1)
                
                if product_key == '3V66T':
                    license_status = "Cracked"
            else:
                license_status = "Undetermined."
                product_key = "N/A"
            return license_status, product_key
        except Exception as e:
            return None, None
    #endregion windows_license()
    #endregion windows_info_and_license()
    #endregion COMPUTER INFORMATION

    # Get current internet speed of unit
    #region get_internet_speed
    def get_internet_speed():
        try:
            output_progress("Internet Speed")
            st = speedtest.Speedtest()

            # Get the best server
            st.get_best_server()

            # Perform download and upload speed tests
            download_speed = st.download() / 1024 / 1024  # Convert to Mbps
            upload_speed = st.upload() / 1024 / 1024  # Convert to Mbps
            
            return {
                "Download Speed": f"{download_speed:.2f} Mbps",
                "Upload Speed": f"{upload_speed:.2f} Mbps"
            }
        except Exception as e:
            print("Encountered problems in connecting the internet.")
            return {
                "Download Speed": "---",
                "Upload Speed": "---"
            }
    #endregion SPEEDTEST    


    # INSTALLED SOFTWARE - related functions
    # Determine essential softwares installed in the system
    #region INSTALLED SOFTWARE

    # Get Office version and its license
    #region get_office_version()

    # Retrieve version using registry
    def has_subkeys_recursive(registry_key, subkey_name):
        try:
            with winreg.OpenKey(registry_key, subkey_name):
                return True
        except FileNotFoundError:
            return False

    def check_subkeys(registry_key_path):
        root_key, *subkeys = registry_key_path.split("\\")
        registry_hive = {
            "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
            "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
            "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
            "HKEY_USERS": winreg.HKEY_USERS,
            "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
        }

        if root_key in registry_hive:
            current_key = registry_hive[root_key]
            for subkey_name in subkeys:
                if not has_subkeys_recursive(current_key, subkey_name):
                    return False
                current_key = winreg.OpenKey(current_key, subkey_name)
            return True
        else:
            return False

    # Get MS Office year using version number
    def office_list(version):
        match version:
            case "12.0":
                return "2007"
            case "14.0":
                return "2010"
            case "15.0":
                return "2013"
            case "16.0":
                # Since Office 2016^ has the same version, use another method
                registry_key_path = r"HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Licensing"
                if check_subkeys(registry_key_path):
                    return "2019 + "
                else:
                    return "2016"
            case _:
                return "Unknown"

    # Grab version number using application
    def get_office_version():
        try:
            output_progress("Installed Applications")
            obj = win32com.client.Dispatch("Word.Application")
            version = office_list(obj.Application.Version)
            return version
        except Exception:
            return "No MS Office installed."            
    #endregion

    # Get all antivirus installed
    #region get_antivirus_software()
    def get_antivirus_software():
        w = wmi.WMI(namespace=r'root\SecurityCenter2')
        antivirus_list = []

        for antivirus in w.AntivirusProduct():
            antivirus_list.append(antivirus.displayName)

        return antivirus_list

    def display_antivirus_software():
        installed_antivirus = get_antivirus_software()
        installed_antivirus_str = ""
        if installed_antivirus:
            print("Installed Antivirus Software:")
            for antivirus in installed_antivirus:
                installed_antivirus_str += antivirus + "\n"
                print(f"   - {antivirus}")
        else:
            print("No antivirus software found.")
            
        return installed_antivirus_str
    #endregion get_antivirus_software()
    #endregion INSTALLED SOFTWARE

    # PROCESSOR - related functions
    #region PROCESSOR
    def get_processor_info():
        try: 
            output_progress("Processor Information")
            info = cpuinfo.get_cpu_info()
            
            return {
                "CPU Model": info["brand_raw"],
                "CPU Frequency": psutil.cpu_freq().max,
                "CPU Cores": psutil.cpu_count(logical=False),
                "CPU Threads": psutil.cpu_count(logical=True),
            }
        except Exception as e:
            print("No processor information obtained.")
            return None
    #endregion PROCESSOR 

    # MEMORY - related functions
    #region MEMORY 
    def get_memory_type_from_wmic():
        try:
            output_progress("Memory Information")
            cmd = "wmic memorychip get MemoryType"
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True)

            memory_types = {
                20: "DDR",
                21: "DDR2",
                24: "DDR3",
                26: "DDR4",
                29: "DDR5"
            }

            for line in result.stdout.splitlines():
                if line.isdigit():
                    memory_type = int(line)
                    if memory_type in memory_types:
                        return memory_types[memory_type]

            return None
        except Exception as e:
            return None

    def get_memory_speed():
        try:
            c = wmi.WMI()
            speeds = set()

            for mem_module in c.Win32_PhysicalMemory():
                speeds.add(mem_module.Speed)
            
            if (min(speeds) == None):
                return "---"
            else:
                return min(speeds)
        except Exception as e:
            return "Unknown"

    def get_ddr_generation(speed):
        ddr_generations = {
            400: "DDR",
            533: "DDR2",
            667: "DDR2",
            800: "DDR2",
            1066: "DDR3",
            1333: "DDR3",
            1600: "DDR3",
            1866: "DDR4",
            2133: "DDR4",
            2400: "DDR4",
            2666: "DDR4",
            2933: "DDR4",
            3200: "DDR4",
            3600: "DDR4",
            4000: "DDR4",
        }

        for speed_limit, generation in ddr_generations.items():
            if speed <= speed_limit:
                return generation

        return "Unknown"

    def get_memory_ddr_generation():
        try: 
            memory_type = get_memory_type_from_wmic()
            if memory_type:
                return memory_type
            else:
                speed = get_memory_speed()
                return get_ddr_generation(speed)
        except Exception as e:
            return "Cannot determine (memory speed missing)."

    def get_physical_memory_size():
        vm = psutil.virtual_memory()
        swap_space = psutil.swap_memory()

        total_physical_memory = vm.total - swap_space.used

        return humanize.naturalsize(total_physical_memory)
    #endregion


    # HARD DISK - related functions
    #region HARD DISK
    def check_powershell_existence():
        try:
            # Run the "powershell" command with the "-version" option to check if PowerShell exists
            subprocess.run(["powershell", "-version"])
            return True
        except FileNotFoundError:
            return False

    
    
    def get_physical_disk_info():
        try: 
            powershell_cmd = 'get-physicaldisk | format-table -Property Size, MediaType, FriendlyName -autosize'
            result = subprocess.run(['powershell', powershell_cmd], capture_output=True, text=True)

            if result.returncode == 0:
                return result.stdout
            else:
                return None
        except Exception as e:
            return None

        if check_powershell_existence():
            print("PowerShell is installed on this computer.")
        else:
            print("PowerShell is NOT installed on this computer.")
        

    def parse_disk_info(disk_info):
        lines = disk_info.strip().split('\n')
        data_start = None

        # Find the position of the header line
        for i, line in enumerate(lines):
            if line.startswith('Size'):
                data_start = i + 1
                break

        if data_start is not None:
            data_lines = lines[data_start:]
            data = []

            for line in data_lines:
                match = re.match(r'\s*(\d+)\s+(\S+)\s+(.+)', line)
                if match:
                    size, media_type, friendly_name = match.groups()
                    data.append((size, media_type, friendly_name))

            size_column = [humanize.naturalsize(item[0]) for item in data]
            media_type_column = [item[1] for item in data]
            friendly_name_column = [item[2] for item in data]

            return size_column, media_type_column, friendly_name_column
        else:
            return [], [], []

    disk_info = get_physical_disk_info()
    disk_list = []

    if disk_info is not None:
        size_column, media_type_column, friendly_name_column = parse_disk_info(disk_info)
        for x in range(len(size_column)):
            disk_list.append({
                "Name": friendly_name_column[x],
                "Type": media_type_column[x],
                "Size": size_column[x]
            })
    else:
        print("Failed to retrieve disk information")
    #endregion
    
    # Print in neat, tabbed order
    #region tab_print()
    def tab_print(column, value):
        print('{:<24} {:<24}'.format(column + ":", value))
    #endregion
    
    # All variable declaration starts here, separated by output_progress()  
    # Note: Antivirus and Disk Info not listed here; check their functions
    #region VARIABLES
    device_type = is_laptop()
    systeminfo_output = get_systeminfo()
    prod_id = get_systeminfo_item("Product ID", systeminfo_output)
    os_version = get_systeminfo_item("OS Name", systeminfo_output)
    archi = get_systeminfo_item("System Type", systeminfo_output)
    # output_progress("License Status")
    license_status, product_key = get_license_info()
    office_version = get_office_version()
    internet_speed = get_internet_speed()
    processor_info = get_processor_info()
    memory_speed = get_memory_speed()
    ddr_generation = get_memory_ddr_generation()
    physical_memory_size = get_physical_memory_size()
    #endregion

    # Content spilling starts here
    #region PRINT ALL INFORMATION
    output_progress('Complete. Computer Information will be shown in 3 seconds.')
    time.sleep(3)
    print("\n\n---------------------------")
    print("## COMPUTER INFORMATION ##")
    print("---------------------------")
    tab_print("Product ID", prod_id)
    tab_print("Computer Type", device_type)
    tab_print("Windows Version", os_version)
    tab_print("System Architecture", archi)
    tab_print("License Status", license_status)
    tab_print("Product Key (last 5 digits)", product_key)

    print("\n## INSTALLED SOFTWARE ##")
    print("---------------------------")
    tab_print("Microsoft Office Version", office_version)
    antivirus = display_antivirus_software() 

    print("\n## INTERNET SPEED ##")
    print("---------------------------")
    for key, value in internet_speed.items():
        tab_print(key, value)
        
    print("\n## PROCESSOR ##")
    print("---------------------------")
    for key, value in processor_info.items():
        tab_print(key, value)

    print("\n## MEMORY ##")
    print("---------------------------")
    tab_print("Memory Size", str(physical_memory_size))
    tab_print("Memory Speed", str(memory_speed) + " MHz")
    tab_print("Memory Type", ddr_generation)

    print("\n## HARD DISK ##")
    print("---------------------------")
    print("Total Number: {}".format(len(disk_list)))
    print("Disk List:") 
    count = 0
    has_ssd = "NO"
    disk_list_for_clipboard = []
    for entry in disk_list:
        count = count + 1
        print('  Disk [{}]'.format(count))
        for key, value in entry.items():
            print('    {:<8} {:<12}'.format(key + ":", value))
            
            # Append code for clipboard
            if (key == "Type" and value == "SSD"):
                has_ssd = "YES"
            disk_list_for_clipboard.append(value)

    #endregion PRINT ALL INFORMATION
    
    # Append data for clipboard
    #region CLIPBOARD DATA
    data = [prod_id, device_type, os_version, archi, license_status, product_key, office_version, antivirus]
    data.extend(internet_speed.values())
    data.extend(list(processor_info.values()))
    data.extend([physical_memory_size, memory_speed, ddr_generation, has_ssd])
    data.extend(disk_list_for_clipboard)
    #endregion CLIPBOARD DATA
    
    return data
#endregion MAIN PROGRAM

def add_to_clipboard(data):
    # Convert all elements in the list to strings and create TSV-formatted data
    tsv_data = "\t".join(['"{}"'.format(str(field).replace('"', '""')) for field in data])

    # Copy TSV-formatted data to clipboard
    pyperclip.copy(tsv_data)

    # Print instructions
    print("\n---------------------------")
    print("Data copied to clipboard. Please do the following:")
    print("- Open the Excel File")
    print("- Click the 'Technical Specs' tab")
    print("- Click the cell next to your computer number")
    print("- Paste the data (or use Ctrl+V)")
    print("---------------------------\n")

if __name__ == '__main__':
    multiprocessing.freeze_support() 
    loading_screen()
    data = main_program()
    print("\n---------------------------")
    print("Info generated successfully.")
    # Display a prompt to the user and wait for input
    add_to_clipboard(data)
    user_input = input("Program ends here. Press Enter to close.")
