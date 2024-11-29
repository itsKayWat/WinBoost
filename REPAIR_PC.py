import os
import subprocess
import ctypes
import sys
import winreg
import shutil
import tempfile
import psutil
import json
import requests
from datetime import datetime
import win32com.client
import win32api
import win32con
import traceback
import logging
import time

# Set up logging
log_file = os.path.join(os.path.expanduser('~'), 'Desktop', 'repair_tool_log.txt')
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_file, mode='w'),
        logging.StreamHandler()
    ]
)

def pause_for_user():
    """Pause and wait for user input."""
    print("\nPress Enter to continue or 'Q' to quit...")
    response = input().lower()
    if response == 'q':
        sys.exit()

def show_error_and_pause(error):
    """Display error and wait for user acknowledgment."""
    print("\n" + "="*50)
    print("ERROR OCCURRED:")
    print(str(error))
    print("\nFull error details:")
    print(traceback.format_exc())
    print("="*50)
    input("\nPress Enter to continue...")

def is_admin():
    """Check if the script is running with administrative privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """Attempt to relaunch the script with administrative privileges if not already."""
    if not is_admin():
        try:
            print("Requesting administrator privileges...")
            ctypes.windll.shell32.ShellExecuteW(
                None,
                "runas",
                sys.executable,
                f'"{os.path.abspath(__file__)}"',
                None,
                1
            )
            sys.exit()
        except Exception as e:
            show_error_and_pause(f"Failed to get admin rights: {e}")
            sys.exit(1)

def create_restore_point():
    """Create a System Restore Point."""
    try:
        print("Creating System Restore Point...")
        logging.info("Creating System Restore Point...")
        
        # Enable System Restore if disabled
        subprocess.run(
            'powershell.exe Enable-ComputerRestore -Drive "C:"',
            shell=True,
            check=True
        )
        
        # Create the restore point
        cmd = (
            'powershell.exe Checkpoint-Computer '
            '-Description "Before System Optimization" '
            '-RestorePointType "MODIFY_SETTINGS"'
        )
        subprocess.run(cmd, shell=True, check=True)
        
        print("Restore point created successfully!")
        logging.info("System Restore Point created successfully.")
    except subprocess.CalledProcessError as e:
        logging.error(f"Failed to create restore point: {e}")
        print(f"Failed to create restore point: {e}")
        input("Press Enter to continue anyway, or CTRL+C to exit...")
    except Exception as e:
        logging.error(f"Unexpected error creating restore point: {e}")
        print(f"Unexpected error creating restore point: {e}")
        input("Press Enter to continue anyway, or CTRL+C to exit...")

def create_batch_file():
    """Create a batch file and a desktop shortcut for easy script execution."""
    try:
        # Get the current script's full path
        script_path = os.path.abspath(__file__)
        batch_path = os.path.join(os.path.dirname(script_path), 'Run_Repair_Tool.bat')
        
        batch_contents = f'''@echo off
echo Starting System Repair Tool...
echo.
python "{script_path}"
if errorlevel 1 (
    echo.
    echo An error occurred! Check the log file on your Desktop.
    pause
) else (
    echo.
    echo Script completed successfully!
    pause
)'''
        
        with open(batch_path, 'w') as f:
            f.write(batch_contents)
        
        # Make batch file executable (not necessary on Windows, but included for completeness)
        try:
            os.chmod(batch_path, 0o755)
        except:
            pass  # Windows doesn't use executable permissions the same way as Unix
        
        print(f"\nCreated batch file at: {batch_path}")
        logging.info(f"Batch file created at: {batch_path}")
        
        # Create desktop shortcut
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        shortcut_path = os.path.join(desktop_path, 'System Repair Tool.bat')
        
        # Copy batch file to desktop
        shutil.copy2(batch_path, shortcut_path)
        print(f"Created desktop shortcut: {shortcut_path}")
        logging.info(f"Desktop shortcut created at: {shortcut_path}")
        
    except Exception as e:
        logging.error(f"Failed to create batch file: {e}")
        print(f"Error creating batch file: {e}")

def install_requirements():
    """Install required Python packages."""
    try:
        print("\nChecking and installing required packages...")
        logging.info("Starting package installation.")
        
        requirements = ['psutil', 'pywin32', 'requests']
        
        # Upgrade pip first
        try:
            print("Upgrading pip...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "pip"])
            logging.info("Pip upgraded successfully.")
        except Exception as e:
            logging.warning(f"Pip upgrade failed: {e}")
            print(f"Pip upgrade failed: {e}. Continuing...")
        
        # Install each package
        for package in requirements:
            try:
                print(f"\nInstalling {package}...")
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])
                print(f"{package} installed successfully!")
                logging.info(f"{package} installed successfully.")
            except Exception as e:
                logging.error(f"Failed to install {package}: {e}")
                print(f"\nError installing {package}. Trying alternative method...")
                try:
                    subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", package])
                    print(f"{package} installed successfully with --user flag!")
                    logging.info(f"{package} installed successfully with --user flag.")
                except Exception as e2:
                    logging.error(f"Alternative installation failed for {package}: {e2}")
                    raise Exception(f"Could not install {package}")
        
        print("\nAll requirements installed successfully!")
        logging.info("Package installation completed successfully.")
        time.sleep(2)
        return True
        
    except Exception as e:
        logging.error(f"Package installation failed: {e}")
        print("\n" + "="*50)
        print("ERROR: Package installation failed!")
        print(f"Error details: {e}")
        print("\nPlease try running these commands manually:")
        print("pip install --upgrade pip")
        print("pip install --user psutil pywin32 requests")
        print("="*50)
        input("\nPress Enter to exit...")
        return False

def safe_import():
    """Safely import required modules after ensuring they are installed."""
    try:
        logging.info("Attempting to import required modules...")
        import ctypes
        import winreg
        import shutil
        import tempfile
        import psutil
        import json
        import requests
        import win32com.client
        import win32api
        import win32con
        import traceback
        logging.info("All modules imported successfully.")
        return True
    except ImportError as e:
        logging.error(f"Import error: {e}")
        return False

def check_python_version():
    """Check if Python version meets the minimum requirement."""
    try:
        logging.info(f"Checking Python version: {sys.version}")
        if sys.version_info < (3, 6):
            logging.error("Incompatible Python version.")
            print("This script requires Python 3.6 or newer.")
            input("Press Enter to exit...")
            return False
        return True
    except Exception as e:
        logging.error(f"Error checking Python version: {e}")
        return False

# Optimization Functions

def disable_startup_programs():
    """Disable common unnecessary startup programs."""
    try:
        logging.info("Disabling startup programs.")
        print("Disabling startup programs...")
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        
        # Backup current startup items
        backup_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'startup_backup.txt')
        with open(backup_path, 'w') as f:
            i = 0
            while True:
                try:
                    name, value, _ = winreg.EnumValue(key, i)
                    f.write(f"{name}: {value}\n")
                    i += 1
                except WindowsError:
                    break
        logging.info(f"Startup programs backed up to {backup_path}")
        print(f"Startup programs backed up to {backup_path}")
        
        # Clear startup items
        i = 0
        while True:
            try:
                name, _, _ = winreg.EnumValue(key, 0)
                winreg.DeleteValue(key, name)
                logging.info(f"Deleted startup item: {name}")
                print(f"Deleted startup item: {name}")
            except WindowsError:
                break
        winreg.CloseKey(key)
        print("Startup programs disabled.")
        logging.info("Startup programs disabled.")
    except Exception as e:
        logging.error(f"Error managing startup programs: {e}")
        print(f"Error managing startup programs: {e}")

def clear_temp_files():
    """Clear temporary files from Windows temp folders."""
    temp_paths = [tempfile.gettempdir(), 
                 os.path.join(os.environ.get('WINDIR', 'C:\\Windows'), 'Temp')]
    
    for temp_path in temp_paths:
        try:
            logging.info(f"Clearing temporary files in {temp_path}")
            print(f"Clearing temporary files in {temp_path}...")
            for item in os.listdir(temp_path):
                item_path = os.path.join(temp_path, item)
                try:
                    if os.path.isfile(item_path):
                        os.unlink(item_path)
                    elif os.path.isdir(item_path):
                        shutil.rmtree(item_path)
                except Exception as e:
                    logging.warning(f"Could not remove {item_path}: {e}")
        except Exception as e:
            logging.error(f"Error accessing {temp_path}: {e}")
            print(f"Error accessing {temp_path}: {e}")

def run_system_commands():
    """Run various Windows system maintenance commands."""
    commands = [
        'sfc /scannow',  # System File Checker
        'DISM /Online /Cleanup-Image /RestoreHealth',  # Windows Image repair
        'chkdsk /f /r C:',  # Check disk with repair
        'ipconfig /flushdns',  # Flush DNS
        'netsh winsock reset',  # Reset Winsock
        'netsh int ip reset',  # Reset TCP/IP stack
        'defrag C: /U /V'  # Defragment C drive
    ]
    
    for cmd in commands:
        try:
            logging.info(f"Executing system command: {cmd}")
            print(f"Executing: {cmd}")
            subprocess.run(cmd, shell=True, check=True)
        except subprocess.CalledProcessError as e:
            logging.error(f"Error running {cmd}: {e}")
            print(f"Error running {cmd}: {e}")

def clear_browser_data():
    """Clear data from multiple browsers."""
    # Paths to browser data
    browsers = {
        'Chrome': os.path.expanduser('~\\AppData\\Local\\Google\\Chrome\\User Data\\Default'),
        'Firefox': os.path.expanduser('~\\AppData\\Local\\Mozilla\\Firefox\\Profiles'),
        'Edge': os.path.expanduser('~\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default')
    }
    
    for browser, path in browsers.items():
        try:
            if os.path.exists(path):
                logging.info(f"Clearing {browser} data at {path}")
                print(f"Clearing {browser} data...")
                if browser == 'Firefox':
                    # Firefox profiles can have multiple folders
                    for profile in os.listdir(path):
                        profile_path = os.path.join(path, profile)
                        shutil.rmtree(profile_path, ignore_errors=True)
                else:
                    shutil.rmtree(path, ignore_errors=True)
                print(f"{browser} data cleared.")
                logging.info(f"{browser} data cleared.")
        except Exception as e:
            logging.warning(f"Error clearing {browser} data: {e}")
            print(f"Error clearing {browser} data: {e}")

def optimize_services():
    """Optimize Windows services."""
    try:
        services_to_disable = [
            'DiagTrack',  # Connected User Experiences and Telemetry
            'dmwappushservice',  # WAP Push Message Routing Service
            'SysMain',  # Superfetch
            'WSearch'  # Windows Search
        ]
        
        for service in services_to_disable:
            try:
                subprocess.run(f'sc config {service} start= disabled', shell=True, check=True)
                subprocess.run(f'sc stop {service}', shell=True, check=True)
                logging.info(f"Service {service} disabled and stopped.")
                print(f"Service {service} disabled and stopped.")
            except subprocess.CalledProcessError as e:
                logging.warning(f"Could not disable service {service}: {e}")
                print(f"Could not disable service {service}: {e}")
    except Exception as e:
        logging.error(f"Error optimizing services: {e}")
        print(f"Error optimizing services: {e}")

def clean_registry():
    """Clean and optimize Windows Registry."""
    reg_commands = [
        'reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RunMRU" /va /f',
        'reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\TypedPaths" /va /f',
        'reg delete "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\RecentDocs" /va /f'
    ]
    
    for cmd in reg_commands:
        try:
            subprocess.run(cmd, shell=True, check=True)
            logging.info(f"Executed registry command: {cmd}")
            print(f"Executed registry command: {cmd}")
        except subprocess.CalledProcessError as e:
            logging.warning(f"Error executing registry command '{cmd}': {e}")
            print(f"Error executing registry command '{cmd}': {e}")

def optimize_performance_settings():
    """Optimize Windows performance settings."""
    try:
        # Disable visual effects
        subprocess.run(
            'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" '
            '/v VisualFXSetting /t REG_DWORD /d 2 /f',
            shell=True,
            check=True
        )
        logging.info("Disabled visual effects.")
        print("Disabled visual effects.")
        
        # Set power plan to high performance
        subprocess.run(
            'powercfg /setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c', 
            shell=True, 
            check=True
        )
        logging.info("Set power plan to High Performance.")
        print("Set power plan to High Performance.")
        
        # Disable hibernation
        subprocess.run('powercfg /hibernate off', shell=True, check=True)
        logging.info("Disabled hibernation.")
        print("Disabled hibernation.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error optimizing performance settings: {e}")
        print(f"Error optimizing performance settings: {e}")

def clean_system_drives():
    """Deep clean system drives."""
    try:
        print("Running Disk Cleanup...")
        logging.info("Running Disk Cleanup.")
        subprocess.run('cleanmgr /sagerun:1', shell=True, check=True)
        
        # Clear Windows update cache
        logging.info("Clearing Windows Update cache.")
        subprocess.run('net stop wuauserv', shell=True, check=True)
        shutil.rmtree('C:\\Windows\\SoftwareDistribution', ignore_errors=True)
        subprocess.run('net start wuauserv', shell=True, check=True)
        print("Windows Update cache cleared.")
        
        # Clear Event Logs
        logging.info("Clearing Event Logs.")
        subprocess.run(
            'powershell.exe "Get-EventLog -LogName * | Clear-EventLog"',
            shell=True,
            check=True
        )
        print("Event Logs cleared.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error cleaning system drives: {e}")
        print(f"Error cleaning system drives: {e}")
    except Exception as e:
        logging.error(f"Unexpected error cleaning system drives: {e}")
        print(f"Unexpected error cleaning system drives: {e}")

def optimize_network():
    """Optimize network settings."""
    network_commands = [
        'netsh int tcp set global autotuninglevel=normal',
        'netsh int tcp set global chimney=enabled',
        'netsh int tcp set global dca=enabled',
        'netsh int tcp set global netdma=enabled',
        'netsh int tcp set global ecncapability=enabled',
        'netsh int tcp set global timestamps=disabled',
        'netsh int tcp set heuristics disabled',
        'netsh int tcp set global rss=enabled'
    ]
    
    for cmd in network_commands:
        try:
            subprocess.run(cmd, shell=True, check=True)
            logging.info(f"Executed network command: {cmd}")
            print(f"Executed network command: {cmd}")
        except subprocess.CalledProcessError as e:
            logging.warning(f"Error executing network command '{cmd}': {e}")
            print(f"Error executing network command '{cmd}': {e}")

def optimize_cpu_power():
    """Optimize CPU power settings."""
    try:
        # Set CPU performance to maximum
        subprocess.run(
            'powercfg -setacvalueindex scheme_current sub_processor PROCTHROTTLEMAX 100', 
            shell=True, 
            check=True
        )
        subprocess.run(
            'powercfg -setacvalueindex scheme_current sub_processor PROCTHROTTLEMIN 100', 
            shell=True, 
            check=True
        )
        subprocess.run('powercfg -setactive scheme_current', shell=True, check=True)
        logging.info("Optimized CPU power settings.")
        print("Optimized CPU power settings.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error optimizing CPU power settings: {e}")
        print(f"Error optimizing CPU power settings: {e}")

def clear_print_spooler():
    """Clear stuck print jobs and reset spooler."""
    try:
        subprocess.run('net stop spooler', shell=True, check=True)
        spooler_path = 'C:\\Windows\\System32\\spool\\PRINTERS'
        if os.path.exists(spooler_path):
            shutil.rmtree(spooler_path)
            os.mkdir(spooler_path)
            logging.info("Cleared print spooler.")
            print("Cleared print spooler.")
        subprocess.run('net start spooler', shell=True, check=True)
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error resetting print spooler: {e}")
        print(f"Error resetting print spooler: {e}")
    except Exception as e:
        logging.error(f"Unexpected error resetting print spooler: {e}")
        print(f"Unexpected error resetting print spooler: {e}")

def optimize_gaming_settings():
    """Optimize Windows for gaming performance."""
    try:
        # Disable Game DVR and Game Bar
        subprocess.run(
            'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" '
            '/v AllowGameDVR /t REG_DWORD /d 0 /f',
            shell=True,
            check=True
        )
        subprocess.run(
            'reg add "HKCU\\System\\GameConfigStore" '
            '/v GameDVR_Enabled /t REG_DWORD /d 0 /f',
            shell=True,
            check=True
        )
        
        # Optimize for performance
        subprocess.run(
            'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" '
            '/v SystemResponsiveness /t REG_DWORD /d 0 /f',
            shell=True,
            check=True
        )
        subprocess.run(
            'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" '
            '/v "GPU Priority" /t REG_DWORD /d 8 /f',
            shell=True,
            check=True
        )
        subprocess.run(
            'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" '
            '/v "Priority" /t REG_DWORD /d 6 /f',
            shell=True,
            check=True
        )
        
        logging.info("Optimized gaming settings.")
        print("Optimized gaming settings.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error optimizing gaming settings: {e}")
        print(f"Error optimizing gaming settings: {e}")

def clear_windows_defender_history():
    """Clear Windows Defender scan history and quarantine."""
    try:
        subprocess.run('MpCmdRun.exe -RemoveDefinitions -All', shell=True, check=True)
        subprocess.run('MpCmdRun.exe -DeleteAllRestorePoints', shell=True, check=True)
        logging.info("Cleared Windows Defender history and quarantine.")
        print("Cleared Windows Defender history and quarantine.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error clearing Windows Defender history: {e}")
        print(f"Error clearing Windows Defender history: {e}")
    except Exception as e:
        logging.error(f"Unexpected error clearing Windows Defender history: {e}")
        print(f"Unexpected error clearing Windows Defender history: {e}")

def optimize_ssd():
    """Optimize SSD if present."""
    try:
        # Check if system drive is SSD
        wmi = win32com.client.GetObject("winmgmts:")
        physical_disks = wmi.InstancesOf("Win32_PhysicalMedia")
        is_ssd = False
        for disk in physical_disks:
            # This is a placeholder check; proper SSD detection requires more detailed querying
            if disk.Tag == "Disk #0":
                is_ssd = True  # Placeholder: Implement actual SSD detection
                break
        
        if is_ssd:
            # Disable defragmentation for SSDs
            subprocess.run('fsutil behavior set DisableLastAccess 1', shell=True, check=True)
            subprocess.run('fsutil behavior set EncryptPagingFile 0', shell=True, check=True)
            
            # Enable TRIM
            subprocess.run('fsutil behavior set DisableDeleteNotify 0', shell=True, check=True)
            
            logging.info("Optimized SSD settings.")
            print("Optimized SSD settings.")
        else:
            logging.info("No SSD detected. Skipping SSD optimizations.")
            print("No SSD detected. Skipping SSD optimizations.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error optimizing SSD: {e}")
        print(f"Error optimizing SSD: {e}")
    except Exception as e:
        logging.error(f"Unexpected error optimizing SSD: {e}")
        print(f"Unexpected error optimizing SSD: {e}")

def repair_windows_updates():
    """Repair Windows Update issues."""
    update_commands = [
        'net stop wuauserv',
        'net stop cryptSvc',
        'net stop bits',
        'net stop msiserver',
        'rename C:\\Windows\\SoftwareDistribution SoftwareDistribution.old',
        'rename C:\\Windows\\System32\\catroot2 catroot2.old',
        'net start wuauserv',
        'net start cryptSvc',
        'net start bits',
        'net start msiserver'
    ]
    
    for cmd in update_commands:
        try:
            subprocess.run(cmd, shell=True, check=True)
            logging.info(f"Executed update command: {cmd}")
            print(f"Executed update command: {cmd}")
        except subprocess.CalledProcessError as e:
            logging.warning(f"Error executing update command '{cmd}': {e}")
            print(f"Error executing update command '{cmd}': {e}")

def clear_font_cache():
    """Clear Windows Font Cache."""
    try:
        subprocess.run('net stop FontCache', shell=True, check=True)
        subprocess.run('net stop FontCache3.0.0.0', shell=True, check=True)
        
        font_cache_path = 'C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\FontCache'
        if os.path.exists(font_cache_path):
            os.remove(font_cache_path)
        
        subprocess.run('net start FontCache', shell=True, check=True)
        subprocess.run('net start FontCache3.0.0.0', shell=True, check=True)
        
        logging.info("Cleared font cache.")
        print("Cleared font cache.")
    except subprocess.CalledProcessError as e:
        logging.warning(f"Error clearing font cache: {e}")
        print(f"Error clearing font cache: {e}")
    except Exception as e:
        logging.error(f"Unexpected error clearing font cache: {e}")
        print(f"Unexpected error clearing font cache: {e}")

def main_optimization_tasks():
    """Run all optimization tasks."""
    try:
        print("\n1. Creating System Restore Point...")
        create_restore_point()
        pause_for_user()
        
        print("\n2. Disabling startup programs...")
        disable_startup_programs()
        pause_for_user()
        
        print("\n3. Clearing temporary files...")
        clear_temp_files()
        pause_for_user()
        
        print("\n4. Clearing browser data...")
        clear_browser_data()
        pause_for_user()
        
        print("\n5. Optimizing Windows services...")
        optimize_services()
        pause_for_user()
        
        print("\n6. Cleaning registry...")
        clean_registry()
        pause_for_user()
        
        print("\n7. Optimizing performance settings...")
        optimize_performance_settings()
        pause_for_user()
        
        print("\n8. Cleaning system drives...")
        clean_system_drives()
        pause_for_user()
        
        print("\n9. Optimizing network settings...")
        optimize_network()
        pause_for_user()
        
        print("\n10. Optimizing CPU power settings...")
        optimize_cpu_power()
        pause_for_user()
        
        print("\n11. Clearing print spooler...")
        clear_print_spooler()
        pause_for_user()
        
        print("\n12. Optimizing gaming settings...")
        optimize_gaming_settings()
        pause_for_user()
        
        print("\n13. Clearing Windows Defender history...")
        clear_windows_defender_history()
        pause_for_user()
        
        print("\n14. Optimizing SSD settings...")
        optimize_ssd()
        pause_for_user()
        
        print("\n15. Repairing Windows Updates...")
        repair_windows_updates()
        pause_for_user()
        
        print("\n16. Clearing font cache...")
        clear_font_cache()
        pause_for_user()
        
        print("\nAll optimization tasks completed successfully!")
        logging.info("All optimization tasks completed successfully.")
        
        print("\nSystem needs to restart to apply changes.")
        logging.info("System restart initiated.")
        input("Press Enter to restart your computer...")
        
        # Create restore point before restarting
        subprocess.run(
            'wmic.exe /Namespace:\\\\root\\default Path SystemRestore Call CreateRestorePoint "Before System Restart", 100, 7', 
            shell=True, 
            check=True
        )
        
        # Restart computer with a 10-second delay
        subprocess.run('shutdown /r /t 10', shell=True, check=True)
        print("System will restart in 10 seconds...")
        logging.info("System restart scheduled in 10 seconds.")
        time.sleep(10)
        
    except Exception as e:
        logging.error(f"Error during optimization tasks: {e}")
        print(f"An error occurred during optimization tasks: {e}")
        input("Press Enter to exit...")

def main():
    """Main function to execute optimization tasks."""
    main_optimization_tasks()

if __name__ == "__main__":
    try:
        print("\nSystem Repair Tool - Initialization")
        print("="*40)
        logging.info("Script initialization started.")
        
        # Create batch file first
        create_batch_file()
        
        # Check Python version
        if not check_python_version():
            sys.exit(1)
        
        # Install required packages
        if not install_requirements():
            sys.exit(1)
        
        # Try importing required modules
        if not safe_import():
            print("\nFailed to import required modules. Please try running the script again.")
            input("Press Enter to exit...")
            sys.exit(1)
        
        print("\nInitialization successful! Starting main program...")
        logging.info("Initialization successful. Starting main program.")
        time.sleep(2)
        
        # Execute main optimization tasks
        main()
        
    except Exception as e:
        logging.critical(f"Critical error during initialization: {e}", exc_info=True)
        print("\n" + "="*50)
        print("A critical error occurred during initialization!")
        print(f"Error details: {e}")
        print(f"\nA detailed log file has been created at: {log_file}")
        print("="*50)
        input("\nPress Enter to exit...")
        sys.exit(1)
    
    finally:
        logging.info("Script execution completed.")
        input("\nPress Enter to exit...")