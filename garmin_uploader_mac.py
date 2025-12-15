#!/usr/bin/env python3
"""
Garmin Workout Uploader for Mac
A guided app to upload .FIT workout files to your Garmin watch.
"""

import os
import sys
import shutil
import subprocess
import webbrowser
import struct
import re
import threading
import time
import json
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from urllib.request import urlopen
from urllib.error import URLError

try:
    from version import __version__, __app_name__, __github_repo__
except ImportError:
    __version__ = "1.0.0"
    __app_name__ = "Garmin Workout Uploader"
    __github_repo__ = "supergeri/garmin-usb-mac-app"

# Try to import tkinterdnd2 for drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# Try to import fitparse for FIT file parsing
try:
    from fitparse import FitFile
    FITPARSE_AVAILABLE = True
except ImportError:
    FITPARSE_AVAILABLE = False

# Try to import amakaflow-fitfiletool for workout repair and FIT parsing
try:
    from amakaflow_fitfiletool import (
        GarminExerciseLookup, build_fit_workout, get_preview_steps, get_fit_metadata,
        parse_fit_file as fitfiletool_parse_fit_file,
        validate_fit_file as fitfiletool_validate_fit_file,
        get_sport_display, get_sport_color, format_duration, format_distance,
        SPORT_COLORS, SPORT_DISPLAY_NAMES, SUB_SPORT_DISPLAY_NAMES, EXERCISE_CATEGORY_NAMES
    )
    FITFILETOOL_AVAILABLE = True
except ImportError:
    FITFILETOOL_AVAILABLE = False
    # Fallback definitions
    SPORT_COLORS = {'training': '#8b5cf6', 'fitness_equipment': '#06b6d4'}
    EXERCISE_CATEGORY_NAMES = {}
    def get_sport_display(sport, sub_sport=None): return sport.replace('_', ' ').title() if sport else 'Workout'
    def get_sport_color(sport, sub_sport=None): return SPORT_COLORS.get(sport, '#6b7280')
    def fitfiletool_parse_fit_file(filepath): return None
    def fitfiletool_validate_fit_file(filepath): return {'valid': True, 'issues': [], 'warnings': []}

# Garmin USB Vendor ID
GARMIN_VENDOR_ID = "0x091e"


# Garmin exercise name mapping (from FIT SDK)
EXERCISE_NAMES = {
    # Strength exercises
    0: "Bench Press", 1: "Calf Raise", 2: "Cardio", 3: "Carry", 4: "Chop",
    5: "Core", 6: "Crunch", 7: "Curl", 8: "Deadlift", 9: "Flye",
    10: "Hip Raise", 11: "Hip Stability", 12: "Hip Swing", 13: "Hyperextension",
    14: "Lateral Raise", 15: "Leg Curl", 16: "Leg Raise", 17: "Lunge",
    18: "Olympic Lift", 19: "Plank", 20: "Plyo", 21: "Pull Up", 22: "Push Up",
    23: "Row", 24: "Shoulder Press", 25: "Shoulder Stability", 26: "Shrug",
    27: "Sit Up", 28: "Squat", 29: "Total Body", 30: "Triceps Extension",
    31: "Warm Up", 32: "Run", 33: "Unknown", 34: "Rest",
    # Cardio
    65534: "Workout", 65535: "Unknown"
}

# Duration type mapping
DURATION_TYPES = {
    0: "time", 1: "distance", 2: "hr_less_than", 3: "hr_greater_than",
    4: "calories", 5: "open", 6: "repeat_until_steps_cmplt",
    7: "repeat_until_time", 8: "repeat_until_distance", 9: "repeat_until_calories",
    10: "repeat_until_hr_less_than", 11: "repeat_until_hr_greater_than",
    12: "repeat_until_power_less_than", 13: "repeat_until_power_greater_than",
    14: "power_less_than", 15: "power_greater_than", 28: "reps"
}


class UpdateChecker:
    """Check for app updates from GitHub releases"""

    @staticmethod
    def _compare_versions(v1, v2):
        """Compare two version strings properly (handles 1.0.10 > 1.0.9)"""
        def parse_version(v):
            return [int(x) for x in v.split('.')]
        try:
            return parse_version(v1) > parse_version(v2)
        except (ValueError, AttributeError):
            return v1 > v2

    @staticmethod
    def check_for_updates():
        """Check if a new version is available on GitHub"""
        import ssl
        try:
            import certifi
            ssl_context = ssl.create_default_context(cafile=certifi.where())
        except ImportError:
            # Fallback if certifi not available (shouldn't happen in bundled app)
            ssl_context = ssl.create_default_context()
        try:
            url = f"https://api.github.com/repos/{__github_repo__}/releases/latest"
            with urlopen(url, timeout=10, context=ssl_context) as response:
                data = json.loads(response.read().decode())
                latest_version = data['tag_name'].lstrip('v')
                download_url = None

                for asset in data.get('assets', []):
                    if asset['name'].endswith('.dmg') or asset['name'].endswith('.pkg'):
                        download_url = asset['browser_download_url']
                        break

                return {
                    'available': UpdateChecker._compare_versions(latest_version, __version__),
                    'version': latest_version,
                    'url': download_url,  # None if no DMG/PKG asset found
                    'release_url': data['html_url'],  # For manual download fallback
                    'notes': data.get('body', '')
                }
        except Exception as e:
            print(f"Update check error: {e}")
            return None

    @staticmethod
    def download_update(url, callback=None):
        """Download the update installer"""
        import ssl
        import tempfile
        try:
            import certifi
            ssl_context = ssl.create_default_context(cafile=certifi.where())
        except ImportError:
            ssl_context = ssl.create_default_context()
        try:
            # Determine filename from URL
            filename = url.split('/')[-1] if '/' in url else 'GarminWorkoutUploader.dmg'
            temp_file = os.path.join(tempfile.gettempdir(), filename)

            with urlopen(url, timeout=60, context=ssl_context) as response:
                total_size = int(response.headers.get('content-length', 0))
                downloaded = 0

                with open(temp_file, 'wb') as f:
                    while True:
                        chunk = response.read(8192)
                        if not chunk:
                            break
                        f.write(chunk)
                        downloaded += len(chunk)
                        if callback and total_size:
                            callback(downloaded / total_size)

            return temp_file
        except Exception as e:
            print(f"Download error: {e}")
            return None


class GarminUploaderMac:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("580x820")
        self.root.resizable(False, False)
        
        # Styling
        self.root.configure(bg='#f5f5f7')
        
        self.style = ttk.Style()
        self.style.configure("Title.TLabel", font=('SF Pro Display', 22, 'bold'), background='#f5f5f7')
        self.style.configure("Subtitle.TLabel", font=('SF Pro Text', 12), background='#f5f5f7', foreground='#666')
        self.style.configure("Step.TLabel", font=('SF Pro Text', 11), background='#fff')
        self.style.configure("StepNum.TLabel", font=('SF Pro Display', 14, 'bold'), background='#007AFF', foreground='white')
        self.style.configure("Big.TButton", font=('SF Pro Text', 13), padding=12)
        self.style.configure("Card.TFrame", background='#fff')
        
        # Paths
        self.home = Path.home()
        self.staging_folder = self.home / "GarminWorkouts"
        self.staging_folder.mkdir(exist_ok=True)
        
        self.selected_files = []
        self.openmtp_installed = self.check_openmtp()
        self.libmtp_installed = self.check_libmtp()
        
        # Track drag state for visual feedback
        self.is_dragging = False
        
        # UI elements initialized later
        self.close_ge_btn = None
        self.refresh_btn = None
        self._monitor_running = True
        self.transfer_btns_frame = None
        self.openmtp_warning_frame = None

        # Track connected device for model-specific adjustments
        self.current_device = None

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        
        self.create_menu()
        self.create_ui()

        # Check for updates in background
        threading.Thread(target=self._check_updates, daemon=True).start()
    
    def _on_close(self):
        """Handle window close"""
        self._monitor_running = False
        self.root.destroy()

    def _check_updates(self):
        """Check for updates in background thread"""
        time.sleep(2)  # Wait for app to fully load
        update_info = UpdateChecker.check_for_updates()
        if update_info and update_info['available']:
            self.root.after(0, lambda: self._show_update_notification(update_info))

    def _show_update_notification(self, update_info):
        """Show update notification banner"""
        update_frame = Frame(self.root, bg='#4CAF50', padx=15, pady=10)
        update_frame.pack(fill=X, side=TOP, before=self.root.winfo_children()[0])

        Label(update_frame, text=f"Update Available: v{update_info['version']}",
              font=('SF Pro Text', 11, 'bold'), bg='#4CAF50', fg='white').pack(side=LEFT)

        Button(update_frame, text="Download Update", font=('SF Pro Text', 10),
               bg='white', fg='#4CAF50', relief=FLAT, padx=12, pady=4,
               cursor='hand2',
               command=lambda: self._download_and_install(update_info)).pack(side=RIGHT, padx=(0, 5))

        Button(update_frame, text="View Release Notes", font=('SF Pro Text', 10),
               bg='#45A049', fg='white', relief=FLAT, padx=12, pady=4,
               cursor='hand2',
               command=lambda: webbrowser.open(f"https://github.com/{__github_repo__}/releases/latest")).pack(side=RIGHT)

    def _download_and_install(self, update_info, skip_confirm=False):
        """Download and install update"""
        # Check if download URL is available (DMG/PKG asset exists)
        if not update_info.get('url'):
            response = messagebox.askyesno(
                "Manual Download Required",
                f"Version {update_info['version']} is available but no installer was found.\n\n"
                "This can happen due to GitHub rate limiting or if the release\n"
                "doesn't have a DMG file attached yet.\n\n"
                "Would you like to open the release page to download manually?"
            )
            if response:
                webbrowser.open(update_info.get('release_url', f"https://github.com/{__github_repo__}/releases/latest"))
            return

        if not skip_confirm:
            response = messagebox.askyesno(
                "Download Update",
                f"Download version {update_info['version']}?\n\n"
                "The installer will open after download."
            )
            if not response:
                return

        # Show progress dialog
        progress_window = Toplevel(self.root)
        progress_window.title("Downloading Update")
        progress_window.geometry("400x100")
        progress_window.resizable(False, False)
        progress_window.transient(self.root)

        Label(progress_window, text="Downloading update...", font=('SF Pro Text', 11)).pack(pady=10)

        progress_bar = ttk.Progressbar(progress_window, length=350, mode='determinate')
        progress_bar.pack(pady=10)

        def update_progress(pct):
            progress_bar['value'] = pct * 100
            progress_window.update()

        def do_download():
            installer_path = UpdateChecker.download_update(update_info['url'], update_progress)
            progress_window.destroy()

            # Validate the downloaded file is actually a DMG (not HTML error page)
            if installer_path and installer_path.endswith('.dmg'):
                file_size = os.path.getsize(installer_path) if os.path.exists(installer_path) else 0
                if file_size < 1000000:  # Less than 1MB - probably not a valid DMG
                    response = messagebox.askyesno(
                        "Download Issue",
                        f"The downloaded file appears to be invalid (only {file_size // 1024} KB).\n\n"
                        "This can happen due to network issues or GitHub rate limiting.\n\n"
                        "Would you like to open the release page to download manually?"
                    )
                    if response:
                        webbrowser.open(update_info.get('release_url', f"https://github.com/{__github_repo__}/releases/latest"))
                    return

            if installer_path and installer_path.endswith('.dmg'):
                # Auto-install from DMG
                try:
                    # Mount the DMG and get mount point from output
                    mount_result = subprocess.run(
                        ['hdiutil', 'attach', installer_path, '-nobrowse'],
                        capture_output=True, text=True
                    )
                    if mount_result.returncode != 0:
                        raise Exception(f"Failed to mount DMG: {mount_result.stderr}")

                    # Parse mount point from output (last line contains the volume path)
                    mount_point = None
                    for line in mount_result.stdout.strip().split('\n'):
                        if '/Volumes/' in line:
                            # Extract the path after /Volumes/
                            parts = line.split('\t')
                            for part in parts:
                                if '/Volumes/' in part:
                                    mount_point = part.strip()
                                    break

                    # Fallback: check known volume names
                    if not mount_point or not os.path.exists(mount_point):
                        import time
                        time.sleep(1)  # Wait for mount to complete
                        for vol in ['/Volumes/Garmin Workout Uploader', '/Volumes/GarminWorkoutUploader']:
                            if os.path.exists(vol):
                                mount_point = vol
                                break

                    if not mount_point or not os.path.exists(mount_point):
                        raise Exception("Could not find mounted volume")

                    # Find the .app in the mounted volume
                    app_source = None
                    for item in os.listdir(mount_point):
                        if item.endswith('.app'):
                            app_source = os.path.join(mount_point, item)
                            break

                    if not app_source:
                        raise Exception("Could not find app in DMG")

                    # Determine where to install - try to install to the same location as current app
                    current_app_path = None
                    install_location = "~/Applications"

                    # In a bundled .app, sys.executable is like:
                    # /Applications/Garmin Workout Uploader.app/Contents/MacOS/Garmin Workout Uploader
                    if getattr(sys, 'frozen', False):
                        exe_path = sys.executable
                        # Go up to find .app bundle
                        path = Path(exe_path)
                        for parent in path.parents:
                            if parent.suffix == '.app':
                                current_app_path = str(parent)
                                break

                    # Determine destination
                    if current_app_path and os.path.exists(current_app_path):
                        # Install to same location as current app
                        app_dest = current_app_path
                        install_location = os.path.dirname(current_app_path)
                    else:
                        # Fallback to ~/Applications
                        home_apps = os.path.expanduser('~/Applications')
                        os.makedirs(home_apps, exist_ok=True)
                        app_dest = os.path.join(home_apps, 'Garmin Workout Uploader.app')

                    # Try to remove old app and copy new one
                    try:
                        if os.path.exists(app_dest):
                            subprocess.run(['rm', '-rf', app_dest], check=True)

                        copy_result = subprocess.run(['cp', '-R', app_source, app_dest], capture_output=True, text=True)
                        if copy_result.returncode != 0:
                            raise PermissionError(copy_result.stderr)
                    except (PermissionError, subprocess.CalledProcessError) as perm_err:
                        # Permission denied - fall back to ~/Applications
                        home_apps = os.path.expanduser('~/Applications')
                        os.makedirs(home_apps, exist_ok=True)
                        app_dest = os.path.join(home_apps, 'Garmin Workout Uploader.app')
                        install_location = "~/Applications"

                        if os.path.exists(app_dest):
                            subprocess.run(['rm', '-rf', app_dest], check=True)

                        copy_result = subprocess.run(['cp', '-R', app_source, app_dest], capture_output=True, text=True)
                        if copy_result.returncode != 0:
                            raise Exception(f"Failed to copy app: {copy_result.stderr}")

                    # Unmount the DMG
                    subprocess.run(['hdiutil', 'detach', mount_point, '-quiet', '-force'])

                    # Ask to restart
                    response = messagebox.askyesno(
                        "Update Installed",
                        f"Version {update_info['version']} has been installed to:\n"
                        f"{install_location}\n\n"
                        f"Restart the app now to use the new version?"
                    )
                    if response:
                        # Relaunch the app
                        subprocess.Popen(['open', app_dest])
                        self.root.quit()

                except Exception as e:
                    # Fallback to manual install
                    subprocess.run(['open', installer_path])
                    messagebox.showinfo("Auto-Install Issue",
                        f"Auto-install encountered an issue:\n{str(e)}\n\n"
                        f"The DMG has been opened. Please drag the app to Applications manually.")
            elif installer_path:
                # Non-DMG file was downloaded - this shouldn't happen but handle gracefully
                response = messagebox.askyesno(
                    "Unexpected File Type",
                    f"The downloaded file is not a DMG installer.\n\n"
                    "Would you like to open the release page to download manually?"
                )
                if response:
                    webbrowser.open(update_info.get('release_url', f"https://github.com/{__github_repo__}/releases/latest"))
            else:
                response = messagebox.askyesno(
                    "Download Failed",
                    "Could not download the update.\n\n"
                    "Would you like to open the release page to download manually?"
                )
                if response:
                    webbrowser.open(update_info.get('release_url', f"https://github.com/{__github_repo__}/releases/latest"))

        threading.Thread(target=do_download, daemon=True).start()

    def check_for_updates_manual(self):
        """Manually check for updates from menu"""
        # Show checking dialog
        checking_window = Toplevel(self.root)
        checking_window.title("Checking for Updates")
        checking_window.geometry("300x80")
        checking_window.resizable(False, False)
        checking_window.transient(self.root)
        checking_window.grab_set()

        # Center on parent
        checking_window.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 300) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 80) // 2
        checking_window.geometry(f"+{x}+{y}")

        Label(checking_window, text="🔄 Checking for updates...",
              font=('SF Pro Text', 12)).pack(expand=True)

        def do_check():
            update_info = UpdateChecker.check_for_updates()
            self.root.after(0, lambda: show_result(update_info))

        def show_result(update_info):
            try:
                checking_window.destroy()
            except:
                pass
            if update_info is None:
                messagebox.showerror("Error", "Could not check for updates.\nPlease check your internet connection.")
            elif update_info['available']:
                # Show dialog asking to download
                response = messagebox.askyesno(
                    "Update Available",
                    f"A new version is available!\n\n"
                    f"Current version: v{__version__}\n"
                    f"New version: v{update_info['version']}\n\n"
                    f"Would you like to download the update?"
                )
                if response:
                    self._download_and_install(update_info, skip_confirm=True)
            else:
                messagebox.showinfo("Up to Date", f"You're running the latest version (v{__version__}).")

        threading.Thread(target=do_check, daemon=True).start()

    def check_openmtp(self):
        """Check if OpenMTP is installed"""
        paths = [
            Path("/Applications/OpenMTP.app"),
            self.home / "Applications/OpenMTP.app"
        ]
        return any(p.exists() for p in paths)
    
    def check_libmtp(self):
        """Check if libmtp is installed via Homebrew"""
        try:
            result = subprocess.run(['which', 'mtp-detect'], 
                                   capture_output=True, text=True, timeout=5)
            return result.returncode == 0
        except:
            return False
    
    def detect_garmin_device(self):
        """Detect connected Garmin device via USB"""
        # Try system_profiler first
        device = self._detect_via_system_profiler()
        if device:
            return device
        
        # Fallback to ioreg for MTP devices
        device = self._detect_via_ioreg()
        if device:
            return device
        
        return None
    
    def _detect_via_system_profiler(self):
        """Detect Garmin via system_profiler"""
        try:
            result = subprocess.run(
                ['system_profiler', 'SPUSBDataType'],
                capture_output=True, text=True, timeout=10
            )
            
            if result.returncode != 0:
                return None
            
            output = result.stdout
            output_lower = output.lower()
            
            # Look for Garmin device patterns
            garmin_patterns = [
                'garmin', 'forerunner', 'fenix', 'edge', 'vivoactive', 
                'venu', 'instinct', 'marq', 'enduro', 'epix', 'approach'
            ]
            
            # Check if any Garmin-related text exists
            found = any(p in output_lower for p in garmin_patterns)
            
            # Also check vendor ID (0x091e)
            if not found:
                found = 'vendor id: 0x091e' in output_lower or '091e' in output_lower
            
            if not found:
                return None
            
            # Try to extract device name
            lines = output.split('\n')
            device_name = None
            
            for i, line in enumerate(lines):
                line_lower = line.lower()
                if any(g in line_lower for g in garmin_patterns):
                    name_match = re.search(r'^\s*(.+?):', line)
                    if name_match:
                        device_name = name_match.group(1).strip()
                    break
            
            return {
                'connected': True,
                'name': device_name or 'Garmin Device',
                'vendor_id': '091e'
            }
            
        except:
            return None
    
    def _detect_via_ioreg(self):
        """Detect Garmin via ioreg (for MTP devices)"""
        try:
            result = subprocess.run(
                ['ioreg', '-p', 'IOUSB', '-l', '-w', '0'],
                capture_output=True, text=True, timeout=10
            )
            
            if result.returncode != 0:
                return None
            
            output = result.stdout
            
            # Look for Garmin signature directly in the output
            # Signature format: <1e09XXYY...> where 1e09 is Garmin vendor ID (little-endian)
            # and XXYY is product ID (little-endian)
            sig_pattern = re.search(r'"UsbDeviceSignature"\s*=\s*<1e09([a-f0-9]{4})', output, re.IGNORECASE)
            
            if not sig_pattern:
                return None
            
            # Extract product ID from signature (little-endian)
            hex_pid = sig_pattern.group(1)
            product_id = int(hex_pid[2:4] + hex_pid[0:2], 16)
            
            # Map known Garmin product IDs to names
            garmin_products = {
                # Special modes
                3: None,  # Charging/initializing mode - will be handled below
                
                # Fenix series
                20920: "Fenix 8",
                20921: "Fenix 8 Solar",
                20922: "Fenix 8 AMOLED",
                20736: "Fenix 7",
                20737: "Fenix 7S",
                20738: "Fenix 7X",
                20480: "Fenix 6",
                20481: "Fenix 6S",
                20482: "Fenix 6X",
                
                # Forerunner series
                20224: "Forerunner 965",
                20096: "Forerunner 265",
                20097: "Forerunner 265S", 
                19968: "Forerunner 955",
                19840: "Forerunner 255",
                19712: "Forerunner 945",
                19584: "Forerunner 745",
                
                # Epix
                20352: "Epix Gen 2",
                20353: "Epix Pro",
                
                # Venu
                19456: "Venu 2",
                19457: "Venu 2S",
                19328: "Venu",
                
                # Instinct
                19200: "Instinct 2",
                19201: "Instinct 2S",
                
                # Edge
                18944: "Edge 1040",
                18688: "Edge 840",
                18432: "Edge 540",
                18176: "Edge 530",
            }
            
            # Handle special modes
            if product_id == 3:
                return {
                    'connected': True,
                    'name': "Garmin Watch (initializing...)",
                    'vendor_id': '091e',
                    'product_id': product_id,
                    'mode': 'charging'
                }
            elif product_id in garmin_products:
                device_name = garmin_products[product_id]
            else:
                device_name = f"Garmin Watch (ID:{product_id})"
            
            return {
                'connected': True,
                'name': device_name,
                'vendor_id': '091e',
                'product_id': product_id,
                'mode': 'mtp'
            }
            
        except Exception as e:
            return None
    
    def check_garmin_express_running(self):
        """Check if Garmin Express is running (blocks MTP)"""
        try:
            result = subprocess.run(['pgrep', '-f', 'Garmin Express'],
                                   capture_output=True, text=True, timeout=5)
            return result.returncode == 0
        except:
            return False
    
    def kill_garmin_express(self):
        """Kill Garmin Express if running"""
        try:
            subprocess.run(['pkill', '-f', 'Garmin Express'], timeout=5, capture_output=True)
            subprocess.run(['pkill', '-f', 'GarminExpressService'], timeout=5, capture_output=True)
            return True
        except:
            return False
    
    def refresh_device_status(self):
        """Refresh the device connection status"""
        # Check if widgets still exist
        try:
            if not self.device_status.winfo_exists():
                return
        except:
            return
        
        device = self.detect_garmin_device()
        garmin_express_running = self.check_garmin_express_running()

        # Store detected device for model-specific adjustments
        self.current_device = device

        # Remove any existing close button
        if hasattr(self, 'close_ge_btn') and self.close_ge_btn:
            try:
                self.close_ge_btn.destroy()
            except:
                pass
            self.close_ge_btn = None
        
        try:
            if device:
                # Check if device is in charging/initializing mode
                if device.get('mode') == 'charging':
                    self.device_status.config(text=f"🔄 {device['name']}", fg='#007AFF')
                    self.device_status_detail.config(text="Wait for watch to enter MTP mode...")
                elif garmin_express_running:
                    self.device_status.config(text=f"⚠️ {device['name']} detected", fg='#FF9500')
                    self.device_status_detail.config(text="Garmin Express is blocking - close it to transfer")
                    
                    # Add close button in the status container
                    parent_frame = self.device_status_detail.master
                    self.close_ge_btn = Button(parent_frame, text="Close Garmin Express",
                                              font=('SF Pro Text', 11), bg='#FF9500', fg='white',
                                              command=self.close_garmin_express_clicked, relief=FLAT,
                                              cursor='hand2', padx=10, pady=4)
                    self.close_ge_btn.pack(anchor='w', pady=(8, 0))
                else:
                    self.device_status.config(text=f"✅ {device['name']} connected", fg='#28a745')
                    self.device_status_detail.config(text="Ready for transfer")
            else:
                self.device_status.config(text="❌ No Garmin device detected", fg='#dc3545')
                self.device_status_detail.config(text="Connect watch via USB (keep screen awake)")
        except:
            pass  # Widget was destroyed
    
    def close_garmin_express_clicked(self):
        """Handle Close Garmin Express button click"""
        self.kill_garmin_express()
        self.device_status.config(text="🔄 Closing Garmin Express...", fg='#666')
        self.root.after(1500, self.refresh_device_status)
    
    def start_device_monitor(self):
        """Start background thread to monitor device connection"""
        def monitor():
            while self._monitor_running:
                try:
                    self.root.after(0, self.refresh_device_status)
                except:
                    break
                time.sleep(3)  # Check every 3 seconds
        
        thread = threading.Thread(target=monitor, daemon=True)
        thread.start()
    
    def create_menu(self):
        """Create the application menu bar"""
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        # Tools menu
        tools_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
        # Local tools
        tools_menu.add_command(label="👁 Preview FIT File...", command=self.preview_fit_file_dialog)
        tools_menu.add_separator()
        
        # GOTOES online tools
        tools_menu.add_command(label="🔧 Repair FIT File", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/Combine_FIT_Files.php'))
        tools_menu.add_command(label="🔗 Merge FIT/GPX Files", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/Combine_GPX_TCX_FIT_Files.php'))
        tools_menu.add_command(label="📊 View FIT File Data", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/View_FIT_Data.php'))
        tools_menu.add_command(label="🕐 Add Timestamps to GPX", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/Add_Timestamps_To_GPX.php'))
        tools_menu.add_separator()
        tools_menu.add_command(label="📉 Shrink FIT File", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/Shrink_FIT_File.php'))
        tools_menu.add_command(label="⏱️ Time-Shift Activity", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/Adjust_Activity_Time.php'))
        tools_menu.add_command(label="🏁 Race Repair (GPS)", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/Race_Repair.php'))
        tools_menu.add_separator()
        tools_menu.add_command(label="🌐 All GOTOES Tools...", 
                              command=lambda: webbrowser.open('https://gotoes.org/strava/index.php'))
        
        # Help menu
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="How to Use", command=self.show_help)
        help_menu.add_command(label="Get OpenMTP",
                             command=lambda: webbrowser.open('https://openmtp.ganeshrvel.com'))
        help_menu.add_separator()
        help_menu.add_command(label="Check for Updates...", command=self.check_for_updates_manual)
        help_menu.add_command(label="About", command=self.show_about)
    
    def create_ui(self):
        """Create the main interface"""
        # Main container
        main = Frame(self.root, bg='#f5f5f7', padx=30, pady=25)
        main.pack(fill=BOTH, expand=True)
        
        # Header
        ttk.Label(main, text="Garmin Workout Uploader", style="Title.TLabel").pack()
        ttk.Label(main, text="Upload .FIT workouts to your Garmin watch", style="Subtitle.TLabel").pack(pady=(5, 20))
        
        # Step 1: Select Files
        self.create_step(main, "1", "Select Your Workout Files", self.create_file_selector)
        
        # Step 2: Prepare Transfer  
        self.create_step(main, "2", "Prepare for Transfer", self.create_prepare_section)
        
        # Step 3: Transfer Files
        self.create_step(main, "3", "Transfer to Watch", self.create_transfer_section)
        
        # Help link
        help_frame = Frame(main, bg='#f5f5f7')
        help_frame.pack(fill=X, pady=(15, 0))
        
        help_btn = Label(help_frame, text="Need help? Click here", fg='#007AFF', bg='#f5f5f7',
                        cursor='hand2', font=('SF Pro Text', 11, 'underline'))
        help_btn.pack()
        help_btn.bind('<Button-1>', lambda e: self.show_help())
    
    def create_step(self, parent, number, title, content_func):
        """Create a step card"""
        # Card frame
        card = Frame(parent, bg='#fff', highlightbackground='#e0e0e0', 
                    highlightthickness=1, padx=15, pady=12)
        card.pack(fill=X, pady=(0, 12))
        
        # Header row
        header = Frame(card, bg='#fff')
        header.pack(fill=X, pady=(0, 10))
        
        # Step number circle
        num_canvas = Canvas(header, width=28, height=28, bg='#fff', highlightthickness=0)
        num_canvas.pack(side=LEFT, padx=(0, 10))
        num_canvas.create_oval(2, 2, 26, 26, fill='#007AFF', outline='')
        num_canvas.create_text(14, 14, text=number, fill='white', font=('SF Pro Display', 13, 'bold'))
        
        # Title
        Label(header, text=title, font=('SF Pro Text', 13, 'bold'), bg='#fff').pack(side=LEFT)
        
        # Content
        content_frame = Frame(card, bg='#fff')
        content_frame.pack(fill=X)
        content_func(content_frame)
    
    def create_file_selector(self, parent):
        """Step 1: File selection with drag and drop support"""
        # Drop zone frame (for visual feedback)
        self.drop_zone = Frame(parent, bg='#fff')
        self.drop_zone.pack(fill=X)
        
        # Listbox (EXTENDED mode for multi-select)
        self.file_listbox = Listbox(self.drop_zone, height=4, font=('SF Pro Text', 11),
                                     selectmode=EXTENDED,
                                     selectbackground='#007AFF', activestyle='none',
                                     highlightthickness=2, highlightbackground='#e0e0e0',
                                     highlightcolor='#007AFF', relief=FLAT)
        self.file_listbox.pack(fill=X, pady=(0, 8))
        
        # Set up drag and drop if available
        if DND_AVAILABLE:
            self.file_listbox.drop_target_register(DND_FILES)
            self.file_listbox.dnd_bind('<<DropEnter>>', self.on_drag_enter)
            self.file_listbox.dnd_bind('<<DropLeave>>', self.on_drag_leave)
            self.file_listbox.dnd_bind('<<Drop>>', self.on_drop)
            
            # Placeholder text with drag hint
            self.file_listbox.insert(END, "  Drop .FIT files here or click 'Add Files'")
        else:
            # Placeholder text without drag hint
            self.file_listbox.insert(END, "  No files selected - click 'Add Files' below")
        
        self.file_listbox.config(fg='#999')
        
        # Buttons
        btn_frame = Frame(parent, bg='#fff')
        btn_frame.pack(fill=X)
        
        self.add_btn = Button(btn_frame, text="＋ Add Files", font=('SF Pro Text', 11),
                              command=self.add_files, bg='#007AFF', fg='white',
                              padx=15, pady=5, relief=FLAT, cursor='hand2')
        self.add_btn.pack(side=LEFT)
        
        self.clear_btn = Button(btn_frame, text="Clear", font=('SF Pro Text', 11),
                                command=self.clear_files, padx=10, pady=5, relief=FLAT)
        self.clear_btn.pack(side=LEFT, padx=(8, 0))
        
        self.preview_btn = Button(btn_frame, text="👁 Preview", font=('SF Pro Text', 11),
                                  command=self.preview_selected_file, padx=10, pady=5, relief=FLAT)
        self.preview_btn.pack(side=LEFT, padx=(8, 0))
        
        # File count on its own row for visibility
        count_frame = Frame(parent, bg='#fff')
        count_frame.pack(fill=X, pady=(5, 0))
        
        self.file_count = Label(count_frame, text="", font=('SF Pro Text', 11), bg='#fff', fg='#666')
        self.file_count.pack(side=LEFT)
        
        # Drag and drop status indicator
        if DND_AVAILABLE:
            self.dnd_status = Label(count_frame, text="📥 Drop enabled", font=('SF Pro Text', 10), 
                                    bg='#fff', fg='#34C759')
            self.dnd_status.pack(side=RIGHT)
    
    def on_drag_enter(self, event):
        """Visual feedback when files are dragged over the listbox"""
        self.is_dragging = True
        self.file_listbox.config(highlightbackground='#007AFF', highlightthickness=3)
        self.drop_zone.config(bg='#e3f2fd')
        return event.action
    
    def on_drag_leave(self, event):
        """Reset visual feedback when drag leaves"""
        self.is_dragging = False
        self.file_listbox.config(highlightbackground='#e0e0e0', highlightthickness=2)
        self.drop_zone.config(bg='#fff')
        return event.action
    
    def on_drop(self, event):
        """Handle dropped files"""
        self.is_dragging = False
        self.file_listbox.config(highlightbackground='#e0e0e0', highlightthickness=2)
        self.drop_zone.config(bg='#fff')
        
        # Parse dropped file paths
        # On macOS, paths may be space-separated or in braces
        files = self.parse_drop_data(event.data)
        
        if not files:
            return
        
        # Add the files
        self.add_files_to_list(files)
    
    def parse_drop_data(self, data):
        """Parse the dropped file data from tkinterdnd2"""
        files = []
        
        # Handle different formats
        # Format 1: {/path/to/file1} {/path/to/file2}
        # Format 2: /path/to/file1 /path/to/file2
        
        if '{' in data:
            # Files are wrapped in braces (common on macOS)
            import re
            matches = re.findall(r'\{([^}]+)\}', data)
            files = matches
        else:
            # Try to split by spaces, but handle spaces in filenames
            # This is tricky - assume each path starts with /
            parts = data.split()
            current_path = ""
            for part in parts:
                if part.startswith('/') and current_path:
                    files.append(current_path)
                    current_path = part
                elif part.startswith('/'):
                    current_path = part
                else:
                    current_path += ' ' + part
            if current_path:
                files.append(current_path)
        
        # Filter to only .fit files
        fit_files = [f for f in files if f.lower().endswith('.fit')]
        
        return fit_files
    
    def add_files_to_list(self, files):
        """Add files to the selection list"""
        if not files:
            return
        
        # Clear placeholder if this is the first file
        if not self.selected_files:
            self.file_listbox.delete(0, END)
            self.file_listbox.config(fg='black')
        
        added_count = 0
        for f in files:
            if f not in self.selected_files:
                if f.lower().endswith('.fit'):
                    self.selected_files.append(f)
                    name = os.path.basename(f)
                    self.file_listbox.insert(END, f"  📄 {name}")
                    added_count += 1
        
        if added_count > 0:
            self.update_ui_state()
            # Flash success feedback
            self.file_listbox.config(highlightbackground='#34C759')
            self.root.after(300, lambda: self.file_listbox.config(highlightbackground='#e0e0e0'))
    
    def create_prepare_section(self, parent):
        """Step 2: Prepare transfer"""
        # Instructions
        instructions = Frame(parent, bg='#fff')
        instructions.pack(fill=X)
        
        steps_text = """Before transferring, make sure:

✓  Your Garmin watch is connected via USB
✓  On your watch: Settings → System → USB Mode → MTP
✓  Accept "Use MTP" prompt on the watch if asked
✓  Garmin Express is closed (quit it if running)"""
        
        Label(instructions, text=steps_text, font=('SF Pro Text', 11), bg='#fff',
              justify=LEFT, anchor='w').pack(fill=X)
        
        # Prepare button
        self.prepare_btn = Button(parent, text="✓ Ready - Stage My Files", 
                                   font=('SF Pro Text', 12, 'bold'),
                                   command=self.stage_files, bg='#34C759', fg='white',
                                   padx=20, pady=8, relief=FLAT, cursor='hand2',
                                   state=DISABLED)
        self.prepare_btn.pack(pady=(12, 0))
    
    def create_transfer_section(self, parent):
        """Step 3: Transfer"""
        self.transfer_frame = parent
        
        # Device status indicator
        device_frame = Frame(parent, bg='#f0f0f0', padx=12, pady=12,
                            highlightbackground='#ccc', highlightthickness=1)
        device_frame.pack(fill=X, pady=(0, 10))
        
        # Header row with refresh button
        header_row = Frame(device_frame, bg='#f0f0f0')
        header_row.pack(fill=X)
        
        Label(header_row, text="Device Status:", font=('SF Pro Text', 11, 'bold'),
              bg='#f0f0f0', fg='#333').pack(side=LEFT)
        
        # Use a proper styled button
        self.refresh_btn = Button(header_row, text="↻ Refresh", font=('SF Pro Text', 11),
                            bg='white', fg='#007AFF', relief=SOLID, cursor='hand2',
                            borderwidth=1, padx=12, pady=4, 
                            activebackground='#007AFF', activeforeground='white',
                            command=self._refresh_clicked)
        self.refresh_btn.pack(side=RIGHT)
        
        # Status container for proper layout
        status_container = Frame(device_frame, bg='#f0f0f0')
        status_container.pack(fill=X, pady=(10, 0))
        
        # Main status text
        self.device_status = Label(status_container, text="🔍 Checking for device...",
                                   font=('SF Pro Text', 13, 'bold'), bg='#f0f0f0', fg='#666',
                                   anchor='w')
        self.device_status.pack(fill=X)
        
        # Detail/tip text
        self.device_status_detail = Label(status_container, text="Please wait...",
                                          font=('SF Pro Text', 11), bg='#f0f0f0', fg='#666',
                                          anchor='w')
        self.device_status_detail.pack(fill=X, pady=(2, 0))
        
        # Initial state - waiting
        self.transfer_status = Label(parent, 
            text="Stage your files first (Step 2), then transfer instructions will appear here.",
            font=('SF Pro Text', 11), bg='#fff', fg='#666', wraplength=480, justify=LEFT)
        self.transfer_status.pack(fill=X, pady=(5, 0))
        
        # Start device monitoring after UI is built
        self.root.after(500, self.refresh_device_status)
        self.root.after(1000, self.start_device_monitor)
    
    def _refresh_clicked(self):
        """Handle refresh button click with visual feedback"""
        self.refresh_btn.config(text="⏳ Checking...", state=DISABLED)
        self.device_status.config(text="🔍 Checking for device...", fg='#666')
        self.device_status_detail.config(text="Please wait...")
        self.root.update()
        
        # Do the refresh
        self.refresh_device_status()
        
        # Reset button
        self.root.after(500, lambda: self.refresh_btn.config(text="↻ Refresh", state=NORMAL))
    
    def add_files(self):
        """Open file dialog to add .FIT files"""
        files = filedialog.askopenfilenames(
            title="Select Workout Files",
            filetypes=[("FIT files", "*.fit *.FIT"), ("All files", "*.*")]
        )
        
        if files:
            self.add_files_to_list(list(files))
    
    def clear_files(self):
        """Clear all selected files"""
        self.selected_files = []
        self.file_listbox.delete(0, END)
        
        if DND_AVAILABLE:
            self.file_listbox.insert(END, "  Drop .FIT files here or click 'Add Files'")
        else:
            self.file_listbox.insert(END, "  No files selected - click 'Add Files' below")
        
        self.file_listbox.config(fg='#999')
        self.update_ui_state()
    
    def update_ui_state(self):
        """Update button states based on current state"""
        count = len(self.selected_files)
        
        if count > 0:
            self.file_count.config(text=f"{count} file{'s' if count > 1 else ''} selected")
            self.prepare_btn.config(state=NORMAL)
        else:
            self.file_count.config(text="")
            self.prepare_btn.config(state=DISABLED)
    
    def stage_files(self):
        """Copy files to staging folder and prepare for transfer"""
        if not self.selected_files:
            return
        
        # Kill Garmin Express
        self.kill_garmin_express()
        
        # Clear staging folder
        for f in self.staging_folder.glob('*.fit'):
            f.unlink()
        for f in self.staging_folder.glob('*.FIT'):
            f.unlink()
        
        # Copy files
        staged = []
        for filepath in self.selected_files:
            try:
                filename = os.path.basename(filepath)
                dest = self.staging_folder / filename
                shutil.copy2(filepath, dest)
                staged.append(filename)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to copy {filename}: {e}")
        
        if not staged:
            return
        
        # Update Step 3 with transfer instructions (keep device status)
        self.show_transfer_instructions(staged)
        
        # Open staging folder and OpenMTP
        subprocess.run(['open', str(self.staging_folder)])
        self.open_openmtp()
        
        # Show success
        self.prepare_btn.config(text="✓ Files Staged!", bg='#666', state=DISABLED)
    
    
    def show_transfer_instructions(self, staged_files):
        """Show transfer instructions in Step 3 while keeping device status"""
        # Only clear transfer_status label, not device status
        if hasattr(self, 'transfer_status'):
            self.transfer_status.config(
                text=f"✓ {len(staged_files)} file(s) ready! Drag from Finder → OpenMTP (GARMIN/NewFiles)",
                fg='#2e7d32'
            )
        
        # Clean up existing buttons frame
        if hasattr(self, 'transfer_btns_frame') and self.transfer_btns_frame:
            try:
                self.transfer_btns_frame.destroy()
            except:
                pass
        
        # Clean up existing warning frame
        if hasattr(self, 'openmtp_warning_frame') and self.openmtp_warning_frame:
            try:
                self.openmtp_warning_frame.destroy()
            except:
                pass
        
        # Add helper buttons below transfer_status
        self.transfer_btns_frame = Frame(self.transfer_frame, bg='#fff')
        self.transfer_btns_frame.pack(fill=X, pady=(8, 0))
        
        Button(self.transfer_btns_frame, text="📂 Open Folder", font=('SF Pro Text', 11),
               command=lambda: subprocess.run(['open', str(self.staging_folder)]),
               padx=10, pady=5, relief=FLAT, cursor='hand2').pack(side=LEFT)
        
        Button(self.transfer_btns_frame, text="🔄 OpenMTP", font=('SF Pro Text', 11),
               command=self.open_openmtp, padx=10, pady=5, relief=FLAT, cursor='hand2').pack(side=LEFT, padx=(8, 0))
        
        # If OpenMTP not installed
        if not self.openmtp_installed:
            self.openmtp_warning_frame = Frame(self.transfer_frame, bg='#fff3e0', padx=10, pady=8)
            self.openmtp_warning_frame.pack(fill=X, pady=(10, 0))
            
            Label(self.openmtp_warning_frame, text="⚠️ OpenMTP not found!", 
                  font=('SF Pro Text', 11, 'bold'), bg='#fff3e0', fg='#e65100').pack()
            
            Label(self.openmtp_warning_frame, text="Download it free from: openmtp.ganeshrvel.com", 
                  font=('SF Pro Text', 11), bg='#fff3e0').pack()
            
            Button(self.openmtp_warning_frame, text="Download OpenMTP", font=('SF Pro Text', 11),
                   command=lambda: webbrowser.open('https://openmtp.ganeshrvel.com'),
                   bg='#ff9800', fg='white', padx=10, pady=5, relief=FLAT,
                   cursor='hand2').pack(pady=(5, 0))
    
    def open_openmtp(self):
        """Open OpenMTP application"""
        paths = [
            "/Applications/OpenMTP.app",
            str(self.home / "Applications/OpenMTP.app")
        ]
        
        for path in paths:
            if os.path.exists(path):
                subprocess.run(['open', path])
                return True
        
        # Try Android File Transfer as fallback
        aft = "/Applications/Android File Transfer.app"
        if os.path.exists(aft):
            subprocess.run(['open', aft])
            return True
        
        return False
    
    def show_help(self):
        """Show help dialog"""
        help_window = Toplevel(self.root)
        help_window.title("Help")
        help_window.geometry("500x450")
        help_window.configure(bg='#f5f5f7')
        help_window.transient(self.root)
        
        frame = Frame(help_window, bg='#f5f5f7', padx=25, pady=20)
        frame.pack(fill=BOTH, expand=True)
        
        Label(frame, text="Help & Troubleshooting", font=('SF Pro Display', 18, 'bold'),
              bg='#f5f5f7').pack(pady=(0, 15))
        
        help_text = """Why do I need OpenMTP?
Mac doesn't support MTP (the protocol Garmin uses).
OpenMTP bridges this gap - it's free and works great.

Watch not showing in OpenMTP?
• Make sure USB cable supports data (not charge-only)
• On watch: Settings → System → USB Mode → MTP
• Quit Garmin Express completely
• Unplug and replug the watch
• Click "Refresh" in OpenMTP

Can't find NewFiles folder?
Look for: GARMIN → NewFiles
If only "Workouts" exists, use that instead.

Workouts not appearing on watch?
• Restart your watch after transfer
• Check: Training → Workouts
• Make sure files are valid .FIT workout files

Where do workouts come from?
• Create in Garmin Connect (web or app)
• Export from TrainingPeaks, Intervals.icu, etc.
• Download from training plan providers

My files are stuck in GarminWorkouts folder?
That's just the staging folder on your Mac.
You still need to drag them to OpenMTP."""
        
        Label(frame, text=help_text, font=('SF Pro Text', 11), bg='#f5f5f7',
              justify=LEFT, anchor='w').pack(fill=X)
        
        Button(frame, text="Get OpenMTP", font=('SF Pro Text', 11),
               command=lambda: webbrowser.open('https://openmtp.ganeshrvel.com'),
               bg='#007AFF', fg='white', padx=15, pady=8, relief=FLAT,
               cursor='hand2').pack(pady=(15, 10))
        
        Button(frame, text="Close", command=help_window.destroy,
               font=('SF Pro Text', 11), padx=15, pady=5, relief=FLAT).pack()
    
    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo("About",
            f"{__app_name__}\n\n"
            f"Version {__version__}\n\n"
            "A simple tool to upload .FIT workout files\n"
            "to your Garmin watch via OpenMTP.\n\n"
            "Tools menu powered by GOTOES.org")
    
    def preview_fit_file_dialog(self):
        """Open file dialog to select and preview a FIT file"""
        filepath = filedialog.askopenfilename(
            title="Select FIT File to Preview",
            filetypes=[("FIT files", "*.fit *.FIT"), ("All files", "*.*")]
        )
        if filepath:
            self.show_fit_preview(filepath)
    
    def preview_selected_file(self):
        """Preview the currently selected FIT file(s)"""
        selection = self.file_listbox.curselection()
        
        if not selection:
            # If nothing selected but files exist, preview all
            if self.selected_files:
                filepaths = self.selected_files
            else:
                messagebox.showinfo("Preview", "Please select a .FIT file first")
                return
        else:
            # Get all selected files
            filepaths = []
            for idx in selection:
                if idx < len(self.selected_files):
                    filepaths.append(self.selected_files[idx])
        
        if len(filepaths) == 1:
            self.show_fit_preview(filepaths[0])
        else:
            self.show_fit_preview_multi(filepaths)
    
    def show_fit_preview(self, filepath):
        """Show FIT file preview matching AmakaFlow app style"""
        # Parse the FIT file
        workout_data = self.parse_fit_file(filepath)

        if not workout_data:
            messagebox.showerror("Error", "Could not parse FIT file. It may be corrupted or not a workout file.")
            return

        # Validate the FIT file for issues
        validation = self.validate_fit_file(filepath)

        # Create preview window
        preview = Toplevel(self.root)
        preview.title(f"Workout Preview - {os.path.basename(filepath)}")
        preview.geometry("450x750" if not validation['valid'] else "450x700")
        preview.configure(bg='#1a1a1a')
        preview.transient(self.root)

        # Unbind mousewheel on close
        def on_close():
            try:
                preview.unbind_all("<MouseWheel>")
            except:
                pass
            preview.destroy()
        preview.protocol("WM_DELETE_WINDOW", on_close)

        # Main container with dark theme
        main = Frame(preview, bg='#1a1a1a', padx=20, pady=20)
        main.pack(fill=BOTH, expand=True)

        # Warning banner if validation failed
        if not validation['valid']:
            warning_frame = Frame(main, bg='#dc3545', padx=10, pady=8)
            warning_frame.pack(fill=X, pady=(0, 10))

            Label(warning_frame, text="⚠️ Compatibility Issue Detected",
                  font=('SF Pro Text', 11, 'bold'), bg='#dc3545', fg='#fff').pack(anchor='w')

            for issue in validation['issues'][:2]:  # Show first 2 issues
                Label(warning_frame, text=issue, font=('SF Pro Text', 9),
                      bg='#dc3545', fg='#fff', wraplength=400, justify=LEFT).pack(anchor='w')

            # Repair button if fitfiletool is available
            if FITFILETOOL_AVAILABLE:
                def do_repair():
                    new_file, error = self.repair_fit_file(filepath, workout_data)
                    if new_file:
                        messagebox.showinfo("Repaired",
                            f"Workout repaired and saved to:\n{os.path.basename(new_file)}\n\n"
                            "The repaired file uses valid exercise categories that work on all Garmin watches.")
                        # Add repaired file to selection
                        if new_file not in self.selected_files:
                            self.selected_files.append(new_file)
                            self.file_listbox.insert(END, f"  📄 {os.path.basename(new_file)} (repaired)")
                            self.file_count.config(text=f"{len(self.selected_files)} file(s) selected")
                    else:
                        messagebox.showerror("Error", f"Could not repair file:\n{error}")

                Button(warning_frame, text="🔧 Repair Workout", font=('SF Pro Text', 10, 'bold'),
                       command=do_repair, bg='#fff', fg='#dc3545',
                       padx=12, pady=4, relief=FLAT, cursor='hand2').pack(anchor='w', pady=(5, 0))

        # Watch face simulation (rounded rectangle effect)
        watch_frame = Frame(main, bg='#000', highlightbackground='#333',
                           highlightthickness=2, padx=15, pady=15)
        watch_frame.pack(fill=BOTH, expand=True, pady=(0, 15))

        # Workout title
        title = workout_data.get('name', 'Workout')
        Label(watch_frame, text=title, font=('SF Pro Display', 14, 'bold'),
              bg='#000', fg='#fff', wraplength=380).pack(pady=(5, 5))

        # Sport type badge
        sport = workout_data.get('sport')
        sub_sport = workout_data.get('sub_sport')
        if sport:
            sport_display = get_sport_display(sport, sub_sport)
            sport_color = get_sport_color(sport, sub_sport)

            sport_badge = Label(watch_frame, text=f"  {sport_display}  ",
                               font=('SF Pro Text', 10, 'bold'),
                               bg=sport_color, fg='#fff')
            sport_badge.pack(pady=(0, 5))

        # Metadata row (source + date)
        meta_frame = Frame(watch_frame, bg='#000')
        meta_frame.pack(fill=X, pady=(0, 10))

        meta_parts = []
        if workout_data.get('source'):
            meta_parts.append(f"📱 {workout_data['source']}")
        if workout_data.get('created'):
            # Format date nicely
            created = workout_data['created'].split(' ')[0] if ' ' in workout_data['created'] else workout_data['created']
            meta_parts.append(f"📅 {created}")

        # Calculate total duration
        total_duration = sum(ex.get('duration', 0) for ex in workout_data.get('steps', []))
        if total_duration > 0:
            meta_parts.append(f"⏱ {self.format_duration(total_duration)}")

        if meta_parts:
            Label(meta_frame, text="  •  ".join(meta_parts), font=('SF Pro Text', 9),
                  bg='#000', fg='#666').pack()

        # Scrollable exercise list
        canvas = Canvas(watch_frame, bg='#000', highlightthickness=0, height=400)
        scrollbar = Scrollbar(watch_frame, orient=VERTICAL, command=canvas.yview)
        exercise_frame = Frame(canvas, bg='#000')

        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)

        canvas_window = canvas.create_window((0, 0), window=exercise_frame, anchor='nw')

        # Bind canvas resize
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
            canvas.itemconfig(canvas_window, width=event.width)

        exercise_frame.bind('<Configure>', configure_canvas)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        # Mouse wheel scrolling - scoped to this canvas
        def on_mousewheel(event):
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Process steps to detect repeat structures
        exercises = workout_data.get('steps', [])
        processed_steps = self.process_steps_for_preview(exercises)

        # Stats counters
        exercise_count = 0
        rest_count = 0
        total_sets = 0
        repeat_count = 0

        # Display processed steps
        for step_info in processed_steps:
            step_type = step_info.get('display_type', 'exercise')

            if step_type == 'repeat_header':
                self.create_repeat_header(exercise_frame, step_info)
                repeat_count += 1
                total_sets += step_info.get('repeat_count', 1)
            elif step_type == 'nested_exercise':
                self.create_nested_exercise_row(exercise_frame, step_info)
                exercise_count += 1
            elif step_type == 'nested_rest':
                self.create_nested_rest_row(exercise_frame, step_info)
                rest_count += 1
            elif step_info.get('is_rest') or step_info.get('step_type') == 'rest':
                self.create_rest_row(exercise_frame, step_info)
                rest_count += 1
            elif step_info.get('step_type') == 'warmup':
                self.create_warmup_row(exercise_frame, step_info)
                exercise_count += 1
            else:
                self.create_exercise_row(exercise_frame, step_info, exercise_count, sport)
                exercise_count += 1
                total_sets += step_info.get('sets', 1)

        # Footer stats
        footer = Frame(watch_frame, bg='#000')
        footer.pack(fill=X, pady=(15, 5))

        # Build stats parts
        stats_parts = [f"{len(exercises)} steps"]
        if exercise_count > 0:
            stats_parts.append(f"{exercise_count} exercises")
        if total_sets > exercise_count:
            stats_parts.append(f"{total_sets} total sets")

        Label(footer, text=" • ".join(stats_parts), font=('SF Pro Text', 11),
              bg='#000', fg='#666').pack()

        # Legend matching app style with icons
        legend_frame = Frame(main, bg='#1a1a1a')
        legend_frame.pack(fill=X, pady=(5, 0))

        legend_items = Frame(legend_frame, bg='#1a1a1a')
        legend_items.pack(fill=X)

        # App-style legend with icons
        self.create_legend_item(legend_items, "⊙", "Warmup", "#eab308")
        self.create_legend_item(legend_items, "※", "Warm-Up Set", "#f97316")
        self.create_legend_item(legend_items, "※", "Exercise", "#fff")
        self.create_legend_item(legend_items, "↷", "Rest", "#9ca3af")
        self.create_legend_item(legend_items, "↻", "Repeat", "#3b82f6")

        # Close button
        Button(main, text="Close", font=('SF Pro Text', 12),
               command=on_close, bg='#333', fg='#fff',
               padx=20, pady=8, relief=FLAT, cursor='hand2').pack(pady=(10, 0))

    def process_steps_for_preview(self, steps):
        """Process flat steps list to detect repeat structures for hierarchical display"""
        processed = []
        i = 0

        while i < len(steps):
            step = steps[i]

            # Check if this step is followed by rest + repeat pattern
            if i + 2 < len(steps):
                next_step = steps[i + 1]
                after_next = steps[i + 2]

                # Pattern: exercise -> rest -> repeat
                is_rest = next_step.get('is_rest') or next_step.get('step_type') == 'rest'
                is_repeat = after_next.get('is_repeat', False)

                if is_rest and is_repeat:
                    repeat_count = after_next.get('repeat_count', 0) + 1

                    # Create repeat header
                    processed.append({
                        'display_type': 'repeat_header',
                        'repeat_count': repeat_count,
                        'text': f"{repeat_count} Sets"
                    })

                    # Add the exercise as nested
                    nested_step = step.copy()
                    nested_step['display_type'] = 'nested_exercise'
                    processed.append(nested_step)

                    # Add the rest as nested
                    nested_rest = next_step.copy()
                    nested_rest['display_type'] = 'nested_rest'
                    processed.append(nested_rest)

                    # Skip the repeat marker
                    i += 3
                    continue

            # Check if step itself is a repeat marker (skip it)
            if step.get('is_repeat', False):
                i += 1
                continue

            # Regular step - add with display info
            step_copy = step.copy()
            step_copy['display_type'] = 'regular'
            processed.append(step_copy)
            i += 1

        return processed

    def create_repeat_header(self, parent, step_info):
        """Create a repeat/sets header row (green background like web app)"""
        row = Frame(parent, bg='#166534', padx=10, pady=8)  # Dark green to match web rgba(34, 197, 94, 0.2)
        row.pack(fill=X, pady=(8, 2), padx=2)

        Label(row, text=f"↻  {step_info.get('text', 'Sets')}",
              font=('SF Pro Text', 12, 'bold'),
              bg='#166534', fg='#4ade80').pack(anchor='w')  # Green text like web #4ade80

    def create_nested_exercise_row(self, parent, exercise):
        """Create an exercise row nested within a repeat block"""
        # Container with left border to show nesting (blue for regular, orange for warmup)
        is_warmup_set = exercise.get('is_warmup_set', False)
        border_color = '#f97316' if is_warmup_set else '#3b82f6'  # Orange for warmup, blue for regular
        bg_color = '#1a1520' if is_warmup_set else '#111827'  # Subtle background tint

        row = Frame(parent, bg=bg_color)
        row.pack(fill=X, pady=1, padx=2)

        # Left border indicator
        border = Frame(row, bg=border_color, width=4)
        border.pack(side=LEFT, fill=Y)

        # Content
        content = Frame(row, bg=bg_color, padx=10, pady=8)
        content.pack(side=LEFT, fill=BOTH, expand=True)

        # Exercise name with icon
        name = exercise.get('name', 'Exercise')
        text_color = '#fbbf24' if is_warmup_set else '#93c5fd'  # Matching web colors
        suffix = " (Warm-Up)" if is_warmup_set else ""

        Label(content, text=f"※  {name}{suffix}", font=('SF Pro Text', 11, 'bold'),
              bg=bg_color, fg=text_color, anchor='w', wraplength=420).pack(fill=X)

        # Badges row
        badges = Frame(content, bg=bg_color)
        badges.pack(fill=X, pady=(4, 0))

        # Reps badge (green for reps)
        if exercise.get('reps'):
            self.create_badge(badges, f"{exercise['reps']} reps", "#22c55e")

        # Duration badge (blue)
        if exercise.get('duration'):
            duration_str = self.format_duration(exercise['duration'])
            self.create_badge(badges, duration_str, "#3b82f6")
        elif exercise.get('duration_type') == 'open':
            self.create_badge(badges, "Lap Button", "#6b7280")

        # Category badge (gray)
        category = exercise.get('category', '')
        if category:
            try:
                cat_id = int(category)
                cat_name = EXERCISE_CATEGORY_NAMES.get(cat_id, '')
                if cat_name and cat_name.lower() not in name.lower():
                    Label(badges, text=cat_name, font=('SF Pro Text', 9),
                          bg='#374151', fg='#d1d5db', padx=6, pady=2).pack(side=LEFT, padx=(0, 5))
            except (ValueError, TypeError):
                if category.lower() not in name.lower():
                    Label(badges, text=category.replace('_', ' ').title(), font=('SF Pro Text', 9),
                          bg='#374151', fg='#d1d5db', padx=6, pady=2).pack(side=LEFT, padx=(0, 5))

    def create_nested_rest_row(self, parent, rest_info):
        """Create a rest row nested within a repeat block"""
        row = Frame(parent, bg='#111')
        row.pack(fill=X, pady=1, padx=2)

        # Left border indicator (gray for rest)
        border = Frame(row, bg='#6b7280', width=4)
        border.pack(side=LEFT, fill=Y)

        # Content
        content = Frame(row, bg='#111', padx=10, pady=6)
        content.pack(side=LEFT, fill=BOTH, expand=True)

        # Rest label with icon
        Label(content, text="↷  Rest", font=('SF Pro Text', 11),
              bg='#111', fg='#9ca3af', anchor='w').pack(side=LEFT)

        # Duration badge
        duration_type = rest_info.get('duration_type', '')
        rest_seconds = rest_info.get('rest_seconds', rest_info.get('duration', 0))

        if duration_type in ('open', 'lap_button') or rest_seconds <= 0:
            self.create_badge(content, "Lap Button", "#6b7280")
        else:
            self.create_badge(content, f"{int(rest_seconds)}s rest", "#6b7280")

    def create_rest_row(self, parent, rest_info):
        """Create a standalone rest row"""
        row = Frame(parent, bg='#1f2937', padx=10, pady=8)
        row.pack(fill=X, pady=2, padx=2)

        # Rest label with icon
        Label(row, text="↷  Rest", font=('SF Pro Text', 11, 'bold'),
              bg='#1f2937', fg='#9ca3af', anchor='w').pack(side=LEFT)

        # Duration badge
        duration_type = rest_info.get('duration_type', '')
        rest_seconds = rest_info.get('rest_seconds', rest_info.get('duration', 0))

        if duration_type in ('open', 'lap_button') or rest_seconds <= 0:
            self.create_badge(row, "Lap Button", "#f97316")
        else:
            self.create_badge(row, f"{int(rest_seconds)}s", "#f97316")

    def create_warmup_row(self, parent, warmup_info):
        """Create a warmup row with timer icon"""
        row = Frame(parent, bg='#1c1917', highlightbackground='#eab308',
                   highlightthickness=1, padx=10, pady=8)
        row.pack(fill=X, pady=2, padx=2)

        # Content
        content = Frame(row, bg='#1c1917')
        content.pack(fill=X)

        # Warmup label with timer icon (yellow/gold)
        name = warmup_info.get('name', 'Warmup')
        Label(content, text=f"⊙  {name}", font=('SF Pro Text', 11, 'bold'),
              bg='#1c1917', fg='#eab308', anchor='w', wraplength=420).pack(fill=X)

        # Duration badge
        badges = Frame(content, bg='#1c1917')
        badges.pack(fill=X, pady=(4, 0))

        duration = warmup_info.get('duration', 0)
        duration_type = warmup_info.get('duration_type', '')

        if duration > 0:
            duration_str = self.format_duration(duration)
            self.create_badge(badges, duration_str, "#3b82f6")
        elif duration_type in ('open', 5):  # 5 is FIT SDK OPEN
            self.create_badge(badges, "Press Lap", "#6b7280")
    
    def show_fit_preview_multi(self, filepaths):
        """Show multiple FIT files in a list summary view with single window navigation"""
        # Create or reuse preview window
        preview = Toplevel(self.root)
        preview.title(f"Workout Preview - {len(filepaths)} files")
        preview.geometry("500x600")
        preview.configure(bg='#1a1a1a')
        preview.transient(self.root)
        
        # Store filepaths for back navigation
        self._preview_window = preview
        self._preview_filepaths = filepaths
        
        # Content frame that can be cleared/rebuilt
        self._preview_content = Frame(preview, bg='#1a1a1a')
        self._preview_content.pack(fill=BOTH, expand=True)
        
        # Unbind mousewheel on close
        def on_close():
            try:
                preview.unbind_all("<MouseWheel>")
            except:
                pass
            preview.destroy()
        preview.protocol("WM_DELETE_WINDOW", on_close)
        
        # Build the list view
        self._build_list_view()
    
    def _build_list_view(self):
        """Build the workout list view"""
        # Clear content
        for widget in self._preview_content.winfo_children():
            widget.destroy()
        
        filepaths = self._preview_filepaths
        preview = self._preview_window
        content = self._preview_content
        
        preview.title(f"Workout Preview - {len(filepaths)} files")
        
        # Header
        header = Frame(content, bg='#1a1a1a')
        header.pack(fill=X, padx=15, pady=(15, 10))
        Label(header, text=f"📋 {len(filepaths)} Workouts", font=('SF Pro Display', 18, 'bold'),
              bg='#1a1a1a', fg='#fff').pack(anchor='w')
        
        # Scrollable list
        canvas = Canvas(content, bg='#1a1a1a', highlightthickness=0)
        scrollbar = Scrollbar(content, orient=VERTICAL, command=canvas.yview)
        list_frame = Frame(canvas, bg='#1a1a1a')
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True, padx=(15, 0))
        
        canvas_window = canvas.create_window((0, 0), window=list_frame, anchor='nw')
        
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
            canvas.itemconfig(canvas_window, width=event.width)
        
        list_frame.bind('<Configure>', configure_canvas)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))
        
        # Mouse wheel scrolling - scoped to this canvas
        def on_mousewheel(event):
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # Parse and display each file as a summary row
        for filepath in filepaths:
            workout_data = self.parse_fit_file(filepath)
            if not workout_data:
                continue

            # Card for each workout
            card = Frame(list_frame, bg='#222', highlightbackground='#333', highlightthickness=1)
            card.pack(fill=X, pady=4, padx=(0, 15))

            card_content = Frame(card, bg='#222', padx=12, pady=10)
            card_content.pack(fill=X)

            # Top row: name + sport badge
            top_row = Frame(card_content, bg='#222')
            top_row.pack(fill=X)

            name = workout_data.get('name', os.path.basename(filepath))
            Label(top_row, text=name, font=('SF Pro Text', 13, 'bold'),
                  bg='#222', fg='#fff').pack(side=LEFT)

            sport = workout_data.get('sport')
            sub_sport = workout_data.get('sub_sport')
            if sport:
                sport_display = get_sport_display(sport, sub_sport)
                sport_color = get_sport_color(sport, sub_sport)
                Label(top_row, text=f" {sport_display} ", font=('SF Pro Text', 9, 'bold'),
                      bg=sport_color, fg='#fff').pack(side=RIGHT)
            
            # Stats row
            stats_row = Frame(card_content, bg='#222')
            stats_row.pack(fill=X, pady=(6, 0))
            
            stats = []
            exercises = workout_data.get('steps', [])
            stats.append(f"{len(exercises)} steps")
            
            total_duration = sum(ex.get('duration', 0) for ex in exercises)
            if total_duration > 0:
                stats.append(f"⏱ {self.format_duration(total_duration)}")
            
            total_sets = sum(ex.get('sets', 1) for ex in exercises)
            if total_sets > len(exercises):
                stats.append(f"{total_sets} sets")
            
            if workout_data.get('created'):
                created = workout_data['created'].split(' ')[0]
                stats.append(f"📅 {created}")
            
            Label(stats_row, text="  •  ".join(stats), font=('SF Pro Text', 10),
                  bg='#222', fg='#888').pack(side=LEFT)
            
            # Preview button - navigate within same window
            btn = Label(stats_row, text="👁", font=('SF Pro Text', 14), 
                       bg='#222', fg='#007AFF', cursor='hand2')
            btn.pack(side=RIGHT)
            btn.bind('<Button-1>', lambda e, fp=filepath: self._show_detail_view(fp))
        
        # Bottom bar
        bottom = Frame(content, bg='#1a1a1a')
        bottom.pack(fill=X, padx=15, pady=15)
        
        Button(bottom, text="Close", font=('SF Pro Text', 12),
               command=self._preview_window.destroy, bg='#333', fg='#fff',
               padx=20, pady=8, relief=FLAT, cursor='hand2').pack(side=RIGHT)
    
    def _show_detail_view(self, filepath):
        """Show detailed workout view with back button"""
        workout_data = self.parse_fit_file(filepath)
        if not workout_data:
            return
        
        # Clear content
        for widget in self._preview_content.winfo_children():
            widget.destroy()
        
        preview = self._preview_window
        content = self._preview_content
        
        preview.title(f"Workout Preview - {workout_data.get('name', 'Workout')}")
        
        # Back button header
        header = Frame(content, bg='#1a1a1a')
        header.pack(fill=X, padx=15, pady=(10, 5))
        
        back_btn = Label(header, text="← Back", font=('SF Pro Text', 12),
                        bg='#1a1a1a', fg='#007AFF', cursor='hand2')
        back_btn.pack(side=LEFT)
        back_btn.bind('<Button-1>', lambda e: self._build_list_view())
        
        # Main container
        main = Frame(content, bg='#1a1a1a', padx=20, pady=10)
        main.pack(fill=BOTH, expand=True)
        
        # Watch face simulation
        watch_frame = Frame(main, bg='#000', highlightbackground='#333', 
                           highlightthickness=2, padx=15, pady=15)
        watch_frame.pack(fill=BOTH, expand=True, pady=(0, 15))
        
        # Workout title
        title = workout_data.get('name', 'Workout')
        Label(watch_frame, text=title, font=('SF Pro Display', 16, 'bold'),
              bg='#000', fg='#fff').pack(pady=(5, 5))
        
        # Sport type badge
        sport = workout_data.get('sport')
        sub_sport = workout_data.get('sub_sport')
        if sport:
            sport_display = get_sport_display(sport, sub_sport)
            sport_color = get_sport_color(sport, sub_sport)

            sport_badge = Label(watch_frame, text=f"  {sport_display}  ",
                               font=('SF Pro Text', 10, 'bold'),
                               bg=sport_color, fg='#fff')
            sport_badge.pack(pady=(0, 5))

        # Metadata row
        meta_frame = Frame(watch_frame, bg='#000')
        meta_frame.pack(fill=X, pady=(0, 10))
        
        meta_parts = []
        if workout_data.get('source'):
            meta_parts.append(f"📱 {workout_data['source']}")
        if workout_data.get('created'):
            created = workout_data['created'].split(' ')[0] if ' ' in workout_data['created'] else workout_data['created']
            meta_parts.append(f"📅 {created}")
        
        total_duration = sum(ex.get('duration', 0) for ex in workout_data.get('steps', []))
        if total_duration > 0:
            meta_parts.append(f"⏱ {self.format_duration(total_duration)}")
        
        if meta_parts:
            Label(meta_frame, text="  •  ".join(meta_parts), font=('SF Pro Text', 9),
                  bg='#000', fg='#666').pack()
        
        # Scrollable exercise list
        canvas = Canvas(watch_frame, bg='#000', highlightthickness=0, height=300)
        scrollbar = Scrollbar(watch_frame, orient=VERTICAL, command=canvas.yview)
        exercise_frame = Frame(canvas, bg='#000')
        
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=RIGHT, fill=Y)
        canvas.pack(side=LEFT, fill=BOTH, expand=True)
        
        canvas_window = canvas.create_window((0, 0), window=exercise_frame, anchor='nw')
        
        def configure_canvas(event):
            canvas.configure(scrollregion=canvas.bbox('all'))
            canvas.itemconfig(canvas_window, width=event.width)
        
        exercise_frame.bind('<Configure>', configure_canvas)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))
        
        # Mouse wheel scrolling
        def on_mousewheel(event):
            if canvas.winfo_exists():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel)
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", on_mousewheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        
        # Display exercises
        exercises = workout_data.get('steps', [])
        total_sets = 0
        rest_count = 0

        for i, exercise in enumerate(exercises):
            self.create_exercise_row(exercise_frame, exercise, i, sport)
            total_sets += exercise.get('sets', 1)
            if exercise.get('is_rest') or exercise.get('step_type') == 'rest':
                rest_count += 1

        # Footer stats
        footer = Frame(watch_frame, bg='#000')
        footer.pack(fill=X, pady=(15, 5))

        # Build stats parts
        stats_parts = [f"{len(exercises)} steps"]
        if total_sets > len(exercises):
            stats_parts.append(f"{total_sets} total sets")
        if rest_count > 0:
            stats_parts.append(f"{rest_count} rest")

        Label(footer, text=" • ".join(stats_parts), font=('SF Pro Text', 11),
              bg='#000', fg='#666').pack()

    def create_exercise_row(self, parent, exercise, index, sport=None):
        """Create a standalone exercise row (not nested in repeat)"""
        bg_color = '#111'
        name = exercise.get('name', f'Exercise {index + 1}')
        duration_type = exercise.get('duration_type', '')

        row = Frame(parent, bg=bg_color, padx=10, pady=8)
        row.pack(fill=X, pady=2, padx=2)

        # Exercise name with icon
        Label(row, text=f"※  {name}", font=('SF Pro Text', 11, 'bold'),
              bg=bg_color, fg='#fff', anchor='w', wraplength=420).pack(fill=X)

        # Badges row
        badges = Frame(row, bg=bg_color)
        badges.pack(fill=X, pady=(4, 0))

        # Reps badge (green)
        if exercise.get('reps'):
            self.create_badge(badges, f"{exercise['reps']} reps", "#22c55e")

        # Duration badge (blue)
        if exercise.get('duration'):
            duration_str = self.format_duration(exercise['duration'])
            self.create_badge(badges, duration_str, "#3b82f6")
        elif duration_type == 'open':
            self.create_badge(badges, "Lap Button", "#6b7280")

        # Sets badge (green, only if > 1)
        sets = exercise.get('sets', 1)
        if sets > 1:
            self.create_badge(badges, f"{sets} sets", "#22c55e")

        # Category badge (gray)
        category = exercise.get('category', '')
        if category:
            try:
                cat_id = int(category)
                cat_name = EXERCISE_CATEGORY_NAMES.get(cat_id, '')
                if cat_name and cat_name.lower() not in name.lower():
                    Label(badges, text=cat_name, font=('SF Pro Text', 9),
                          bg='#374151', fg='#d1d5db', padx=6, pady=2).pack(side=LEFT, padx=(0, 5))
            except (ValueError, TypeError):
                if category.lower() not in name.lower():
                    Label(badges, text=category.replace('_', ' ').title(), font=('SF Pro Text', 9),
                          bg='#374151', fg='#d1d5db', padx=6, pady=2).pack(side=LEFT, padx=(0, 5))

    def create_badge(self, parent, text, color):
        """Create a colored badge"""
        badge = Label(parent, text=text, font=('SF Pro Text', 10, 'bold'),
                     bg=color, fg='#fff', padx=8, pady=2)
        badge.pack(side=LEFT, padx=(0, 5))

    def create_legend_item(self, parent, icon, text, color):
        """Create a legend item with icon matching app style"""
        item = Frame(parent, bg='#1a1a1a')
        item.pack(side=LEFT, padx=(0, 12))

        Label(item, text=icon, font=('SF Pro Text', 10), bg='#1a1a1a', fg=color).pack(side=LEFT)
        Label(item, text=f" {text}", font=('SF Pro Text', 9), bg='#1a1a1a', fg='#888').pack(side=LEFT)

    def create_legend_badge(self, parent, text, color):
        """Create a legend badge (legacy)"""
        item = Frame(parent, bg='#1a1a1a')
        item.pack(side=LEFT, padx=(0, 15))

        Label(item, text="●", font=('SF Pro Text', 10), bg='#1a1a1a', fg=color).pack(side=LEFT)
        Label(item, text=text, font=('SF Pro Text', 10), bg='#1a1a1a', fg='#888').pack(side=LEFT, padx=(3, 0))
    
    def format_duration(self, seconds):
        """Format duration in seconds to human readable string"""
        if seconds < 60:
            return f"{int(seconds)}s"
        elif seconds < 3600:
            mins = int(seconds // 60)
            secs = int(seconds % 60)
            return f"{mins}:{secs:02d}" if secs else f"{mins}min"
        else:
            hours = int(seconds // 3600)
            mins = int((seconds % 3600) // 60)
            return f"{hours}h {mins}m" if mins else f"{hours}h"
    
    def format_distance(self, meters):
        """Format distance in meters to human readable string"""
        if meters < 1000:
            return f"{int(meters)}m"
        else:
            km = meters / 1000
            return f"{km:.1f}km"

    def validate_fit_file(self, filepath):
        """Validate FIT file for issues that may prevent it from working on Garmin watches.
        Returns dict with 'valid' boolean and 'issues' list."""
        # Try fitfiletool's validator first
        if FITFILETOOL_AVAILABLE:
            result = fitfiletool_validate_fit_file(filepath)
            # Add empty invalid_categories for compatibility
            if 'invalid_categories' not in result:
                result['invalid_categories'] = []
            return result

        # Fall back to local validation
        if not FITPARSE_AVAILABLE:
            return {'valid': True, 'issues': [], 'warnings': [], 'invalid_categories': []}

        try:
            fitfile = FitFile(filepath)
            issues = []
            invalid_categories = []

            # Valid FIT SDK exercise categories are 0-32
            VALID_CATEGORIES = set(range(33))

            for record in fitfile.get_messages('workout_step'):
                for field in record.fields:
                    if field.name == 'exercise_category' and field.value is not None:
                        # Check if it's a raw number (invalid) vs named category
                        if isinstance(field.value, int) and field.value not in VALID_CATEGORIES:
                            invalid_categories.append(field.value)

            if invalid_categories:
                unique_invalid = list(set(invalid_categories))
                issues.append(f"Invalid exercise categories found: {unique_invalid}")
                issues.append("These may cause the workout to not appear on your Garmin watch.")

            return {
                'valid': len(issues) == 0,
                'issues': issues,
                'warnings': [],
                'invalid_categories': list(set(invalid_categories))
            }
        except Exception as e:
            return {'valid': False, 'issues': [f"Error validating file: {str(e)}"], 'warnings': [], 'invalid_categories': []}

    def repair_fit_file(self, filepath, workout_data):
        """Repair a FIT file by regenerating it with valid exercise categories.
        Uses amakaflow-fitfiletool for proper FIT generation."""
        if not FITFILETOOL_AVAILABLE:
            return None, "amakaflow-fitfiletool not available"

        try:
            lookup = GarminExerciseLookup()

            # Convert parsed workout data to fitfiletool format
            exercises = []
            for step in workout_data.get('steps', []):
                name = step.get('name', 'Exercise')

                # Determine reps/duration
                if step.get('reps'):
                    reps = step['reps']
                elif step.get('distance'):
                    # Convert distance (meters) to string format
                    dist = step['distance']
                    if dist >= 1000:
                        reps = f"{dist/1000:.1f}km"
                    else:
                        reps = f"{int(dist)}m"
                elif step.get('duration'):
                    reps = f"{int(step['duration'])}s"
                else:
                    reps = 10  # Default

                exercises.append({
                    'name': name,
                    'reps': reps,
                    'sets': step.get('sets', 1)
                })

            workout = {
                'title': workout_data.get('name', 'Repaired Workout'),
                'blocks': [{'exercises': exercises, 'rest_between_sec': 0}]
            }

            # Generate new FIT file
            fit_bytes = build_fit_workout(workout, use_lap_button=False)

            # Save to new file
            base, ext = os.path.splitext(filepath)
            new_filepath = f"{base}_repaired{ext}"

            with open(new_filepath, 'wb') as f:
                f.write(fit_bytes)

            return new_filepath, None
        except Exception as e:
            return None, f"Error repairing file: {str(e)}"

    def parse_fit_file(self, filepath):
        """Parse a FIT file and extract workout data"""
        # Try fitfiletool's parser first (uses fitparse internally)
        if FITFILETOOL_AVAILABLE:
            result = fitfiletool_parse_fit_file(filepath)
            if result:
                return result
        # Fall back to local fitparse implementation
        if FITPARSE_AVAILABLE:
            return self.parse_fit_with_fitparse(filepath)
        # Last resort: basic binary parsing
        return self.parse_fit_basic(filepath)
    
    def parse_fit_with_fitparse(self, filepath):
        """Parse FIT file using fitparse library"""
        try:
            fitfile = FitFile(filepath)
            
            workout_data = {
                'name': 'Workout',
                'sport': None,
                'steps': [],
                'created': None,
                'source': None
            }
            
            # Get file metadata
            for record in fitfile.get_messages('file_id'):
                for field in record.fields:
                    if field.name == 'time_created' and field.value:
                        workout_data['created'] = str(field.value)
                    elif field.name == 'manufacturer' and field.value:
                        workout_data['manufacturer'] = str(field.value)
                    elif field.name == 'garmin_product' and field.value:
                        workout_data['source'] = str(field.value).replace('_', ' ').title()
            
            # First pass: collect exercise titles for lookup (strength workouts)
            exercise_titles = {}
            for record in fitfile.get_messages('exercise_title'):
                title_data = {}
                for field in record.fields:
                    if field.name == 'wkt_step_name':
                        title_data['name'] = field.value
                    elif field.name == 'exercise_category':
                        title_data['category'] = str(field.value) if field.value else None
                    elif field.name == 'exercise_name':
                        title_data['exercise_id'] = field.value
                
                if title_data.get('category') and title_data.get('name'):
                    key = (title_data.get('category'), title_data.get('exercise_id'))
                    exercise_titles[key] = title_data['name']
                    exercise_titles[title_data.get('category')] = title_data['name']
            
            # Get workout name and sport type
            for record in fitfile.get_messages('workout'):
                for field in record.fields:
                    if field.name == 'wkt_name' and field.value:
                        workout_data['name'] = field.value
                    elif field.name == 'sport' and field.value:
                        workout_data['sport'] = str(field.value)
                    elif field.name == 'sub_sport' and field.value:
                        workout_data['sub_sport'] = str(field.value)
            
            # Second pass: get workout steps
            steps_raw = []
            for record in fitfile.get_messages('workout_step'):
                step = {'is_rest': False, 'is_repeat': False}
                for field in record.fields:
                    if field.name == 'wkt_step_name' and field.value:
                        step['name'] = field.value
                    elif field.name == 'exercise_category' and field.value:
                        step['category'] = str(field.value)
                    elif field.name == 'exercise_name':
                        step['exercise_id'] = field.value
                    elif field.name == 'duration_type':
                        dtype_str = str(field.value) if field.value else ''
                        step['duration_type'] = dtype_str
                        # FIT SDK: repeat types indicate this is a repeat step
                        # repeat_until_steps_cmplt=6, repeat_until_time=7, etc.
                        if 'repeat' in dtype_str.lower() or dtype_str in ('6', '7', '8', '9'):
                            step['is_repeat'] = True
                    elif field.name == 'duration_step' and field.value is not None:
                        # This is which step to repeat back to (for repeat steps)
                        step['duration_step'] = int(field.value)
                    elif field.name == 'duration_value' and field.value is not None:
                        # For repeat steps, this is the repeat count
                        if step.get('is_repeat'):
                            step['repeat_count'] = int(field.value)
                    elif field.name == 'duration_reps' and field.value:
                        step['reps'] = int(field.value)
                    elif field.name == 'duration_time' and field.value:
                        step['duration'] = float(field.value)
                    elif field.name == 'duration_distance' and field.value:
                        step['distance'] = float(field.value)
                    elif field.name == 'intensity':
                        intensity_raw = field.value
                        intensity = str(intensity_raw) if intensity_raw is not None else None
                        step['intensity'] = intensity
                        # FIT SDK intensity: 0=active, 1=rest, 2=warmup, 3=cooldown
                        # fitparse may return string or numeric
                        if intensity in ('rest', '1', 1):
                            step['is_rest'] = True
                        elif intensity in ('warmup', '2', 2):
                            step['is_warmup'] = True
                    elif field.name == 'repeat_steps' and field.value:
                        step['is_repeat'] = True
                        step['repeat_count'] = int(field.value)
                    elif field.name == 'exercise_weight' and field.value:
                        step['weight'] = float(field.value)
                    elif field.name == 'weight_display_unit':
                        step['weight_unit'] = str(field.value) if field.value else 'kg'
                    elif field.name == 'notes' and field.value:
                        step['notes'] = field.value
                    elif field.name == 'target_type' and field.value:
                        step['target_type'] = str(field.value)
                    elif field.name == 'target_value' and field.value:
                        step['target_value'] = field.value
                
                steps_raw.append(step)
            
            # Determine if this is a cardio workout (running, cycling, etc.) vs strength
            sport_lower = (workout_data.get('sport') or '').lower()
            sub_sport_lower = (workout_data.get('sub_sport') or '').lower()
            cardio_sports = ['running', 'cycling', 'swimming', 'walking', 'hiking', 'run', 'bike', 'swim', 'walk', 'hike', 'cardio', 'trail_running', 'treadmill']
            is_cardio = sport_lower in cardio_sports or sub_sport_lower in cardio_sports or 'run' in sport_lower or 'run' in sub_sport_lower
            
            # Third pass: process steps
            # Keep rest and repeat steps as separate entries for grouped display
            exercises = []
            i = 0
            while i < len(steps_raw):
                step = steps_raw[i]

                # Handle repeat markers - keep as separate step for grouped display
                if step.get('is_repeat'):
                    repeat_step = {
                        'is_repeat': True,
                        'repeat_count': step.get('repeat_count', 0),
                        'name': f"{step.get('repeat_count', 0) + 1} Sets",
                        'step_type': 'repeat'
                    }
                    # Also update previous exercise's sets for badge display
                    if exercises and step.get('repeat_count'):
                        for ex in reversed(exercises):
                            if not ex.get('is_rest') and not ex.get('is_repeat'):
                                ex['sets'] = step['repeat_count'] + 1
                                break
                    exercises.append(repeat_step)
                    i += 1
                    continue

                # Handle rest steps - keep as separate entries for grouped display
                if step.get('is_rest'):
                    rest_step = {
                        'is_rest': True,
                        'step_type': 'rest',
                        'name': 'Rest',
                        'duration_type': step.get('duration_type', 'time'),
                        'rest_seconds': step.get('duration', 0),
                        'duration': step.get('duration', 0),
                        'category': step.get('category')
                    }
                    if step.get('duration_type') in ('open', 'repeat_until_steps_cmplt'):
                        rest_step['duration_type'] = 'open'
                    exercises.append(rest_step)
                    i += 1
                    continue

                # Handle warmup steps - keep as separate entries
                if step.get('is_warmup'):
                    warmup_step = {
                        'step_type': 'warmup',
                        'name': step.get('name') or 'Warm-Up',
                        'duration_type': step.get('duration_type', 'time'),
                        'duration': step.get('duration', 0),
                        'category': step.get('category')
                    }
                    if step.get('duration_type') in ('open', 'repeat_until_steps_cmplt'):
                        warmup_step['duration_type'] = 'open'
                    exercises.append(warmup_step)
                    i += 1
                    continue

                exercise = {}
                cat = step.get('category')
                ex_id = step.get('exercise_id')
                intensity = step.get('intensity')
                notes = step.get('notes')
                
                # Build step name based on workout type
                if is_cardio:
                    # For cardio workouts, use intensity + notes
                    sport_name = workout_data.get('sport', 'exercise').title()

                    # FIT SDK intensity: 0=active, 1=rest, 2=warmup, 3=cooldown
                    if intensity in ('warmup', '2', 2):
                        exercise['name'] = 'Warm Up'
                        exercise['step_type'] = 'warmup'
                    elif intensity in ('cooldown', '3', 3):
                        exercise['name'] = 'Cool Down'
                        exercise['step_type'] = 'cooldown'
                    elif intensity in ('rest', '1', 1):
                        exercise['name'] = 'Recovery'
                        exercise['step_type'] = 'rest'
                    elif intensity in ('active', '0', 0):
                        exercise['name'] = notes if notes else sport_name
                        exercise['step_type'] = 'active'
                    else:
                        exercise['name'] = notes if notes else sport_name
                        exercise['step_type'] = 'active'
                    
                    # Add notes as subtitle if we used intensity for name
                    if notes and exercise['name'] != notes:
                        exercise['notes'] = notes
                else:
                    # For strength workouts, use exercise title lookup
                    if step.get('name'):
                        exercise['name'] = step['name']
                    elif cat and (cat, ex_id) in exercise_titles:
                        exercise['name'] = exercise_titles[(cat, ex_id)]
                    elif cat and cat in exercise_titles:
                        exercise['name'] = exercise_titles[cat]
                    elif cat:
                        exercise['name'] = cat.replace('_', ' ').title()
                    else:
                        exercise['name'] = 'Exercise'
                
                # Copy over exercise data
                if step.get('reps'):
                    exercise['reps'] = step['reps']
                if step.get('duration'):
                    exercise['duration'] = step['duration']
                if step.get('distance'):
                    exercise['distance'] = step['distance']
                if step.get('weight'):
                    weight = step['weight']
                    unit = step.get('weight_unit', 'kg')
                    if unit == 'pound':
                        exercise['weight'] = f"{weight:.0f} lbs"
                    else:
                        exercise['weight'] = f"{weight:.1f} kg"
                if step.get('notes') and is_cardio:
                    exercise['zone'] = step['notes']
                
                exercise['sets'] = 1
                exercise['type'] = cat.replace('_', ' ').title() if cat else ''
                exercise['category'] = cat  # Keep original category for display lookup

                exercises.append(exercise)
                i += 1
            
            workout_data['steps'] = exercises
            return workout_data if exercises else None
            
        except Exception as e:
            # Silently fall back to basic parsing
            return self.parse_fit_basic(filepath)
    
    def parse_fit_basic(self, filepath):
        """Basic FIT file parsing without fitparse library"""
        try:
            with open(filepath, 'rb') as f:
                data = f.read()
            
            # Check FIT header
            if len(data) < 14:
                return None
            
            header_size = data[0]
            if header_size < 12:
                return None
            
            # Check for ".FIT" signature
            if data[8:12] != b'.FIT':
                return None
            
            # Basic parsing - look for workout name in data
            workout_data = {
                'name': 'Workout',
                'steps': []
            }
            
            # Try to find readable strings that might be workout/exercise names
            # This is a simplified approach
            text_start = None
            for i in range(header_size, len(data) - 4):
                # Look for printable ASCII sequences
                if 32 <= data[i] <= 126:
                    if text_start is None:
                        text_start = i
                else:
                    if text_start is not None and i - text_start >= 4:
                        text = data[text_start:i].decode('ascii', errors='ignore')
                        # Filter for likely workout/exercise names
                        if len(text) >= 4 and not text.startswith(('.', '/', '\\')):
                            if any(kw in text.lower() for kw in ['workout', 'exercise', 'run', 'bike', 'swim', 'strength']):
                                if not workout_data['steps']:
                                    workout_data['name'] = text
                            elif len(text) < 30:
                                workout_data['steps'].append({'name': text, 'type': 'exercise'})
                    text_start = None
            
            # If we couldn't parse steps, create a placeholder
            if not workout_data['steps']:
                workout_data['steps'] = [
                    {'name': 'Workout content', 'type': 'workout'},
                    {'name': '(Install fitparse for detailed view)', 'type': 'info'}
                ]
            
            return workout_data
            
        except Exception as e:
            # Silently return None for invalid files
            return None


def main():
    # Use TkinterDnD if available, otherwise fall back to regular Tk
    global DND_AVAILABLE
    root = None
    
    if DND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except RuntimeError:
            DND_AVAILABLE = False
            root = Tk()
    else:
        root = Tk()
    
    # Center on screen
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 580) // 2
    y = (root.winfo_screenheight() - 620) // 2
    root.geometry(f"+{x}+{y}")
    
    # Mac-specific styling
    root.tk.call('tk', 'scaling', 2.0)  # Retina support
    
    app = GarminUploaderMac(root)
    root.mainloop()


if __name__ == "__main__":
    main()
