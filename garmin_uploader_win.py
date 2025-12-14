#!/usr/bin/env python3
"""Garmin Workout Uploader for Windows"""

import os, sys, shutil, subprocess, re, threading, time, ctypes, json
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from urllib.request import urlopen
from urllib.error import URLError

# Windows subprocess flag to prevent console windows from appearing
CREATE_NO_WINDOW = 0x08000000

try:
    from version import __version__, __app_name__, __github_repo__
except ImportError:
    __version__ = "1.0.0"
    __app_name__ = "Garmin Workout Uploader"
    __github_repo__ = "supergeri/garmin-usb-mac-app"

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
    pass

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

# Try to import win32com for MTP file transfer
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

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
            ssl_context = ssl.create_default_context()
        try:
            url = f"https://api.github.com/repos/{__github_repo__}/releases/latest"
            with urlopen(url, timeout=10, context=ssl_context) as response:
                data = json.loads(response.read().decode())
                latest_version = data['tag_name'].lstrip('v')
                download_url = None

                # Find Windows installer in assets
                for asset in data.get('assets', []):
                    if asset['name'].endswith('.exe'):
                        download_url = asset['browser_download_url']
                        break

                return {
                    'available': UpdateChecker._compare_versions(latest_version, __version__),
                    'version': latest_version,
                    'url': download_url or data['html_url'],
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
            temp_file = os.path.join(tempfile.gettempdir(), 'GarminWorkoutUploaderSetup.exe')

            with urlopen(url, timeout=30, context=ssl_context) as response:
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

class GarminUploaderWin:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{__app_name__} v{__version__}")
        self.root.geometry("580x750")
        self.root.minsize(520, 650)
        self.root.resizable(True, True)
        self.root.update_idletasks()
        self.root.configure(bg='#f5f5f7')
        
        self.style = ttk.Style()
        self.style.configure("Title.TLabel", font=('Segoe UI', 18, 'bold'), background='#f5f5f7')
        self.style.configure("Subtitle.TLabel", font=('Segoe UI', 11), background='#f5f5f7', foreground='#666')
        
        self.home = Path.home()
        self.staging_folder = self.home / "GarminWorkouts"
        self.staging_folder.mkdir(exist_ok=True)
        
        self.selected_files = []
        self.close_ge_btn = None
        self._monitor_running = True
        self.garmin_drive = None
        self.garmin_newfiles = None
        self.is_mtp = False
        self.mtp_device_name = None

        # Track connected device for model-specific adjustments
        self.current_device = None

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.create_ui()
        self.root.update()

        # Check for updates in background
        threading.Thread(target=self._check_updates, daemon=True).start()
    
    def _on_close(self):
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
        update_frame.pack(fill=X, side=TOP)

        Label(update_frame, text=f"🎉 Update Available: v{update_info['version']}",
              font=('Segoe UI', 11, 'bold'), bg='#4CAF50', fg='white').pack(side=LEFT)

        Button(update_frame, text="Download Update", font=('Segoe UI', 10),
               bg='white', fg='#4CAF50', relief=FLAT, padx=12, pady=4,
               command=lambda: self._download_and_install(update_info)).pack(side=RIGHT, padx=(0, 5))

        Button(update_frame, text="View Release Notes", font=('Segoe UI', 10),
               bg='#45A049', fg='white', relief=FLAT, padx=12, pady=4,
               command=lambda: subprocess.run(['explorer', update_info['url']])).pack(side=RIGHT)

    def _download_and_install(self, update_info):
        """Download and install update"""
        response = messagebox.askyesno(
            "Download Update",
            f"Download and install version {update_info['version']}?\n\n"
            "The installer will open after download. The app will close automatically."
        )

        if response:
            # Show progress dialog
            progress_window = Toplevel(self.root)
            progress_window.title("Downloading Update")
            progress_window.geometry("400x100")
            progress_window.resizable(False, False)

            Label(progress_window, text="Downloading update...", font=('Segoe UI', 11)).pack(pady=10)

            progress_bar = ttk.Progressbar(progress_window, length=350, mode='determinate')
            progress_bar.pack(pady=10)

            def update_progress(pct):
                progress_bar['value'] = pct * 100
                progress_window.update()

            def do_download():
                installer_path = UpdateChecker.download_update(update_info['url'], update_progress)
                progress_window.destroy()

                if installer_path:
                    subprocess.Popen([installer_path])
                    self.root.quit()
                else:
                    messagebox.showerror("Download Failed", "Could not download the update. Please try again.")

            threading.Thread(target=do_download, daemon=True).start()
    
    def detect_garmin_device(self):
        import string
        for letter in string.ascii_uppercase:
            drive = f"{letter}:\\"
            garmin_path = os.path.join(drive, "GARMIN")
            try:
                if os.path.exists(garmin_path):
                    self.garmin_drive = drive
                    self.garmin_newfiles = os.path.join(garmin_path, "NewFiles")
                    self.is_mtp = False
                    return {'connected': True, 'name': f"Garmin ({letter}:)", 'mode': 'drive'}
            except:
                continue
        try:
            ps_cmd = 'Get-PnpDevice -Class WPD -Status OK | Select-Object -ExpandProperty FriendlyName'
            result = subprocess.run(['powershell', '-Command', ps_cmd], capture_output=True, text=True, timeout=10, creationflags=CREATE_NO_WINDOW)
            if result.returncode == 0 and result.stdout.strip():
                keywords = ['garmin', 'fenix', 'forerunner', 'venu', 'instinct', 'epix', 'edge', 'vivoactive']
                for line in result.stdout.strip().split('\n'):
                    line = line.strip()
                    if any(kw in line.lower() for kw in keywords):
                        self.is_mtp = True
                        self.mtp_device_name = line
                        return {'connected': True, 'name': line, 'mode': 'mtp'}
        except:
            pass
        self.is_mtp = False
        return None
    
    def check_garmin_express(self):
        try:
            result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq GarminExpress.exe'], capture_output=True, text=True, timeout=5, creationflags=CREATE_NO_WINDOW)
            return 'GarminExpress.exe' in result.stdout
        except:
            return False
    
    def kill_garmin_express(self):
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'GarminExpress.exe'], timeout=5, capture_output=True, creationflags=CREATE_NO_WINDOW)
        except:
            pass

    def transfer_mtp_files(self, files):
        """Transfer files to Garmin device via MTP using Windows Shell COM"""
        if not WIN32COM_AVAILABLE:
            return False, "pywin32 library not available"

        try:
            shell = win32com.client.Dispatch("Shell.Application")

            # Find the Garmin device in "This PC"
            # Namespace 17 = This PC (My Computer)
            this_pc = shell.Namespace(17)
            device_item = None

            # Search for Garmin device
            for item in this_pc.Items():
                if self.mtp_device_name and self.mtp_device_name.lower() in item.Name.lower():
                    device_item = item
                    break

            if not device_item:
                return False, "Could not find Garmin device in This PC"

            # Get device folder using GetFolder (works better for MTP)
            device_folder = device_item.GetFolder
            if not device_folder:
                return False, "Could not access device folder"

            # Look for Internal Storage or similar
            storage_folder = None
            for item in device_folder.Items():
                if 'storage' in item.Name.lower():
                    storage_folder = item.GetFolder
                    break

            # If no storage subfolder, use device folder directly
            if not storage_folder:
                storage_folder = device_folder

            # Navigate to GARMIN folder
            garmin_folder = None
            for item in storage_folder.Items():
                if item.Name.upper() == "GARMIN":
                    garmin_folder = item.GetFolder
                    break

            if not garmin_folder:
                return False, "Could not find GARMIN folder on device"

            # Try Workouts folder first (newer watches), then NewFiles (older watches)
            target_folder = None
            folder_used = None

            # Check for existing Workouts folder
            for item in garmin_folder.Items():
                if item.Name.upper() == "WORKOUTS":
                    target_folder = item.GetFolder
                    folder_used = "Workouts"
                    break

            # If no Workouts folder, try NewFiles
            if not target_folder:
                for item in garmin_folder.Items():
                    if item.Name.upper() == "NEWFILES":
                        target_folder = item.GetFolder
                        folder_used = "NewFiles"
                        break

            # If neither exists, create NewFiles as fallback
            if not target_folder:
                try:
                    garmin_folder.NewFolder("NewFiles")
                    time.sleep(1.0)  # Give MTP time to create folder
                    for item in garmin_folder.Items():
                        if item.Name.upper() == "NEWFILES":
                            target_folder = item.GetFolder
                            folder_used = "NewFiles"
                            break
                except:
                    pass

            if not target_folder:
                return False, "Could not access or create Workouts/NewFiles folder"

            # Copy files
            copied = 0
            for filepath in files:
                try:
                    filename = os.path.basename(filepath)

                    # Check if file already exists
                    file_exists = False
                    for item in target_folder.Items():
                        if item.Name == filename:
                            file_exists = True
                            break

                    # CopyHere with no flags to show progress (blocking)
                    # This ensures the copy completes before continuing
                    target_folder.CopyHere(filepath, 0)

                    # Wait and verify file appeared
                    max_wait = 10  # seconds
                    for i in range(max_wait * 2):
                        time.sleep(0.5)
                        found = False
                        for item in target_folder.Items():
                            if item.Name == filename:
                                found = True
                                break
                        if found and not file_exists:
                            break

                    copied += 1
                except Exception as e:
                    print(f"Error copying {filepath}: {e}")
                    continue

            if copied > 0:
                return True, f"{copied} file(s) transferred to GARMIN/{folder_used}"
            else:
                return False, "No files were copied"

        except Exception as e:
            return False, f"MTP transfer error: {str(e)}"

    def refresh_device_status(self):
        try:
            if not self.device_status.winfo_exists():
                return
        except:
            return
        device = self.detect_garmin_device()
        ge = self.check_garmin_express()

        # Store detected device for model-specific adjustments
        self.current_device = device

        if self.close_ge_btn:
            try:
                self.close_ge_btn.destroy()
            except:
                pass
            self.close_ge_btn = None
        try:
            if device:
                if ge:
                    self.device_status.config(text=f"⚠️ {device['name']}", fg='#FF9500')
                    self.device_detail.config(text="Close Garmin Express first")
                    self.close_ge_btn = Button(self.status_container, text="Close Garmin Express", font=('Segoe UI', 10), bg='#FF9500', fg='white', command=lambda: [self.kill_garmin_express(), self.root.after(1500, self.refresh_device_status)], relief=FLAT, padx=8, pady=3)
                    self.close_ge_btn.pack(anchor='w', pady=(6, 0))
                elif device.get('mode') == 'mtp':
                    self.device_status.config(text=f"✅ {device['name']}", fg='#28a745')
                    if WIN32COM_AVAILABLE:
                        self.device_detail.config(text="MTP mode - automatic transfer enabled")
                    else:
                        self.device_detail.config(text="MTP mode - install pywin32 for auto transfer")
                else:
                    self.device_status.config(text=f"✅ {device['name']}", fg='#28a745')
                    self.device_detail.config(text="Ready for direct transfer")
            else:
                self.device_status.config(text="❌ No device detected", fg='#dc3545')
                self.device_detail.config(text="Connect watch via USB")
        except:
            pass
    
    def start_monitor(self):
        def monitor():
            while self._monitor_running:
                try:
                    self.root.after(0, self.refresh_device_status)
                except:
                    break
                time.sleep(3)
        threading.Thread(target=monitor, daemon=True).start()
    
    def create_ui(self):
        main = Frame(self.root, bg='#f5f5f7', padx=20, pady=15)
        main.pack(fill=BOTH, expand=True)
        ttk.Label(main, text="Garmin Workout Uploader", style="Title.TLabel").pack()
        ttk.Label(main, text="Upload .FIT workouts to your Garmin watch", style="Subtitle.TLabel").pack(pady=(3, 15))
        
        # Step 1: Select Files
        c1 = Frame(main, bg='#fff', highlightbackground='#e0e0e0', highlightthickness=1, padx=12, pady=10)
        c1.pack(fill=X, pady=(0, 10))
        Label(c1, text="① Select Your Workout Files", font=('Segoe UI', 11, 'bold'), bg='#fff').pack(anchor='w')
        self.file_listbox = Listbox(c1, height=4, font=('Segoe UI', 10), selectbackground='#007AFF', relief=FLAT, highlightthickness=1, highlightbackground='#e0e0e0')
        self.file_listbox.pack(fill=X, pady=(8, 8))
        self.file_listbox.insert(END, "  Click 'Add Files' to select .FIT files")
        self.file_listbox.config(fg='#999')
        bf = Frame(c1, bg='#fff')
        bf.pack(fill=X)
        Button(bf, text="+ Add Files", font=('Segoe UI', 10), command=self.add_files, bg='#007AFF', fg='white', padx=12, pady=4, relief=FLAT, cursor='hand2').pack(side=LEFT)
        Button(bf, text="Clear", font=('Segoe UI', 10), command=self.clear_files, padx=10, pady=4, relief=FLAT).pack(side=LEFT, padx=(8, 0))
        Button(bf, text="Preview", font=('Segoe UI', 10), command=self.preview_file, padx=10, pady=4, relief=FLAT).pack(side=LEFT, padx=(8, 0))
        self.file_count = Label(bf, text="", font=('Segoe UI', 10), bg='#fff', fg='#666')
        self.file_count.pack(side=LEFT, padx=(12, 0))
        
        # Step 2: Prepare Transfer
        c2 = Frame(main, bg='#fff', highlightbackground='#e0e0e0', highlightthickness=1, padx=12, pady=10)
        c2.pack(fill=X, pady=(0, 10))
        Label(c2, text="② Prepare for Transfer", font=('Segoe UI', 11, 'bold'), bg='#fff').pack(anchor='w')
        Label(c2, text="Before transferring, make sure:\n✓ Your Garmin watch is connected via USB\n✓ Watch: Settings → System → USB Mode\n✓ Garmin Express is closed", font=('Segoe UI', 10), bg='#fff', justify=LEFT).pack(anchor='w', pady=(6, 8))
        self.transfer_btn = Button(c2, text="Transfer to Watch", font=('Segoe UI', 11, 'bold'), command=self.transfer, bg='#34C759', fg='white', padx=20, pady=6, relief=FLAT, cursor='hand2', state=DISABLED)
        self.transfer_btn.pack(pady=(0, 4))
        
        # Step 3: Device Status
        c3 = Frame(main, bg='#fff', highlightbackground='#e0e0e0', highlightthickness=1, padx=12, pady=10)
        c3.pack(fill=X, pady=(0, 10))
        Label(c3, text="③ Device Status", font=('Segoe UI', 11, 'bold'), bg='#fff').pack(anchor='w')
        sf = Frame(c3, bg='#f0f0f0', padx=10, pady=8)
        sf.pack(fill=X, pady=(8, 5))
        hr = Frame(sf, bg='#f0f0f0')
        hr.pack(fill=X)
        self.device_status = Label(hr, text="🔍 Checking for device...", font=('Segoe UI', 11, 'bold'), bg='#f0f0f0', fg='#666')
        self.device_status.pack(side=LEFT)
        Button(hr, text="↻ Refresh", font=('Segoe UI', 10), command=self.refresh_device_status, bg='white', fg='#007AFF', relief=SOLID, borderwidth=1, padx=10, pady=2, cursor='hand2').pack(side=RIGHT)
        self.status_container = Frame(sf, bg='#f0f0f0')
        self.status_container.pack(fill=X)
        self.device_detail = Label(self.status_container, text="Please wait...", font=('Segoe UI', 10), bg='#f0f0f0', fg='#666')
        self.device_detail.pack(anchor='w', pady=(4, 0))
        self.transfer_status = Label(c3, text="Select files in Step 1, then click 'Transfer to Watch'", font=('Segoe UI', 10), bg='#fff', fg='#666')
        self.transfer_status.pack(anchor='w', pady=(5, 0))
        
        self.root.after(500, self.refresh_device_status)
        self.root.after(1000, self.start_monitor)
    
    def add_files(self):
        files = filedialog.askopenfilenames(title="Select .FIT Files", filetypes=[("FIT files", "*.fit *.FIT"), ("All files", "*.*")])
        if files:
            if not self.selected_files:
                self.file_listbox.delete(0, END)
                self.file_listbox.config(fg='black')
            for f in files:
                if f not in self.selected_files:
                    self.selected_files.append(f)
                    self.file_listbox.insert(END, f"  📄 {os.path.basename(f)}")
            self.file_count.config(text=f"{len(self.selected_files)} file(s) selected")
            self.transfer_btn.config(state=NORMAL)
    
    def clear_files(self):
        self.selected_files = []
        self.file_listbox.delete(0, END)
        self.file_listbox.insert(END, "  Click 'Add Files' to select .FIT files")
        self.file_listbox.config(fg='#999')
        self.file_count.config(text="")
        self.transfer_btn.config(state=DISABLED)
        self.transfer_btn.config(text="Transfer to Watch", bg='#34C759')
        self.transfer_status.config(text="Select files in Step 1, then click 'Transfer to Watch'", fg='#666')
    
    def preview_file(self):
        """Preview the currently selected FIT file(s)"""
        if not self.selected_files:
            messagebox.showinfo("Preview", "Select a .FIT file first")
            return

        # Get selected index or use first file
        sel = self.file_listbox.curselection()
        idx = sel[0] if sel else 0
        if idx < len(self.selected_files):
            filepath = self.selected_files[idx]
            self.show_fit_preview(filepath)

    def show_fit_preview(self, filepath):
        """Show FIT file preview in Garmin watch style"""
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
        preview.geometry("420x700" if not validation['valid'] else "420x650")
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
                  font=('Segoe UI', 11, 'bold'), bg='#dc3545', fg='#fff').pack(anchor='w')

            for issue in validation['issues'][:2]:
                Label(warning_frame, text=issue, font=('Segoe UI', 9),
                      bg='#dc3545', fg='#fff', wraplength=360, justify=LEFT).pack(anchor='w')

            if FITFILETOOL_AVAILABLE:
                def do_repair():
                    new_file, error = self.repair_fit_file(filepath, workout_data)
                    if new_file:
                        messagebox.showinfo("Repaired",
                            f"Workout repaired and saved to:\n{os.path.basename(new_file)}\n\n"
                            "The repaired file uses valid exercise categories that work on all Garmin watches.")
                        if new_file not in self.selected_files:
                            self.selected_files.append(new_file)
                            self.file_listbox.insert(END, f"  📄 {os.path.basename(new_file)} (repaired)")
                            self.file_count.config(text=f"{len(self.selected_files)} file(s) selected")
                    else:
                        messagebox.showerror("Error", f"Could not repair file:\n{error}")

                Button(warning_frame, text="🔧 Repair Workout", font=('Segoe UI', 10, 'bold'),
                       command=do_repair, bg='#fff', fg='#dc3545',
                       padx=12, pady=4, relief=FLAT, cursor='hand2').pack(anchor='w', pady=(5, 0))

        # Watch face simulation (rounded rectangle effect)
        watch_frame = Frame(main, bg='#000', highlightbackground='#333',
                           highlightthickness=2, padx=15, pady=15)
        watch_frame.pack(fill=BOTH, expand=True, pady=(0, 15))

        # Workout title
        title = workout_data.get('name', 'Workout')
        Label(watch_frame, text=title, font=('Segoe UI', 16, 'bold'),
              bg='#000', fg='#fff').pack(pady=(5, 5))

        # Sport type badge
        sport = workout_data.get('sport')
        sub_sport = workout_data.get('sub_sport')
        if sport:
            sport_display = get_sport_display(sport, sub_sport)
            sport_color = get_sport_color(sport, sub_sport)

            sport_badge = Label(watch_frame, text=f"  {sport_display}  ",
                               font=('Segoe UI', 10, 'bold'),
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
            Label(meta_frame, text="  •  ".join(meta_parts), font=('Segoe UI', 9),
                  bg='#000', fg='#666').pack()

        # Scrollable exercise list
        canvas = Canvas(watch_frame, bg='#000', highlightthickness=0, height=350)
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

        # Display exercises
        exercises = workout_data.get('steps', [])
        total_sets = 0

        for i, exercise in enumerate(exercises):
            self.create_exercise_row(exercise_frame, exercise, i, workout_data.get('sport'))
            total_sets += exercise.get('sets', 1)

        # Footer stats
        footer = Frame(watch_frame, bg='#000')
        footer.pack(fill=X, pady=(15, 5))

        stats_text = f"{len(exercises)} steps • {total_sets} total sets" if total_sets > len(exercises) else f"{len(exercises)} steps"
        Label(footer, text=stats_text, font=('Segoe UI', 11),
              bg='#000', fg='#666').pack()

        # Legend - different for cardio vs strength
        legend_frame = Frame(main, bg='#1a1a1a')
        legend_frame.pack(fill=X)

        Label(legend_frame, text="Legend:", font=('Segoe UI', 10, 'bold'),
              bg='#1a1a1a', fg='#888').pack(anchor='w')

        legend_items = Frame(legend_frame, bg='#1a1a1a')
        legend_items.pack(fill=X, pady=(5, 0))

        sport_lower = (sport or '').lower()
        if sport_lower in ['running', 'cycling', 'swimming', 'walking', 'hiking'] or 'run' in sport_lower:
            # Cardio legend
            self.create_legend_badge(legend_items, "Zone/Target", "#3b82f6")
            self.create_legend_badge(legend_items, "Duration", "#8b5cf6")
            self.create_legend_badge(legend_items, "Warmup", "#22c55e")
            self.create_legend_badge(legend_items, "Cooldown", "#6b7280")
        else:
            # Strength legend
            self.create_legend_badge(legend_items, "Reps", "#3b82f6")
            self.create_legend_badge(legend_items, "Duration", "#8b5cf6")
            self.create_legend_badge(legend_items, "Sets", "#22c55e")
            self.create_legend_badge(legend_items, "Weight", "#f97316")
            self.create_legend_badge(legend_items, "Rest", "#6b7280")

        # Close button
        Button(main, text="Close", font=('Segoe UI', 12),
               command=on_close, bg='#333', fg='#fff',
               padx=20, pady=8, relief=FLAT, cursor='hand2').pack(pady=(10, 0))

    def create_exercise_row(self, parent, exercise, index, sport=None):
        """Create a single exercise row in the preview"""
        sport_lower = (sport or '').lower()
        is_cardio = sport_lower in ['running', 'cycling', 'swimming', 'walking', 'hiking'] or 'run' in sport_lower
        step_type = exercise.get('step_type', 'active')

        # Different background colors for different step types
        if step_type == 'warmup':
            bg_color = '#0a2e1a'  # Dark green tint
            border_color = '#22c55e'
        elif step_type == 'cooldown':
            bg_color = '#1a1a2e'  # Dark blue tint
            border_color = '#6b7280'
        elif step_type == 'rest':
            bg_color = '#1a1a1a'
            border_color = '#333'
        else:
            bg_color = '#111'
            border_color = '#222'

        row = Frame(parent, bg=bg_color, highlightbackground=border_color, highlightthickness=1)
        row.pack(fill=X, pady=2, padx=2)

        # Exercise content
        content = Frame(row, bg=bg_color, padx=10, pady=8)
        content.pack(fill=X)

        # Exercise name
        name = exercise.get('name', f'Exercise {index + 1}')
        Label(content, text=name, font=('Segoe UI', 12, 'bold'),
              bg=bg_color, fg='#fff', anchor='w').pack(fill=X)

        # Badges row
        badges = Frame(content, bg=bg_color)
        badges.pack(fill=X, pady=(5, 0))

        if is_cardio:
            # Zone badge (blue) for cardio
            zone = exercise.get('zone')
            if zone:
                self.create_badge(badges, zone, "#3b82f6")

            # Duration badge (purple)
            if exercise.get('duration'):
                duration_str = self.format_duration(exercise['duration'])
                self.create_badge(badges, duration_str, "#8b5cf6")

            # Distance badge (green) for cardio
            if exercise.get('distance'):
                dist_str = self.format_distance(exercise['distance'])
                self.create_badge(badges, dist_str, "#22c55e")

            # Step type indicator
            if step_type == 'warmup':
                Label(badges, text="🔥 Warm Up", font=('Segoe UI', 9),
                      bg='#22c55e', fg='#fff', padx=6, pady=2).pack(side=LEFT, padx=(5, 0))
            elif step_type == 'cooldown':
                Label(badges, text="❄️ Cool Down", font=('Segoe UI', 9),
                      bg='#6b7280', fg='#fff', padx=6, pady=2).pack(side=LEFT, padx=(5, 0))
        else:
            # Strength workout badges
            # Reps badge (blue)
            if exercise.get('reps'):
                self.create_badge(badges, f"{exercise['reps']} reps", "#3b82f6")

            # Duration badge (purple)
            if exercise.get('duration'):
                duration_str = self.format_duration(exercise['duration'])
                self.create_badge(badges, duration_str, "#8b5cf6")

            # Distance badge (green)
            if exercise.get('distance'):
                dist_str = self.format_distance(exercise['distance'])
                self.create_badge(badges, dist_str, "#22c55e")

            # Sets badge (green, only if > 1)
            sets = exercise.get('sets', 1)
            if sets > 1:
                self.create_badge(badges, f"{sets} sets", "#22c55e")

            # Weight badge (orange)
            if exercise.get('weight'):
                self.create_badge(badges, exercise['weight'], "#f97316")

            # Rest badge (gray)
            if exercise.get('rest'):
                rest_str = self.format_duration(exercise['rest'])
                self.create_badge(badges, f"Rest {rest_str}", "#6b7280")

            # Exercise type
            ex_type = exercise.get('type', '')
            if ex_type and ex_type not in name.lower():
                Label(badges, text=ex_type.title(), font=('Segoe UI', 9),
                      bg='#333', fg='#999', padx=6, pady=2).pack(side=LEFT, padx=(5, 0))

    def create_badge(self, parent, text, color):
        """Create a colored badge"""
        badge = Label(parent, text=text, font=('Segoe UI', 10, 'bold'),
                     bg=color, fg='#fff', padx=8, pady=2)
        badge.pack(side=LEFT, padx=(0, 5))

    def create_legend_badge(self, parent, text, color):
        """Create a legend badge"""
        item = Frame(parent, bg='#1a1a1a')
        item.pack(side=LEFT, padx=(0, 15))

        Label(item, text="●", font=('Segoe UI', 10), bg='#1a1a1a', fg=color).pack(side=LEFT)
        Label(item, text=text, font=('Segoe UI', 10), bg='#1a1a1a', fg='#888').pack(side=LEFT, padx=(3, 0))

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
        """Validate FIT file for issues that may prevent it from working on Garmin watches."""
        # Try fitfiletool's validator first
        if FITFILETOOL_AVAILABLE:
            result = fitfiletool_validate_fit_file(filepath)
            if result:
                # Convert to expected format if needed
                return {
                    'valid': result.get('valid', True),
                    'issues': result.get('issues', []),
                    'invalid_categories': result.get('invalid_categories', [])
                }

        # Fall back to local implementation
        if not FITPARSE_AVAILABLE:
            return {'valid': True, 'issues': [], 'invalid_categories': []}

        try:
            fitfile = FitFile(filepath)
            issues = []
            invalid_categories = []

            VALID_CATEGORIES = set(range(33))

            for record in fitfile.get_messages('workout_step'):
                for field in record.fields:
                    if field.name == 'exercise_category' and field.value is not None:
                        if isinstance(field.value, int) and field.value not in VALID_CATEGORIES:
                            invalid_categories.append(field.value)

            if invalid_categories:
                unique_invalid = list(set(invalid_categories))
                issues.append(f"Invalid exercise categories found: {unique_invalid}")
                issues.append("These may cause the workout to not appear on your Garmin watch.")

            return {
                'valid': len(issues) == 0,
                'issues': issues,
                'invalid_categories': list(set(invalid_categories))
            }
        except Exception as e:
            return {'valid': False, 'issues': [f"Error validating file: {str(e)}"], 'invalid_categories': []}

    def repair_fit_file(self, filepath, workout_data):
        """Repair a FIT file by regenerating it with valid exercise categories."""
        if not FITFILETOOL_AVAILABLE:
            return None, "amakaflow-fitfiletool not available"

        try:
            exercises = []
            for step in workout_data.get('steps', []):
                name = step.get('name', 'Exercise')

                if step.get('reps'):
                    reps = step['reps']
                elif step.get('distance'):
                    dist = step['distance']
                    if dist >= 1000:
                        reps = f"{dist/1000:.1f}km"
                    else:
                        reps = f"{int(dist)}m"
                elif step.get('duration'):
                    reps = f"{int(step['duration'])}s"
                else:
                    reps = 10

                exercises.append({
                    'name': name,
                    'reps': reps,
                    'sets': step.get('sets', 1)
                })

            workout = {
                'title': workout_data.get('name', 'Repaired Workout'),
                'blocks': [{'exercises': exercises, 'rest_between_sec': 0}]
            }

            fit_bytes = build_fit_workout(workout, use_lap_button=False)

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
                        step['duration_type'] = str(field.value)
                    elif field.name == 'duration_reps' and field.value:
                        step['reps'] = int(field.value)
                    elif field.name == 'duration_time' and field.value:
                        step['duration'] = float(field.value)
                    elif field.name == 'duration_distance' and field.value:
                        step['distance'] = float(field.value)
                    elif field.name == 'intensity':
                        intensity = str(field.value) if field.value else None
                        step['intensity'] = intensity
                        if intensity == 'rest':
                            step['is_rest'] = True
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
            exercises = []
            i = 0
            while i < len(steps_raw):
                step = steps_raw[i]

                # Handle repeat markers for strength workouts
                if step.get('is_repeat'):
                    if exercises and step.get('repeat_count'):
                        exercises[-1]['sets'] = step['repeat_count'] + 1
                    i += 1
                    continue

                # For strength workouts, skip pure rest steps
                if not is_cardio and step.get('is_rest'):
                    if exercises and step.get('duration'):
                        exercises[-1]['rest'] = step['duration']
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

                    if intensity == 'warmup':
                        exercise['name'] = 'Warm Up'
                        exercise['step_type'] = 'warmup'
                    elif intensity == 'cooldown':
                        exercise['name'] = 'Cool Down'
                        exercise['step_type'] = 'cooldown'
                    elif intensity == 'rest':
                        exercise['name'] = 'Recovery'
                        exercise['step_type'] = 'rest'
                    elif intensity == 'active':
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
                        exercise['name'] = f'Exercise {i + 1}'

                    if cat:
                        exercise['type'] = cat.replace('_', ' ')

                # Add duration
                if step.get('duration'):
                    exercise['duration'] = step['duration']

                # Add distance
                if step.get('distance'):
                    exercise['distance'] = step['distance']

                # Add reps for strength workouts
                if step.get('reps'):
                    exercise['reps'] = step['reps']

                # Add weight
                if step.get('weight'):
                    unit = step.get('weight_unit', 'kg')
                    exercise['weight'] = f"{int(step['weight'])} {unit}"

                # Add target/zone info for cardio
                if step.get('target_type') and step.get('target_value'):
                    target_type = step['target_type']
                    if 'heart_rate' in target_type:
                        exercise['zone'] = f"HR Zone {int(step['target_value'])}"
                    elif 'speed' in target_type or 'pace' in target_type:
                        exercise['zone'] = f"Pace {step['target_value']}"
                    elif 'power' in target_type:
                        exercise['zone'] = f"{int(step['target_value'])}W"

                exercise['sets'] = 1  # default
                exercises.append(exercise)
                i += 1

            workout_data['steps'] = exercises
            return workout_data

        except Exception as e:
            print(f"Error parsing FIT file with fitparse: {e}")
            return None

    def parse_fit_basic(self, filepath):
        """Basic FIT file parsing without fitparse library"""
        try:
            with open(filepath, 'rb') as f:
                data = f.read()

            # Very basic parsing - just show it's a workout file
            workout_data = {
                'name': 'Workout',
                'sport': 'training',
                'steps': [{'name': '(Install fitparse for detailed view)', 'type': 'info'}],
                'created': None,
                'source': 'FIT File'
            }

            # Try to extract file size at least
            workout_data['size'] = len(data)

            return workout_data
        except Exception as e:
            print(f"Error parsing FIT file: {e}")
            return None

    
    def transfer(self):
        if not self.selected_files:
            return
        self.kill_garmin_express()
        device = self.detect_garmin_device()
        if not device:
            messagebox.showwarning("No Device", "Garmin watch not detected.\n\nMake sure:\n• Watch is connected via USB\n• Watch screen is awake")
            return

        if not self.is_mtp and self.garmin_newfiles:
            # Mass Storage Mode - direct file copy
            # Try Workouts folder first (newer watches), then NewFiles (older watches)
            garmin_root = os.path.dirname(self.garmin_newfiles)
            workouts_path = os.path.join(garmin_root, "Workouts")

            target_path = None
            folder_name = None

            if os.path.exists(workouts_path):
                target_path = workouts_path
                folder_name = "Workouts"
            elif os.path.exists(self.garmin_newfiles):
                target_path = self.garmin_newfiles
                folder_name = "NewFiles"
            else:
                # Create NewFiles as fallback
                os.makedirs(self.garmin_newfiles, exist_ok=True)
                target_path = self.garmin_newfiles
                folder_name = "NewFiles"

            count = 0
            for f in self.selected_files:
                try:
                    shutil.copy2(f, os.path.join(target_path, os.path.basename(f)))
                    count += 1
                except:
                    pass
            if count:
                self.transfer_btn.config(text="✓ Transferred!", bg='#28a745', state=DISABLED)
                self.transfer_status.config(text=f"✅ {count} file(s) transferred to GARMIN/{folder_name}!", fg='#2e7d32')
                messagebox.showinfo("Success", f"✅ {count} file(s) transferred to GARMIN/{folder_name}!\n\nYou can now disconnect your watch.")
        elif self.is_mtp:
            # MTP Mode - open File Explorer for manual drag-and-drop
            # Stage files first
            for f in self.staging_folder.glob('*.fit'):
                f.unlink()
            for f in self.staging_folder.glob('*.FIT'):
                f.unlink()

            count = 0
            for f in self.selected_files:
                try:
                    shutil.copy2(f, self.staging_folder / os.path.basename(f))
                    count += 1
                except:
                    pass

            if count:
                # Open staging folder first
                subprocess.Popen(['explorer', str(self.staging_folder)])
                time.sleep(0.3)

                # Try to open Garmin Workouts folder directly using COM
                workouts_opened = False
                if WIN32COM_AVAILABLE:
                    try:
                        shell = win32com.client.Dispatch("Shell.Application")
                        this_pc = shell.Namespace(17)

                        # Find Garmin device
                        for item in this_pc.Items():
                            if self.mtp_device_name and self.mtp_device_name.lower() in item.Name.lower():
                                device_folder = item.GetFolder

                                # Navigate to Internal Storage
                                storage_folder = device_folder
                                for storage_item in device_folder.Items():
                                    if 'storage' in storage_item.Name.lower():
                                        storage_folder = storage_item.GetFolder
                                        break

                                # Navigate to GARMIN
                                for garmin_item in storage_folder.Items():
                                    if garmin_item.Name.upper() == "GARMIN":
                                        garmin_folder = garmin_item.GetFolder

                                        # Navigate to Workouts
                                        for workout_item in garmin_folder.Items():
                                            if workout_item.Name.upper() == "WORKOUTS":
                                                # Open the Workouts folder
                                                workout_path = workout_item.Path
                                                subprocess.Popen(['explorer', workout_path])
                                                workouts_opened = True
                                                break
                                        break
                                break
                    except Exception as e:
                        print(f"Error opening Workouts folder: {e}")

                if not workouts_opened:
                    # Fallback: open This PC
                    subprocess.Popen(['explorer', 'shell::{{20D04FE0-3AEA-1069-A2D8-08002B30309D}}'])

                self.transfer_btn.config(text="✓ Files Ready!", bg='#FF9500', state=DISABLED)
                self.transfer_status.config(text=f"✓ {count} file(s) ready. Drag from left window to right window.", fg='#FF9500')

                msg = f"✓ {count} file(s) are ready!\n\n"
                if workouts_opened:
                    msg += "Two File Explorer windows have opened:\n\n"
                    msg += "1. Left: Your workout files\n"
                    msg += "2. Right: Garmin Workouts folder\n\n"
                    msg += "Drag the files from left to right."
                else:
                    msg += "Two File Explorer windows have opened:\n\n"
                    msg += "1. Left: Your workout files\n"
                    msg += "2. Right: This PC\n\n"
                    msg += f"Navigate to: {self.mtp_device_name} → Internal Storage → GARMIN → Workouts\n\n"
                    msg += "Then drag the files from left to right."

                messagebox.showinfo("Manual Transfer", msg)
        else:
            # Fallback - stage files for manual drag and drop
            for f in self.staging_folder.glob('*.fit'):
                f.unlink()
            for f in self.staging_folder.glob('*.FIT'):
                f.unlink()
            count = 0
            for f in self.selected_files:
                try:
                    shutil.copy2(f, self.staging_folder / os.path.basename(f))
                    count += 1
                except:
                    pass
            if count:
                self.transfer_btn.config(text="✓ Files Staged!", bg='#666', state=DISABLED)
                if not WIN32COM_AVAILABLE:
                    msg = f"✓ {count} file(s) staged. Install pywin32 for automatic MTP transfer, or drag files manually."
                else:
                    msg = f"✓ {count} file(s) staged. Drag them to your watch in File Explorer."
                self.transfer_status.config(text=msg, fg='#2e7d32')
                subprocess.run(['explorer', str(self.staging_folder)])
                subprocess.run(['explorer', 'shell:MyComputerFolder'])

def main():
    try:
        if DND_AVAILABLE:
            try:
                root = TkinterDnD.Tk()
            except:
                root = Tk()
        else:
            root = Tk()
        GarminUploaderWin(root)
        root.mainloop()
    except Exception as e:
        # Write error to log file for debugging windowed builds
        import traceback
        log_file = os.path.join(os.path.expanduser('~'), 'garmin_uploader_error.log')
        with open(log_file, 'w') as f:
            f.write(f"Error starting Garmin Workout Uploader:\n")
            f.write(f"{str(e)}\n\n")
            f.write(traceback.format_exc())

        # Try to show error dialog if possible
        try:
            from tkinter import Tk, messagebox
            error_root = Tk()
            error_root.withdraw()
            messagebox.showerror(
                "Startup Error",
                f"Failed to start application.\nError log saved to:\n{log_file}\n\nError: {str(e)}"
            )
            error_root.destroy()
        except:
            pass
        raise

if __name__ == "__main__":
    main()