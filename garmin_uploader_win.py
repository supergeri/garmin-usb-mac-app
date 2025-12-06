#!/usr/bin/env python3
"""Garmin Workout Uploader for Windows"""

import os, sys, shutil, subprocess, re, threading, time, ctypes
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox

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

# Try to import win32com for MTP file transfer
try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False

class GarminUploaderWin:
    def __init__(self, root):
        self.root = root
        self.root.title("Garmin Workout Uploader")
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
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.create_ui()
        self.root.update()
    
    def _on_close(self):
        self._monitor_running = False
        self.root.destroy()
    
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
            result = subprocess.run(['powershell', '-Command', ps_cmd], capture_output=True, text=True, timeout=10)
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
            result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq GarminExpress.exe'], capture_output=True, text=True, timeout=5)
            return 'GarminExpress.exe' in result.stdout
        except:
            return False
    
    def kill_garmin_express(self):
        try:
            subprocess.run(['taskkill', '/F', '/IM', 'GarminExpress.exe'], timeout=5, capture_output=True)
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
            garmin_folder = None

            # Search for Garmin device
            for item in this_pc.Items():
                if self.mtp_device_name and self.mtp_device_name.lower() in item.Name.lower():
                    garmin_folder = shell.Namespace(item.Path)
                    break

            if not garmin_folder:
                return False, "Could not find Garmin device in This PC"

            # Navigate to GARMIN folder
            garmin_main = None
            for item in garmin_folder.Items():
                if item.Name.upper() == "GARMIN":
                    garmin_main = shell.Namespace(item.Path)
                    break

            if not garmin_main:
                return False, "Could not find GARMIN folder on device"

            # Navigate to or create NewFiles folder
            newfiles_folder = None
            for item in garmin_main.Items():
                if item.Name.upper() == "NEWFILES":
                    newfiles_folder = shell.Namespace(item.Path)
                    break

            if not newfiles_folder:
                # Try to create NewFiles folder
                garmin_main.NewFolder("NewFiles")
                time.sleep(0.5)
                for item in garmin_main.Items():
                    if item.Name.upper() == "NEWFILES":
                        newfiles_folder = shell.Namespace(item.Path)
                        break

            if not newfiles_folder:
                return False, "Could not access or create NewFiles folder"

            # Copy files
            copied = 0
            for filepath in files:
                try:
                    # CopyHere flags: 4 = no progress dialog, 16 = yes to all
                    newfiles_folder.CopyHere(filepath, 4 | 16)
                    copied += 1
                    time.sleep(0.3)  # Small delay between files
                except Exception as e:
                    print(f"Error copying {filepath}: {e}")
                    continue

            if copied > 0:
                return True, f"{copied} file(s) transferred"
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

        # Create preview window
        preview = Toplevel(self.root)
        preview.title(f"Workout Preview - {os.path.basename(filepath)}")
        preview.geometry("420x650")
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
            sport_colors = {
                'running': '#22c55e',
                'cycling': '#f97316',
                'swimming': '#3b82f6',
                'strength_training': '#ef4444',
                'training': '#8b5cf6',
                'walking': '#84cc16',
                'hiking': '#a3e635'
            }
            sport_display = sub_sport.replace('_', ' ').title() if sub_sport else sport.replace('_', ' ').title()
            sport_color = sport_colors.get(sport, '#6b7280')

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

        if sport in ['running', 'cycling', 'swimming', 'walking', 'hiking']:
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
        is_cardio = sport in ['running', 'cycling', 'swimming', 'walking', 'hiking']
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

    def parse_fit_file(self, filepath):
        """Parse a FIT file and extract workout data"""
        if FITPARSE_AVAILABLE:
            return self.parse_fit_with_fitparse(filepath)
        else:
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
            is_cardio = workout_data.get('sport') in ['running', 'cycling', 'swimming', 'walking', 'hiking']

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
            if not os.path.exists(self.garmin_newfiles):
                os.makedirs(self.garmin_newfiles)
            count = 0
            for f in self.selected_files:
                try:
                    shutil.copy2(f, os.path.join(self.garmin_newfiles, os.path.basename(f)))
                    count += 1
                except:
                    pass
            if count:
                self.transfer_btn.config(text="✓ Transferred!", bg='#28a745', state=DISABLED)
                self.transfer_status.config(text=f"✅ {count} file(s) transferred to Garmin!", fg='#2e7d32')
                messagebox.showinfo("Success", f"✅ {count} file(s) transferred to your Garmin!\n\nYou can now disconnect your watch.")
        elif self.is_mtp and WIN32COM_AVAILABLE:
            # MTP Mode - use COM API for transfer
            self.transfer_btn.config(text="Transferring...", state=DISABLED)
            self.transfer_status.config(text="Transferring files via MTP...", fg='#666')
            self.root.update()

            success, message = self.transfer_mtp_files(self.selected_files)

            if success:
                self.transfer_btn.config(text="✓ Transferred!", bg='#28a745')
                self.transfer_status.config(text=f"✅ {message} to Garmin!", fg='#2e7d32')
                messagebox.showinfo("Success", f"✅ {message} to your Garmin!\n\nYou can now disconnect your watch.")
            else:
                self.transfer_btn.config(text="Transfer Files", bg='#007AFF', state=NORMAL)
                self.transfer_status.config(text=f"❌ {message}", fg='#dc3545')
                messagebox.showerror("Transfer Failed", f"Could not transfer files via MTP:\n\n{message}\n\nTry reconnecting your watch or using Mass Storage mode.")
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
    if DND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except:
            root = Tk()
    else:
        root = Tk()
    GarminUploaderWin(root)
    root.mainloop()

if __name__ == "__main__":
    main()