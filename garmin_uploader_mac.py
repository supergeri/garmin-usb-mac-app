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
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog, messagebox

# Try to import tkinterdnd2 for drag and drop support
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


class GarminUploaderMac:
    def __init__(self, root):
        self.root = root
        self.root.title("Garmin Workout Uploader")
        self.root.geometry("580x620")
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
        
        # Track drag state for visual feedback
        self.is_dragging = False
        
        self.create_menu()
        self.create_ui()
    
    def check_openmtp(self):
        """Check if OpenMTP is installed"""
        paths = [
            Path("/Applications/OpenMTP.app"),
            self.home / "Applications/OpenMTP.app"
        ]
        return any(p.exists() for p in paths)
    
    def create_menu(self):
        """Create the application menu bar"""
        menubar = Menu(self.root)
        self.root.config(menu=menubar)
        
        # Tools menu
        tools_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        
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
        
        # Listbox
        self.file_listbox = Listbox(self.drop_zone, height=4, font=('SF Pro Text', 11),
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
        
        self.file_count = Label(btn_frame, text="", font=('SF Pro Text', 11), bg='#fff', fg='#666')
        self.file_count.pack(side=RIGHT)
        
        # Drag and drop status indicator
        if DND_AVAILABLE:
            self.dnd_status = Label(btn_frame, text="📥 Drop enabled", font=('SF Pro Text', 10), 
                                    bg='#fff', fg='#34C759')
            self.dnd_status.pack(side=RIGHT, padx=(0, 10))
    
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
        
        # Initial state - waiting
        self.transfer_status = Label(parent, 
            text="Stage your files first (Step 2), then transfer instructions will appear here.",
            font=('SF Pro Text', 11), bg='#fff', fg='#666', wraplength=480, justify=LEFT)
        self.transfer_status.pack(fill=X)
    
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
        
        # Update Step 3 with transfer instructions
        self.show_transfer_instructions(staged)
        
        # Open OpenMTP and staging folder
        self.open_openmtp()
        subprocess.run(['open', str(self.staging_folder)])
        
        # Show success
        self.prepare_btn.config(text="✓ Files Staged!", bg='#666', state=DISABLED)
    
    def show_transfer_instructions(self, staged_files):
        """Show detailed transfer instructions in Step 3"""
        # Clear current content
        for widget in self.transfer_frame.winfo_children():
            widget.destroy()
        
        # Success message
        success_frame = Frame(self.transfer_frame, bg='#e8f5e9', padx=10, pady=8)
        success_frame.pack(fill=X, pady=(0, 10))
        
        Label(success_frame, text=f"✓ {len(staged_files)} file(s) ready to transfer!", 
              font=('SF Pro Text', 12, 'bold'), bg='#e8f5e9', fg='#2e7d32').pack()
        
        # Instructions with visual
        instr = Frame(self.transfer_frame, bg='#fff')
        instr.pack(fill=X)
        
        # Visual diagram
        diagram = """
┌─────────────────────┐         ┌─────────────────────┐
│   📁 GarminWorkouts │   ───►  │  📁 GARMIN          │
│   (Finder window)   │  DRAG   │     └─ 📁 NewFiles  │
│                     │         │        (in OpenMTP) │
└─────────────────────┘         └─────────────────────┘
"""
        
        Label(instr, text="Drag your files:", font=('SF Pro Text', 11, 'bold'), 
              bg='#fff', anchor='w').pack(fill=X)
        
        Label(instr, text=diagram, font=('Menlo', 10), bg='#f8f8f8', 
              justify=LEFT, padx=10, pady=5).pack(fill=X, pady=5)
        
        steps = """1. A Finder window opened with your workout files
2. OpenMTP should show your Garmin watch
3. In OpenMTP: Navigate to GARMIN → NewFiles
4. Drag all .fit files from Finder into NewFiles
5. Wait for transfer to complete
6. Disconnect watch - done! 🎉"""
        
        Label(instr, text=steps, font=('SF Pro Text', 11), bg='#fff',
              justify=LEFT, anchor='w').pack(fill=X, pady=(5, 0))
        
        # Buttons
        btn_frame = Frame(self.transfer_frame, bg='#fff')
        btn_frame.pack(fill=X, pady=(12, 0))
        
        Button(btn_frame, text="📂 Open Folder Again", font=('SF Pro Text', 11),
               command=lambda: subprocess.run(['open', str(self.staging_folder)]),
               padx=10, pady=5, relief=FLAT).pack(side=LEFT)
        
        Button(btn_frame, text="🔄 Open OpenMTP", font=('SF Pro Text', 11),
               command=self.open_openmtp, padx=10, pady=5, relief=FLAT).pack(side=LEFT, padx=(8, 0))
        
        # If OpenMTP not installed
        if not self.openmtp_installed:
            warning = Frame(self.transfer_frame, bg='#fff3e0', padx=10, pady=8)
            warning.pack(fill=X, pady=(10, 0))
            
            Label(warning, text="⚠️ OpenMTP not found!", 
                  font=('SF Pro Text', 11, 'bold'), bg='#fff3e0', fg='#e65100').pack()
            
            Label(warning, text="Download it free from: openmtp.ganeshrvel.com", 
                  font=('SF Pro Text', 11), bg='#fff3e0').pack()
            
            Button(warning, text="Download OpenMTP", font=('SF Pro Text', 11),
                   command=lambda: webbrowser.open('https://openmtp.ganeshrvel.com'),
                   bg='#ff9800', fg='white', padx=10, pady=5, relief=FLAT,
                   cursor='hand2').pack(pady=(5, 0))
    
    def kill_garmin_express(self):
        """Kill Garmin Express to free MTP access"""
        subprocess.run(['pkill', '-f', 'Garmin Express'], capture_output=True)
        subprocess.run(['pkill', '-f', 'GarminExpressService'], capture_output=True)
    
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
            "Garmin Workout Uploader\n\n"
            "Version 1.0\n\n"
            "A simple tool to upload .FIT workout files\n"
            "to your Garmin watch via OpenMTP.\n\n"
            "Tools menu powered by GOTOES.org")


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
