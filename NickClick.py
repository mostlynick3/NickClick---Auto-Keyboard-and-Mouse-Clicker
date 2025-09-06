import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog, Menu
import pyautogui
import json
import threading
import time
import keyboard
import sv_ttk
import hashlib
import os
import psutil
import sys
from pathlib import Path
from pynput import mouse, keyboard as pynput_keyboard
import ctypes
from datetime import datetime, timedelta
from tkcalendar import Calendar
try:
   ctypes.windll.shcore.SetProcessDpiAwareness(1)
except:
   pass

try:
   from tkcalendar import DateEntry
   TKCALENDAR_AVAILABLE = True
except ImportError:
   TKCALENDAR_AVAILABLE = False

class AutomationGUI:
   def __init__(self):
   
       self.root = tk.Tk()
       self.setup_icon()
       self.root.geometry("1050x700")
       self.root.minsize(width=1050, height=400)
       self.actions = []
       self.recording = False
       self.editable = True
       self.dark_mode = False
       self.record_key = None
       self.unsaved_changes = False
       self.script_running = False
       self.stop_script = False
       self.scheduled_tasks = []
       self.schedule_thread = None
       self.schedule_active = False
       self.load_preferences()
       self.setup_ui()
       self.setup_recording()
       self.apply_theme()
       self.setup_empty_state()
       self.start_scheduler()
       self.update_window_title()
       pyautogui.FAILSAFE = False
       self.root.protocol("WM_DELETE_WINDOW", self.on_closing)


       if self.check_for_other_instances():
           self.dialog_result = None
           self.show_instance_warning_dialog()
           if self.dialog_result == "exit":
               sys.exit(0)

   def show_instance_warning_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Warning")
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_force()

        self.center_window(dialog, 400, 200)

        label = ttk.Label(
            dialog,
            text="Warning: Multiple instances of this script are running.\n \n"
                 "This can cause issues with overlapping script execution and scheduling.\n \n"
                 "Proceed with caution.",
            justify=tk.CENTER,
            wraplength=380
        )
        label.pack(pady=20, padx=20)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        ok_button = ttk.Button(
            button_frame,
            text="OK",
            command=lambda: self.on_dialog_ok(dialog)
        )
        ok_button.pack(side=tk.LEFT, padx=10)

        exit_button = ttk.Button(
            button_frame,
            text="Exit",
            command=lambda: self.on_dialog_exit(dialog)
        )
        exit_button.pack(side=tk.LEFT, padx=10)

        dialog.bind('<Return>', lambda e: self.on_dialog_ok(dialog))
        dialog.bind('<Escape>', lambda e: self.on_dialog_exit(dialog))

        self.root.wait_window(dialog)
        return self.dialog_result

   def on_dialog_ok(self, dialog):
        self.dialog_result = "ok"
        dialog.destroy()

   def on_dialog_exit(self, dialog):
        self.dialog_result = "exit"
        dialog.destroy()


   def check_for_other_instances(self):
        current_pid = os.getpid()
        current_process = psutil.Process(current_pid)
        try:
            current_cmdline = current_process.cmdline()
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            return False
            
        script_path = None
        if len(current_cmdline) > 1:
            script_path = current_cmdline[1] if current_cmdline[0].endswith('python.exe') or current_cmdline[0].endswith('python') else current_cmdline[0]

        if not script_path:
            return False

        matching_processes = []
        for proc in psutil.process_iter(['pid', 'cmdline']):
            try:
                if proc.info['pid'] == current_pid:
                    continue
                if not proc.info['cmdline']:
                    continue
                proc_script_path = None
                if len(proc.info['cmdline']) > 1:
                    proc_script_path = proc.info['cmdline'][1] if proc.info['cmdline'][0].endswith('python.exe') or proc.info['cmdline'][0].endswith('python') else proc.info['cmdline'][0]
                if proc_script_path and proc_script_path == script_path:
                    matching_processes.append(proc.info['pid'])
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                pass

        if matching_processes:
            return True
        else:
            return False

   def new_script(self):
       if not self.prompt_unsaved_changes():
           return
       if self.actions:
           result = messagebox.askyesno("New Script", "This will clear your actions table. Please save any changes before proceeding. Continue?")
           if not result:
               return
       self.actions.clear()
       if hasattr(self, 'current_file'):
           delattr(self, 'current_file')
       self.unsaved_changes = False
       self.clear_schedule()
       self.update_tree()
       self.update_window_title()

   def delete_script(self):
       if not hasattr(self, 'current_file'):
           if self.actions:
               result = messagebox.askyesno("Delete Script", "Are you sure you want to delete your script?")
               if result:
                   self.actions.clear()
                   self.clear_schedule()
                   self.update_tree()
           else:
               messagebox.showwarning("Delete Script", "No script file is currently loaded.")
           return

       script_name = os.path.basename(self.current_file)
       result = messagebox.askyesno("Delete Script",
                                   f"Are you sure you want to delete script '{script_name}'?")
       if result:
           try:
               os.remove(self.current_file)
               self.remove_from_recent(self.current_file)
               delattr(self, 'current_file')
               self.actions.clear()
               self.clear_schedule()
               self.update_tree()
               messagebox.showinfo("Deleted", f"Script '{script_name}' has been deleted.")
           except Exception as e:
               messagebox.showerror("Error", f"Failed to delete script: {str(e)}")

   def add_to_recent(self, filepath):
       prefs = self.load_preferences()
       recent = prefs.get('recent_files', [])

       if filepath in recent:
           recent.remove(filepath)

       recent.insert(0, filepath)
       recent = recent[:10]

       prefs['recent_files'] = recent
       config_path = self.get_config_path()
       with open(config_path, 'w') as f:
           json.dump(prefs, f, indent=2)

       self.rebuild_file_menu()

   def remove_from_recent(self, filepath):
      prefs = self.load_preferences()
      recent = prefs.get('recent_files', [])

      if filepath in recent:
       recent.remove(filepath)
       prefs['recent_files'] = recent
       config_path = self.get_config_path()
       with open(config_path, 'w') as f:
           json.dump(prefs, f, indent=2)
       self.rebuild_file_menu()

   def load_script_data(self, filepath):
       try:
           with open(filepath, 'r') as f:
               data = json.load(f)

           if isinstance(data, dict) and "actions" in data:
               self.actions = data["actions"]
               locked = data.get("locked", False)
               self.editable_var.set(not locked)
               self.update_ui_editable_state()
           else:
               self.actions = data
               self.editable_var.set(True)
               self.update_ui_editable_state()

           self.current_file = filepath
           self.clear_schedule()
           self.update_tree()
           self.update_window_title()
           return True
       except Exception as e:
           return False, str(e)

   def load_recent_file(self, filepath):
       if not self.prompt_unsaved_changes():
           return
       result = self.load_script_data(filepath)
       if result is True:
           messagebox.showinfo("Loaded", f"Script loaded from {os.path.basename(filepath)}")
           self.unsaved_changes = False
       else:
           success, error = result
           result = messagebox.askyesno(
               "Load Error",
               f"Failed to load script '{os.path.basename(filepath)}':\n{error}\n\nWould you like to remove it from the recent files list?"
           )
           if result:
               self.remove_from_recent(filepath)

   def load_script(self):
       if not self.prompt_unsaved_changes():
           return
       filename = filedialog.askopenfilename(
           filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
       )
       if filename:
           result = self.load_script_data(filename)
           if result is True:
               self.add_to_recent(filename)
               messagebox.showinfo("Loaded", f"Script loaded from {filename}")
               self.unsaved_changes = False
           else:
               success, error = result
               messagebox.showerror("Error", f"Failed to load script: {error}")

   def on_closing(self):
       if not self.prompt_unsaved_changes():
           return
       self.clear_schedule()
       self.save_preferences()
       self.root.destroy()

   def get_config_path(self):
       if os.name == 'nt':
           config_dir = Path(os.environ.get('APPDATA', '')) / 'NickClick'
       else:
           config_dir = Path.home() / '.config' / 'NickClick'

       config_dir.mkdir(parents=True, exist_ok=True)
       return config_dir / 'preferences.json'

   def get_schedule_path(self):
       if os.name == 'nt':
           config_dir = Path(os.environ.get('APPDATA', '')) / 'NickClick'
       else:
           config_dir = Path.home() / '.config' / 'NickClick'

       config_dir.mkdir(parents=True, exist_ok=True)
       return config_dir / 'schedule.json'

   def save_preferences(self):
       try:
           prefs = self.load_preferences()
           prefs.update({
               'dark_mode': self.dark_mode,
               'window_geometry': self.root.geometry(),
               'editable': self.editable_var.get() if hasattr(self, 'editable_var') else True
           })

           config_path = self.get_config_path()
           with open(config_path, 'w') as f:
               json.dump(prefs, f, indent=2)
       except Exception as e:
           print(f"Failed to save preferences: {e}")

   def load_preferences(self):
       try:
           config_path = self.get_config_path()
           if config_path.exists():
               with open(config_path, 'r') as f:
                   prefs = json.load(f)

               self.dark_mode = prefs.get('dark_mode', False)

               geometry = prefs.get('window_geometry')
               if geometry:
                   self.root.geometry(geometry)

               return prefs
       except Exception as e:
           print(f"Failed to load preferences: {e}")

       return {}

   def save_schedule(self, schedule_data):
       try:
           schedule_path = self.get_schedule_path()
           with open(schedule_path, 'w') as f:
               json.dump(schedule_data, f, indent=2)
       except Exception as e:
           messagebox.showerror("Error", f"Failed to save schedule: {e}")

   def load_schedule(self):
       try:
           schedule_path = self.get_schedule_path()
           if schedule_path.exists():
               with open(schedule_path, 'r') as f:
                   return json.load(f)
       except Exception as e:
           print(f"Failed to load schedule: {e}")
       return None

   def clear_schedule(self):
       self.schedule_active = False
       self.update_window_title()
       self.scheduled_tasks.clear()
       try:
           schedule_path = self.get_schedule_path()
           if schedule_path.exists():
               os.remove(schedule_path)
       except Exception as e:
           print(f"Failed to clear schedule file: {e}")

   def setup_icon(self):
       try:
           from PIL import Image, ImageDraw, ImageTk

           img = Image.new('RGBA', (32, 32), (255, 255, 255, 0))
           draw = ImageDraw.Draw(img)

           draw.rectangle([6, 12, 26, 24], fill='#4a90e2', outline='#2c5aa0', width=2)
           draw.rectangle([7, 13, 25, 15], fill='#6bb6ff')

           draw.line([8, 18, 24, 18], fill='#2c5aa0', width=1)
           draw.line([8, 20, 22, 20], fill='#2c5aa0', width=1)

           pointer_points = [(2, 2), (2, 18), (6, 14), (9, 17), (12, 13), (7, 8), (2, 2)]
           draw.polygon(pointer_points, fill='white', outline='black')

           draw.ellipse([15, 16, 18, 19], fill='#ff6b6b')
           draw.ellipse([19, 18, 21, 20], fill='#ff9999')

           img = img.resize((16, 16), Image.Resampling.LANCZOS)
           photo = ImageTk.PhotoImage(img)
           self.root.iconphoto(True, photo)

       except ImportError:
           self.root.title("🖱️ " + self.root.title())
       except Exception as e:
           self.root.title("🖱️ " + self.root.title())

   def center_window(self, window, width, height):
       screen_width = window.winfo_screenwidth()
       screen_height = window.winfo_screenheight()
       x = (screen_width - width) // 2
       y = (screen_height - height) // 2
       window.geometry(f"{width}x{height}+{x}+{y}")

   def format_key(self, key_str):
       if not key_str:
           return key_str
       escape_map = {
           '\u0000': 'CTRL+@',
           '\u0001': 'CTRL+A',
           '\u0002': 'CTRL+B',
           '\u0003': 'CTRL+C',
           '\u0004': 'CTRL+D',
           '\u0005': 'CTRL+E',
           '\u0006': 'CTRL+F',
           '\u0007': 'CTRL+G',
           '\u0008': 'BACKSPACE',
           '\u0009': 'TAB',
           '\u000A': 'ENTER',
           '\u000B': 'CTRL+K',
           '\u000C': 'CTRL+L',
           '\u000D': 'ENTER',
           '\u000E': 'CTRL+N',
           '\u000F': 'CTRL+O',
           '\u0010': 'CTRL+P',
           '\u0011': 'CTRL+Q',
           '\u0012': 'CTRL+R',
           '\u0013': 'CTRL+S',
           '\u0014': 'CTRL+T',
           '\u0015': 'CTRL+U',
           '\u0016': 'CTRL+V',
           '\u0017': 'CTRL+W',
           '\u0018': 'CTRL+X',
           '\u0019': 'CTRL+Y',
           '\u001A': 'CTRL+Z',
           '\u001B': 'ESC',
           '\u001C': 'CTRL+\\',
           '\u001D': 'CTRL+]',
           '\u001E': 'CTRL+^',
           '\u001F': 'CTRL+_',
           '\u007F': 'DEL',
       }
       if '+' in key_str:
           parts = key_str.split('+')
           formatted = []
           for part in parts:
               if part.lower() == 'ctrl':
                   formatted.append('Ctrl')
               elif part.lower() == 'alt':
                   formatted.append('Alt')
               elif part.lower() == 'shift':
                   formatted.append('Shift')
               else:
                   if part in escape_map:
                       formatted.append(escape_map[part])
                   else:
                       try:
                           if part.startswith('\\u'):
                               char = chr(int(part[2:], 16))
                               formatted.append(char)
                           else:
                               formatted.append(part)
                       except:
                           formatted.append(part)
           return '+'.join(formatted)
       if key_str in escape_map:
           return escape_map[key_str]
       if key_str.startswith('\\u'):
           try:
               char = chr(int(key_str[2:], 16))
               return char
           except:
               return key_str
       return key_str

   def setup_ui(self):
       self.root.columnconfigure(0, weight=1)
       self.root.rowconfigure(1, weight=1)
       self.setup_toolbar()
       self.setup_main_content()
       self.setup_bottom_controls()

   def setup_toolbar(self):
       toolbar = ttk.Frame(self.root)
       toolbar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)

       file_btn = ttk.Menubutton(toolbar, text="File", width=5)
       self.file_menu = tk.Menu(file_btn, tearoff=0)
       file_btn['menu'] = self.file_menu

       self.rebuild_file_menu()
       file_btn.pack(side=tk.LEFT, padx=(0, 5))
       ttk.Button(toolbar, text="Save", command=self.save_script).pack(side=tk.LEFT, padx=(0, 5))
       ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
       self.record_btn = ttk.Button(toolbar, text="Record Macro", command=self.start_recording)
       self.record_btn.pack(side=tk.LEFT, padx=(0, 5))
       self.stop_record_btn = ttk.Button(toolbar, text="Stop Recording", command=self.stop_recording, state=tk.DISABLED)
       self.stop_record_btn.pack(side=tk.LEFT, padx=(0, 5))
       ttk.Label(toolbar, text="Repetitions:").pack(side=tk.LEFT, padx=(10, 5))
       self.repetitions = tk.StringVar(value="1")
       ttk.Entry(toolbar, textvariable=self.repetitions, width=5).pack(side=tk.LEFT, padx=(0, 10))
       self.run_btn = ttk.Button(toolbar, text="Run Script", command=self.run_script)
       self.run_btn.pack(side=tk.LEFT, padx=(0, 5))
       self.stop_btn = ttk.Button(toolbar, text="Stop Script", command=self.stop_script_execution, state=tk.DISABLED)
       self.stop_btn.pack(side=tk.LEFT, padx=(0, 5))
       ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
       self.schedule_btn = ttk.Button(toolbar, text="Schedule Script", command=self.show_schedule_dialog)
       self.schedule_btn.pack(side=tk.LEFT, padx=(0, 5))
       ttk.Separator(toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
       ttk.Button(toolbar, text="Preferences", command=self.show_preferences).pack(side=tk.LEFT)

   def show_schedule_dialog(self):
        if not self.actions:
            messagebox.showwarning("Warning", "No actions to schedule")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("Schedule Script")
        self.center_window(dialog, 500, 535)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog)
        main_frame.pack(expand=True, fill='both', padx=20, pady=20)
        
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        ttk.Label(main_frame, text="First Execution Date & Time:", font=("Arial", 11, "bold"), anchor='center').grid(
            row=0, column=0, columnspan=2, pady=(0, 10), sticky='ew')

        date_frame = ttk.Frame(main_frame)
        date_frame.grid(row=1, column=0, columnspan=2, pady=5, sticky='ew')
        date_frame.grid_columnconfigure(0, weight=1)
        
        if TKCALENDAR_AVAILABLE:
            self.date_entry = DateEntry(date_frame, width=12, background='darkblue',
                                      foreground='white', borderwidth=2, date_pattern='yyyy-mm-dd', justify='center')
            self.date_entry.grid(row=0, column=0)
        else:
            self.date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
            ttk.Entry(date_frame, textvariable=self.date_var, width=15, justify='center').grid(row=0, column=0)

        ttk.Label(main_frame, text="Time, 24-hour format (HH:MM):", anchor='center').grid(
            row=2, column=0, columnspan=2, pady=(10, 5), sticky='ew')

        time_frame = ttk.Frame(main_frame)
        time_frame.grid(row=3, column=0, columnspan=2, pady=5, sticky='ew')
        time_frame.grid_columnconfigure(0, weight=1)
        
        time_controls = ttk.Frame(time_frame)
        time_controls.grid(row=0, column=0)

        current_time = datetime.now()
        self.hour_var = tk.StringVar(value=f"{current_time.hour:02d}")
        self.minute_var = tk.StringVar(value=f"{current_time.minute:02d}")

        hour_spin = ttk.Spinbox(time_controls, from_=0, to=23, width=3, textvariable=self.hour_var, format="%02.0f", justify='center')
        hour_spin.pack(side=tk.LEFT)
        ttk.Label(time_controls, text=":").pack(side=tk.LEFT)
        minute_spin = ttk.Spinbox(time_controls, from_=0, to=59, width=3, textvariable=self.minute_var, format="%02.0f", justify='center')
        minute_spin.pack(side=tk.LEFT)

        ttk.Label(main_frame, text="Repetitions per Execution:", font=("Arial", 11, "bold"), anchor='center').grid(
            row=4, column=0, columnspan=2, pady=(20, 5), sticky='ew')
        
        rep_frame = ttk.Frame(main_frame)
        rep_frame.grid(row=5, column=0, columnspan=2, pady=5, sticky='ew')
        rep_frame.grid_columnconfigure(0, weight=1)
        
        self.exec_reps_var = tk.StringVar(value="1")
        ttk.Entry(rep_frame, textvariable=self.exec_reps_var, width=10, justify='center').grid(row=0, column=0)

        ttk.Label(main_frame, text="Total Executions:", font=("Arial", 11, "bold"), anchor='center').grid(
            row=6, column=0, columnspan=2, pady=(20, 5), sticky='ew')
        
        total_frame = ttk.Frame(main_frame)
        total_frame.grid(row=7, column=0, columnspan=2, pady=5, sticky='ew')
        total_frame.grid_columnconfigure(0, weight=1)
        
        self.total_execs_var = tk.StringVar(value="1")
        ttk.Entry(total_frame, textvariable=self.total_execs_var, width=10, justify='center').grid(row=0, column=0)

        ttk.Label(main_frame, text="Repeat Frequency:", font=("Arial", 11, "bold"), anchor='center').grid(
            row=8, column=0, columnspan=2, pady=(20, 5), sticky='ew')
        
        freq_frame = ttk.Frame(main_frame)
        freq_frame.grid(row=9, column=0, columnspan=2, pady=5, sticky='ew')
        freq_frame.grid_columnconfigure(0, weight=1)
        
        self.frequency_var = tk.StringVar(value="Once")
        frequency_combo = ttk.Combobox(freq_frame, textvariable=self.frequency_var, 
                                     values=["Once", "Daily", "Weekly", "Monthly", "Custom Interval"],
                                     state="readonly", width=15, justify='center')
        frequency_combo.grid(row=0, column=0)

        custom_frame = ttk.Frame(main_frame)
        custom_frame.grid(row=10, column=0, columnspan=2, pady=(10, 0), sticky='ew')
        custom_frame.configure(height=30)

        custom_controls = ttk.Frame(custom_frame)
        custom_controls.place(x=1000, y=0)

        ttk.Label(custom_controls, text="Custom Interval (minutes):").pack(side=tk.LEFT, padx=(0, 5))
        self.custom_interval_var = tk.StringVar(value="60")
        self.custom_interval_entry = ttk.Entry(custom_controls, textvariable=self.custom_interval_var, 
                                            width=10, justify='center')
        self.custom_interval_entry.pack(side=tk.LEFT)

        def on_frequency_change(*args):
           if self.frequency_var.get() == "Custom Interval":
               custom_frame.update_idletasks()
               frame_width = custom_frame.winfo_width()
               control_width = custom_controls.winfo_reqwidth()
               center_x = (frame_width - control_width) // 2
               custom_controls.place(x=center_x, y=0)
           else:
               custom_controls.place(x=1000, y=0)
               
        frequency_combo.bind('<<ComboboxSelected>>', on_frequency_change)

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=11, column=0, columnspan=2, pady=(30, 0), sticky='ew')
        button_frame.grid_columnconfigure(0, weight=1)
        
        button_container = ttk.Frame(button_frame)
        button_container.grid(row=0, column=0)

        def save_schedule():
            try:
                if TKCALENDAR_AVAILABLE:
                    date_str = self.date_entry.get()
                else:
                    date_str = self.date_var.get()
                
                time_str = f"{self.hour_var.get().zfill(2)}:{self.minute_var.get().zfill(2)}"
                datetime_str = f"{date_str} {time_str}"
                first_exec = datetime.strptime(datetime_str, '%Y-%m-%d %H:%M')

                if first_exec <= datetime.now():
                    messagebox.showerror("Error", "First execution time must be in the future")
                    return

                schedule_data = {
                    'first_execution': datetime_str,
                    'repetitions_per_execution': int(self.exec_reps_var.get()),
                    'total_executions': int(self.total_execs_var.get()),
                    'frequency': self.frequency_var.get(),
                    'custom_interval': int(self.custom_interval_var.get()) if self.frequency_var.get() == "Custom Interval" else None,
                    'actions': self.actions.copy()
                }

                self.save_schedule(schedule_data)
                self.load_scheduled_tasks()
                messagebox.showinfo("Scheduled", "Script has been scheduled successfully. Keep application active to run according to schedule. Closing the application will cancel the schedule.")
                dialog.destroy()

            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save schedule: {e}")

        def load_schedule_file():
            filename = filedialog.askopenfilename(
                title="Load Schedule",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            if filename:
                try:
                    with open(filename, 'r') as f:
                        schedule_data = json.load(f)
                    
                    first_exec = datetime.strptime(schedule_data['first_execution'], '%Y-%m-%d %H:%M')
                    
                    if TKCALENDAR_AVAILABLE:
                        self.date_entry.set_date(first_exec.date())
                    else:
                        self.date_var.set(first_exec.strftime('%Y-%m-%d'))
                    
                    self.hour_var.set(f"{first_exec.hour:02d}")
                    self.minute_var.set(f"{first_exec.minute:02d}")
                    self.exec_reps_var.set(str(schedule_data['repetitions_per_execution']))
                    self.total_execs_var.set(str(schedule_data['total_executions']))
                    self.frequency_var.set(schedule_data['frequency'])
                    
                    if schedule_data.get('custom_interval'):
                        self.custom_interval_var.set(str(schedule_data['custom_interval']))
                    
                    on_frequency_change()
                    
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load schedule: {e}")

        def save_schedule_file():
            try:
                if TKCALENDAR_AVAILABLE:
                    date_str = self.date_entry.get()
                else:
                    date_str = self.date_var.get()
                
                time_str = f"{self.hour_var.get().zfill(2)}:{self.minute_var.get().zfill(2)}"
                datetime_str = f"{date_str} {time_str}"

                schedule_data = {
                    'first_execution': datetime_str,
                    'repetitions_per_execution': int(self.exec_reps_var.get()),
                    'total_executions': int(self.total_execs_var.get()),
                    'frequency': self.frequency_var.get(),
                    'custom_interval': int(self.custom_interval_var.get()) if self.frequency_var.get() == "Custom Interval" else None,
                    'actions': self.actions.copy()
                }

                filename = filedialog.asksaveasfilename(
                    title="Save Schedule",
                    defaultextension=".json",
                    filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
                )
                if filename:
                    with open(filename, 'w') as f:
                        json.dump(schedule_data, f, indent=2)
                    messagebox.showinfo("Saved", f"Schedule saved to {filename}")

            except ValueError as e:
                messagebox.showerror("Error", f"Invalid input: {e}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save schedule: {e}")

        ttk.Button(button_container, text="Load Schedule", command=load_schedule_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_container, text="Save Schedule", command=save_schedule_file).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_container, text="Start Schedule", command=save_schedule).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_container, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT)
        
   def load_scheduled_tasks(self):
       schedule_data = self.load_schedule()
       if not schedule_data:
           return

       self.scheduled_tasks.clear()
       
       try:
           first_exec = datetime.strptime(schedule_data['first_execution'], '%Y-%m-%d %H:%M')
           frequency = schedule_data['frequency']
           total_executions = schedule_data['total_executions']
           
           current_time = first_exec
           
           for i in range(total_executions):
               if current_time > datetime.now():
                   self.scheduled_tasks.append({
                       'execution_time': current_time,
                       'repetitions': schedule_data['repetitions_per_execution'],
                       'actions': schedule_data['actions']
                   })
               
               if i < total_executions - 1:
                   if frequency == "Daily":
                       current_time += timedelta(days=1)
                   elif frequency == "Weekly":
                       current_time += timedelta(weeks=1)
                   elif frequency == "Monthly":
                       if current_time.month == 12:
                           current_time = current_time.replace(year=current_time.year + 1, month=1)
                       else:
                           current_time = current_time.replace(month=current_time.month + 1)
                   elif frequency == "Custom Interval":
                       current_time += timedelta(minutes=schedule_data['custom_interval'])
           
           self.schedule_active = True
           self.update_window_title()
           
       except Exception as e:
           messagebox.showerror("Error", f"Failed to load scheduled tasks: {e}")

   def start_scheduler(self):
       def scheduler_loop():
           while True:
               if self.schedule_active and self.scheduled_tasks:
                   current_time = datetime.now()
                   tasks_to_remove = []
                   
                   for i, task in enumerate(self.scheduled_tasks):
                       if current_time >= task['execution_time']:
                           if not self.script_running:
                               self.execute_scheduled_task(task)
                           tasks_to_remove.append(i)
                   
                   for i in reversed(tasks_to_remove):
                       self.scheduled_tasks.pop(i)
                   
                   if not self.scheduled_tasks:
                       self.schedule_active = False
               
               time.sleep(30)
       
       self.schedule_thread = threading.Thread(target=scheduler_loop, daemon=True)
       self.schedule_thread.start()
       
       self.load_scheduled_tasks()

   def execute_scheduled_task(self, task):
        def execute():
            original_actions = self.actions.copy()
            self.actions = task['actions']
            
            self.script_running = True
            self.run_btn.configure(state=tk.DISABLED)
            self.stop_btn.configure(state=tk.NORMAL)
            
            try:
                for rep in range(task['repetitions']):
                    if self.stop_script:
                        break
                    for action in self.actions:
                        if self.stop_script:
                            break
                        if action["type"] == "click":
                            time.sleep(action["delay"] / 1000.0)
                            button = action.get("button", "left")
                            if button == "double":
                                pyautogui.doubleClick(action["x"], action["y"])
                            elif button == "scroll_up":
                                scroll_amount = action.get("scroll_amount", 1)
                                pyautogui.scroll(scroll_amount, x=action["x"], y=action["y"])
                            elif button == "scroll_down":
                                scroll_amount = action.get("scroll_amount", 1)
                                pyautogui.scroll(-scroll_amount, x=action["x"], y=action["y"])
                            else:
                                pyautogui.click(action["x"], action["y"], button=button)
                        elif action["type"] == "key":
                            time.sleep(action["delay"] / 1000.0)
                            key = action["key"]
                            if len(key) == 1 and ord(key) < 32:
                                ctrl_keys = {
                                    '\u0003': ['ctrl', 'c'],
                                    '\u0016': ['ctrl', 'v'],
                                    '\u0001': ['ctrl', 'a'],
                                    '\u0018': ['ctrl', 'x'],
                                    '\u001A': ['ctrl', 'z'],
                                }
                                if key in ctrl_keys:
                                    pyautogui.hotkey(*ctrl_keys[key])
                                else:
                                    pyautogui.press('enter' if key == '\n' else key)
                            elif "+" in key:
                                keys = [k.strip().lower() for k in key.split("+")]
                                pyautogui.hotkey(*keys)
                            else:
                                pyautogui.press(key.lower())
                        elif action["type"] == "delay":
                            time.sleep(action["delay"] / 1000.0)
            finally:
                self.actions = original_actions
                self.script_running = False
                self.run_btn.configure(state=tk.NORMAL)
                self.stop_btn.configure(state=tk.DISABLED)
                
                if not self.stop_script:
                    self.root.after(0, lambda: messagebox.showinfo("Scheduled Execution", "Scheduled script execution completed"))
        
        threading.Thread(target=execute, daemon=True).start()
        
   def rebuild_file_menu(self):
       self.file_menu.delete(0, 'end')

       self.file_menu.add_command(label="New Script", command=self.new_script)
       self.file_menu.add_command(label="Load Script", command=self.load_script)
       self.file_menu.add_command(label="Save As", command=self.save_as_script)
       self.file_menu.add_command(label="Delete Script", command=self.delete_script)
       self.file_menu.add_separator()

       self.file_menu.add_command(label="Recent files:")

       prefs = self.load_preferences()
       recent = prefs.get('recent_files', [])

       if not recent:
           self.file_menu.add_command(label="No recent files", state='disabled')
       else:
           for filepath in recent:
               filename = os.path.basename(filepath)
               self.file_menu.add_command(
                   label=filename,
                   command=lambda f=filepath: self.load_recent_file(f)
               )

       self.file_menu.add_separator()
       self.file_menu.add_command(label="Exit", command=self.on_closing)

   def setup_main_content(self):
       main_frame = ttk.Frame(self.root)
       main_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=5)
       main_frame.columnconfigure(0, weight=1)
       main_frame.rowconfigure(0, weight=1)
       self.tree = ttk.Treeview(main_frame, columns=("Type", "Details", "Delay"), show="tree headings")
       self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
       style = ttk.Style()
       style.configure("Treeview", rowheight=25)
       style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
       scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=self.tree.yview)
       scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
       self.tree.configure(yscrollcommand=scrollbar.set)
       self.tree.heading("#0", text="#")
       self.tree.heading("Type", text="Action Type")
       self.tree.heading("Details", text="Details")
       self.tree.heading("Delay", text="Delay (ms)")
       self.tree.column("#0", width=50)
       self.tree.column("Type", width=120)
       self.tree.column("Details", width=400)
       self.tree.column("Delay", width=80)
       self.tree.bind("<Double-1>", self.on_double_click)

   def setup_bottom_controls(self):
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=5, pady=5)
        
        left_frame = ttk.Frame(bottom_frame)
        left_frame.pack(side=tk.LEFT)
        self.editable_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(left_frame, text="Editable", variable=self.editable_var, command=self.toggle_editable).pack(side=tk.LEFT)
        
        middle_frame = ttk.Frame(bottom_frame)
        middle_frame.pack(side=tk.LEFT, padx=20)
        self.move_up_btn = ttk.Button(middle_frame, text="Move Up", command=self.move_up)
        self.move_up_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.move_down_btn = ttk.Button(middle_frame, text="Move Down", command=self.move_down)
        self.move_down_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.delay_adjust_btn = ttk.Button(middle_frame, text="Delay adjustment", command=self.show_delay_adjust_dialog)
        self.delay_adjust_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        right_frame = ttk.Frame(bottom_frame)
        right_frame.pack(side=tk.RIGHT)
        self.copy_btn = ttk.Button(right_frame, text="Copy Action", command=self.copy_action)
        self.copy_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.add_btn = ttk.Button(right_frame, text="Add Action", command=self.show_add_dialog)
        self.add_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_btn = ttk.Button(right_frame, text="Delete Selected", command=self.delete_action)
        self.delete_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.clear_btn = ttk.Button(right_frame, text="Clear All", command=self.clear_actions)
        self.clear_btn.pack(side=tk.LEFT)

   def show_delay_adjust_dialog(self):
       if self.check_no_actions_loaded():
           return
       dialog = tk.Toplevel(self.root)
       dialog.title("Delay Adjustment")
       self.center_window(dialog, 300, 150)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()
       ttk.Label(dialog, text="Enter delay factor (e.g., 2 for 2x, 0.5 for half):").pack(pady=10)
       factor_var = tk.StringVar(value="1")
       ttk.Entry(dialog, textvariable=factor_var, width=10).pack(pady=5)
       def apply_factor():
           try:
               factor = float(factor_var.get())
               for action in self.actions:
                   if "delay" in action:
                       action["delay"] = int(action["delay"] * factor)
               self.update_tree()
               dialog.destroy()
           except ValueError:
               messagebox.showerror("Error", "Please enter a valid number")
       ttk.Button(dialog, text="Apply", command=apply_factor).pack(pady=10)

   def on_double_click(self, event):
       if not self.editable_var.get():
           return
       item = self.tree.selection()[0] if self.tree.selection() else None
       if not item:
           return
       column = self.tree.identify_column(event.x)
       if column in ['#1', '#2', '#3']:
           self.edit_cell(item, column)

   def edit_cell(self, item, column):
        index = int(self.tree.item(item, "text")) - 1
        action = self.actions[index]
        x, y, width, height = self.tree.bbox(item, column)
        entry = tk.Entry(self.tree)
        entry.place(x=x, y=y, width=width, height=height)
        
        if column == '#1':
            if action["type"] == "click":
                entry.insert(0, "Mouse Click")
            elif action["type"] == "key":
                entry.insert(0, "Keyboard Input")
            elif action["type"] == "delay":
                entry.insert(0, "Delay")
        elif column == '#2':
            if action["type"] == "click":
                button_type = action.get("button", "left")
                if button_type == "scroll_up":
                    scroll_amount = action.get("scroll_amount", 1)
                    entry.insert(0, f"({action['x']}, {action['y']}) - scroll up ({scroll_amount} clicks)")
                elif button_type == "scroll_down":
                    scroll_amount = action.get("scroll_amount", 1)
                    entry.insert(0, f"({action['x']}, {action['y']}) - scroll down ({scroll_amount} clicks)")
                else:
                    entry.insert(0, f"({action['x']}, {action['y']}) - {button_type} click")
            elif action["type"] == "key":
                entry.insert(0, self.format_key(action["key"]))
            elif action["type"] == "delay":
                entry.insert(0, f"{action['delay']}ms")
        elif column == '#3':
            if action["type"] != "delay":
                entry.insert(0, str(action["delay"]))
                
        def save_edit():
            new_value = entry.get()
            if column == '#2':
                if action["type"] == "key":
                    action["key"] = new_value
                elif action["type"] == "click":
                    try:
                        parts = new_value.split(" - ")
                        coords = parts[0].strip("()")
                        x, y = map(int, coords.split(", "))
                        
                        if len(parts) > 1:
                            action_desc = parts[1]
                            if "scroll up" in action_desc:
                                button = "scroll_up"
                                if "(" in action_desc and ")" in action_desc:
                                    try:
                                        amount_str = action_desc.split("(")[1].split(")")[0].replace("clicks", "").strip()
                                        action["scroll_amount"] = int(amount_str)
                                    except:
                                        action["scroll_amount"] = 1
                            elif "scroll down" in action_desc:
                                button = "scroll_down"
                                if "(" in action_desc and ")" in action_desc:
                                    try:
                                        amount_str = action_desc.split("(")[1].split(")")[0].replace("clicks", "").strip()
                                        action["scroll_amount"] = int(amount_str)
                                    except:
                                        action["scroll_amount"] = 1
                            else:
                                button = action_desc.split(" ")[0]
                        else:
                            button = "left"
                        
                        action["x"] = x
                        action["y"] = y
                        action["button"] = button
                    except:
                        pass
                elif action["type"] == "delay":
                    try:
                        action["delay"] = int(new_value.replace("ms", ""))
                    except:
                        pass
            elif column == '#3' and action["type"] != "delay":
                try:
                    action["delay"] = int(new_value)
                except:
                    pass
            self.update_tree()
            entry.destroy()
            
        def cancel_edit():
            entry.destroy()
            
        entry.bind('<Return>', lambda e: save_edit())
        entry.bind('<Escape>', lambda e: cancel_edit())
        entry.bind('<FocusOut>', lambda e: save_edit())
        entry.focus_set()
        entry.select_range(0, tk.END)
    
   def move_up(self):
       if self.check_no_actions_loaded():
           return
       if not self.editable_var.get():
           return
       selected = self.tree.selection()
       if not selected:
           self.missing_selection()
           return
       item = selected[0]
       index = int(self.tree.item(item, "text")) - 1
       if index > 0:
           self.actions[index], self.actions[index-1] = self.actions[index-1], self.actions[index]
           self.update_tree()
           new_item = self.tree.get_children()[index-1]
           self.tree.selection_set(new_item)

   def move_down(self):
       if self.check_no_actions_loaded():
           return
       if not self.editable_var.get():
           return
       selected = self.tree.selection()
       if not selected:
           self.missing_selection()
           return
       item = selected[0]
       index = int(self.tree.item(item, "text")) - 1
       if index < len(self.actions) - 1:
           self.actions[index], self.actions[index+1] = self.actions[index+1], self.actions[index]
           self.update_tree()
           new_item = self.tree.get_children()[index+1]
           self.tree.selection_set(new_item)

   def show_preferences(self):
       pref_window = tk.Toplevel(self.root)
       pref_window.title("Preferences")
       self.center_window(pref_window, 300, 150)
       pref_window.resizable(False, False)
       pref_window.transient(self.root)
       pref_window.grab_set()
       ttk.Label(pref_window, text="Theme:").pack(pady=10)
       theme_var = tk.StringVar(value="Dark" if self.dark_mode else "Light")
       theme_frame = ttk.Frame(pref_window)
       theme_frame.pack(pady=10)
       ttk.Radiobutton(theme_frame, text="Light", variable=theme_var, value="Light").pack(side=tk.LEFT, padx=10)
       ttk.Radiobutton(theme_frame, text="Dark", variable=theme_var, value="Dark").pack(side=tk.LEFT, padx=10)
       def apply_theme():
           self.dark_mode = theme_var.get() == "Dark"
           self.apply_theme()
           self.save_preferences()
           pref_window.destroy()
       ttk.Button(pref_window, text="Apply", command=apply_theme).pack(pady=20)

   def apply_theme(self):
       if self.dark_mode:
           sv_ttk.set_theme("dark")
       else:
           sv_ttk.set_theme("light")

   def show_add_dialog(self):
       dialog = tk.Toplevel(self.root)
       dialog.title("Add Action")
       self.center_window(dialog, 400, 300)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()
       ttk.Label(dialog, text="Select Action Type:", font=("Arial", 12)).pack(pady=20)
       button_frame = ttk.Frame(dialog)
       button_frame.pack(pady=10)
       ttk.Button(button_frame, text="Mouse Click", command=lambda: self.show_mouse_dialog(dialog)).pack(side=tk.LEFT, padx=10)
       ttk.Button(button_frame, text="Keyboard Input", command=lambda: self.show_keyboard_dialog(dialog)).pack(side=tk.LEFT, padx=10)
       ttk.Button(button_frame, text="Delay", command=lambda: self.show_delay_dialog(dialog)).pack(side=tk.LEFT, padx=10)

   def show_mouse_dialog(self, parent):
        parent.destroy()
        dialog = tk.Toplevel(self.root)
        dialog.title("Mouse Click Action")
        self.center_window(dialog, 350, 350)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        main_frame = ttk.Frame(dialog)
        main_frame.pack(expand=True, fill='both', padx=20, pady=20)

        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(1, weight=1)

        ttk.Button(main_frame, text="Select Position On Screen",
                   command=lambda: self.select_position_dialog(x_var, y_var, dialog)).grid(
                   row=0, column=0, columnspan=2, pady=(0, 15), sticky='ew')

        ttk.Label(main_frame, text="X Position:").grid(row=1, column=0, sticky=tk.E, padx=(0, 10), pady=5)
        x_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=x_var, width=15).grid(row=1, column=1, sticky=tk.W, pady=5)

        ttk.Label(main_frame, text="Y Position:").grid(row=2, column=0, sticky=tk.E, padx=(0, 10), pady=5)
        y_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=y_var, width=15).grid(row=2, column=1, sticky=tk.W, pady=5)

        ttk.Label(main_frame, text="Click Type:").grid(row=3, column=0, sticky=tk.E, padx=(0, 10), pady=5)
        click_type = ttk.Combobox(main_frame, values=["Left Click", "Double Click", "Right Click", "Middle Click", "Scroll Up", "Scroll Down"],
                                  state="readonly", width=12)
        click_type.grid(row=3, column=1, sticky=tk.W, pady=5)
        click_type.set("Left Click")

        scroll_label = ttk.Label(main_frame, text="Scroll Amount:")
        scroll_var = tk.StringVar(value="3")
        scroll_entry = ttk.Entry(main_frame, textvariable=scroll_var, width=15)
        
        def on_click_type_change(*args):
            if click_type.get() in ["Scroll Up", "Scroll Down"]:
                scroll_label.grid(row=4, column=0, sticky=tk.E, padx=(0, 10), pady=5)
                scroll_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
                delay_label.grid(row=5, column=0, sticky=tk.E, padx=(0, 10), pady=5)
                delay_entry.grid(row=5, column=1, sticky=tk.W, pady=5)
                add_button.grid(row=6, column=0, columnspan=2, pady=(15, 0), sticky='ew')
            else:
                scroll_label.grid_remove()
                scroll_entry.grid_remove()
                delay_label.grid(row=4, column=0, sticky=tk.E, padx=(0, 10), pady=5)
                delay_entry.grid(row=4, column=1, sticky=tk.W, pady=5)
                add_button.grid(row=5, column=0, columnspan=2, pady=(15, 0), sticky='ew')

        click_type.bind('<<ComboboxSelected>>', on_click_type_change)

        delay_label = ttk.Label(main_frame, text="Delay (ms):")
        delay_label.grid(row=4, column=0, sticky=tk.E, padx=(0, 10), pady=5)
        delay_var = tk.StringVar(value="100")
        delay_entry = ttk.Entry(main_frame, textvariable=delay_var, width=15)
        delay_entry.grid(row=4, column=1, sticky=tk.W, pady=5)

        def add_mouse_action():
            try:
                x = int(x_var.get()) if x_var.get() else 0
                y = int(y_var.get()) if y_var.get() else 0
                delay = int(delay_var.get()) if delay_var.get() else 100
                
                click_map = {
                    "Left Click": "left", 
                    "Double Click": "double", 
                    "Right Click": "right", 
                    "Middle Click": "middle",
                    "Scroll Up": "scroll_up",
                    "Scroll Down": "scroll_down"
                }
                
                action = {"type": "click", "x": x, "y": y, "button": click_map[click_type.get()], "delay": delay}
                
                if click_type.get() in ["Scroll Up", "Scroll Down"]:
                    scroll_amount = int(scroll_var.get()) if scroll_var.get() else 3
                    action["scroll_amount"] = scroll_amount
                
                self.actions.append(action)
                self.unsaved_changes = True
                self.update_tree(scroll_to_bottom=True)
                self.update_window_title()
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Error", "Invalid values entered")

        add_button = ttk.Button(main_frame, text="Add", command=add_mouse_action)
        add_button.grid(row=5, column=0, columnspan=2, pady=(15, 0), sticky='ew')
    
   def show_keyboard_dialog(self, parent):
       parent.destroy()
       dialog = tk.Toplevel(self.root)
       dialog.title("Keyboard Input Action")
       self.center_window(dialog, 400, 300)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()
       main_frame = ttk.Frame(dialog)
       main_frame.pack(expand=True, fill='both', padx=20, pady=20)
       main_frame.grid_columnconfigure(0, weight=1)
       ttk.Button(
           main_frame, text="Click here to capture key input",
           command=lambda: self.capture_key_input(key_var, dialog)
       ).grid(row=0, column=0, pady=(0, 15), sticky='ew')
       ttk.Label(
           main_frame, text="Key/Text (supports ctrl+f, alt+tab, etc.):"
       ).grid(row=1, column=0, pady=5, sticky='')
       key_var = tk.StringVar()
       ttk.Entry(
           main_frame, textvariable=key_var, width=30
       ).grid(row=2, column=0, pady=5, sticky='ew')
       ttk.Label(main_frame, text="Delay (ms):").grid(
           row=3, column=0, pady=(10, 5), sticky='')
       delay_var = tk.StringVar(value="100")
       ttk.Entry(
           main_frame, textvariable=delay_var, width=15
       ).grid(row=4, column=0, pady=5, sticky='ew')
       def add_keyboard_action():
           key = key_var.get()
           if not key:
               messagebox.showerror("Error", "Key input cannot be empty")
               return
           delay = int(delay_var.get()) if delay_var.get() else 100
           action = {"type": "key", "key": key, "delay": delay}
           self.actions.append(action)
           self.update_tree()
           self.unsaved_changes = True
           self.update_tree(scroll_to_bottom=True)
           self.update_window_title()
           dialog.destroy()
       ttk.Button(
           main_frame, text="Add", command=add_keyboard_action
       ).grid(row=5, column=0, pady=(15, 0), sticky='ew')

   def show_delay_dialog(self, parent):
       parent.destroy()
       dialog = tk.Toplevel(self.root)
       dialog.title("Delay Action")
       self.center_window(dialog, 250, 150)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()
       ttk.Label(dialog, text="Delay (ms):").grid(row=0, column=0, sticky=tk.W, padx=10, pady=20)
       delay_var = tk.StringVar(value="1000")
       ttk.Entry(dialog, textvariable=delay_var, width=10).grid(row=0, column=1, padx=10, pady=20)
       def add_delay_action():
           try:
               delay = int(delay_var.get())
               action = {"type": "delay", "delay": delay}
               self.actions.append(action)
               self.unsaved_changes = True
               self.update_tree(scroll_to_bottom=True)
               self.update_window_title()
               self.update_tree()
               dialog.destroy()
           except ValueError:
               messagebox.showerror("Error", "Invalid delay value")
       ttk.Button(dialog, text="Add", command=add_delay_action).grid(row=1, column=0, columnspan=2, pady=20)

   def capture_key_input(self, key_var, dialog):
       messagebox.showinfo("Key Capture", "Click OK then enter the key combination you want to capture.")
       def on_key_combination(e):
           keys = []
           if e.state & 0x4:
               keys.append("ctrl")
           if e.state & 0x8:
               keys.append("alt")
           if e.state & 0x1:
               keys.append("shift")
           if e.keysym not in ['Control_L', 'Control_R', 'Alt_L', 'Alt_R', 'Shift_L', 'Shift_R']:
               keys.append(e.keysym.lower())
           if keys:
               key_var.set("+".join(keys))
               dialog.unbind("<KeyPress>")
       dialog.bind("<KeyPress>", on_key_combination)
       dialog.focus_set()

   def select_position_dialog(self, x_var, y_var, dialog):
        dialog.withdraw()
        messagebox.showinfo("Position Selection", "Click anywhere on screen to capture position. Press ESC to cancel.")

        def restore_windows():
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.root.attributes('-topmost', True)
            
            dialog.deiconify()
            dialog.lift()
            dialog.focus_force()
            dialog.attributes('-topmost', True)
            dialog.grab_set()
            
            self.root.after(100, lambda: (
                self.root.attributes('-topmost', False),
                dialog.attributes('-topmost', False)
            ))

        def on_click(x, y, button, pressed):
            if pressed:
                x_var.set(str(x))
                y_var.set(str(y))
                self.root.after(50, restore_windows)
                return False

        def on_key(key):
            if key == pynput_keyboard.Key.esc:
                self.root.after(50, restore_windows)
                return False

        mouse_listener = mouse.Listener(on_click=on_click)
        key_listener = pynput_keyboard.Listener(on_press=on_key)
        mouse_listener.start()
        key_listener.start()
        
   def setup_empty_state(self):
       if hasattr(self, 'empty_frame'):
           self.empty_frame.destroy()

       if not self.actions:
           self.empty_frame = ttk.Frame(self.tree.master)
           self.empty_frame.place(relx=0.5, rely=0.4, anchor='center')

           msg_label = ttk.Label(self.empty_frame,
                                text="No actions loaded. Either record or add actions, or load a script.",
                                font=("Arial", 11),
                                justify='center')
           msg_label.pack(pady=(0, 20))

           button_frame = ttk.Frame(self.empty_frame)
           button_frame.pack()

           ttk.Button(button_frame, text="Add Action",
                     command=self.show_add_dialog).pack(side=tk.LEFT, padx=5)
           ttk.Button(button_frame, text="Record Macro",
                     command=self.start_recording).pack(side=tk.LEFT, padx=5)
           ttk.Button(button_frame, text="Load Script",
                     command=self.load_script).pack(side=tk.LEFT, padx=5)
       elif hasattr(self, 'empty_frame'):
           self.empty_frame.destroy()

   def update_window_title(self):
       base_title = "NickClick - Auto Keyboard and Mouse Clicker"
       
       if hasattr(self, 'current_file'):
           filename = os.path.basename(self.current_file)
           file_status = filename
       elif self.actions:
           file_status = "New script"
       else:
           file_status = ""
       
       title_parts = [base_title]
       
       if file_status:
           title_parts.append(f" - {file_status}")
       
       if self.unsaved_changes:
           title_parts.append("*")
       
       status_indicators = []
       
       if self.script_running:
           status_indicators.append("RUNNING")
       
       if self.schedule_active and self.scheduled_tasks:
           status_indicators.append("SCHEDULED")
       
       if status_indicators:
           title_parts.append(f" [{' | '.join(status_indicators)}]")
       
       self.root.title(''.join(title_parts))
       
   def update_tree(self, scroll_to_bottom=False):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, action in enumerate(self.actions):
            if action["type"] == "click":
                button_type = action.get("button", "left")
                if button_type == "scroll_up":
                    scroll_amount = action.get("scroll_amount", 1)
                    details = f"({action['x']}, {action['y']}) - scroll up ({scroll_amount} clicks)"
                elif button_type == "scroll_down":
                    scroll_amount = action.get("scroll_amount", 1)
                    details = f"({action['x']}, {action['y']}) - scroll down ({scroll_amount} clicks)"
                else:
                    details = f"({action['x']}, {action['y']}) - {button_type} click"
                self.tree.insert("", "end", text=str(i+1), values=("Mouse Click", details, action["delay"]))
            elif action["type"] == "key":
                self.tree.insert("", "end", text=str(i+1), values=("Keyboard Input", self.format_key(action["key"]), action["delay"]))
            elif action["type"] == "delay":
                self.tree.insert("", "end", text=str(i+1), values=("Delay", f"{action['delay']}ms", ""))

        if scroll_to_bottom and self.actions:
            children = self.tree.get_children()
            if children:
                self.tree.see(children[-1])

        self.setup_empty_state()
    
   def delete_action(self):
       if self.check_no_actions_loaded():
           return
       if not self.editable_var.get():
           return
       selected = self.tree.selection()
       if not selected:
           self.missing_selection()
           return
       result = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected action?")
       if result:
           item = selected[0]
           index = int(self.tree.item(item, "text")) - 1
           del self.actions[index]
           self.update_tree()

   def missing_selection(self):
       messagebox.showwarning("No Selection", "No line selected.")

   def check_no_actions_loaded(self):
       if not self.actions:
           messagebox.showwarning("No Script", "No script is loaded.")
           return True
       return False

   def clear_actions(self):
       if self.check_no_actions_loaded():
           return
       if not self.editable_var.get():
           return
       if self.actions:
           result = messagebox.askyesno("Confirm Clear", "Are you sure you want to clear all actions?")
           if result:
               self.actions.clear()
               self.clear_schedule()
               self.update_tree()

   def toggle_editable(self):
       if self.check_no_actions_loaded():
           self.editable_var.set(not self.editable_var.get())
           return
       if self.editable and not self.editable_var.get():
           self.show_lock_dialog()
       elif not self.editable and self.editable_var.get():
           if self.is_script_locked():
               self.unlock_script()
           else:
               self.update_ui_editable_state()
       else:
           self.update_ui_editable_state()

   def is_script_locked(self):
       if not hasattr(self, 'current_file'):
           return False

       try:
           with open(self.current_file, 'r') as f:
               data = json.load(f)

           if isinstance(data, dict) and "actions" in data:
               return data.get("locked", False)
       except:
           pass

       return False

   def show_lock_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Lock Script")
        self.center_window(dialog, 400, 200)
        dialog.resizable(False, False)
        dialog.transient(self.root)
        dialog.grab_set()

        def on_dialog_close():
            dont_lock()
        
        dialog.protocol("WM_DELETE_WINDOW", on_dialog_close)

        ttk.Label(dialog, text="Would you like to lock the script against future edits?").pack(pady=20)

        def lock_with_password():
            dialog.destroy()
            self.lock_script_with_password()

        def lock_without_password():
            dialog.destroy()
            self.lock_script_without_password()

        def dont_lock():
            dialog.destroy()
            self.update_ui_editable_state()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=10)

        ttk.Button(button_frame, text="Yes, with password protection", command=lock_with_password).pack(pady=5)
        ttk.Button(button_frame, text="Yes, with no protection", command=lock_without_password).pack(pady=5)
        ttk.Button(button_frame, text="No", command=dont_lock).pack(pady=5)
               
   def hash_password(self, password):
       return hashlib.sha256(password.encode()).hexdigest()

   def show_password_creation_dialog(self):
       dialog = tk.Toplevel(self.root)
       dialog.title("Set Password")
       self.center_window(dialog, 350, 200)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()

       ttk.Label(dialog, text="Enter password:").grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
       password_var = tk.StringVar()
       password_entry = ttk.Entry(dialog, textvariable=password_var, show='*', width=25)
       password_entry.grid(row=0, column=1, padx=10, pady=5)

       ttk.Label(dialog, text="Confirm password:").grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
       confirm_var = tk.StringVar()
       confirm_entry = ttk.Entry(dialog, textvariable=confirm_var, show='*', width=25)
       confirm_entry.grid(row=1, column=1, padx=10, pady=5)

       show_password_var = tk.BooleanVar()
       def toggle_password_visibility():
           show_char = '' if show_password_var.get() else '*'
           password_entry.config(show=show_char)
           confirm_entry.config(show=show_char)

       ttk.Checkbutton(dialog, text="Show passwords", variable=show_password_var,
                      command=toggle_password_visibility).grid(row=2, column=0, columnspan=2, pady=10)

       button_frame = ttk.Frame(dialog)
       button_frame.grid(row=3, column=0, columnspan=2, pady=20)

       result = {'password': None}

       def confirm_password():
           pwd1 = password_var.get()
           pwd2 = confirm_var.get()

           if not pwd1:
               messagebox.showerror("Error", "Password cannot be empty")
               return

           if pwd1 != pwd2:
               messagebox.showerror("Error", "Passwords do not match")
               return

           result['password'] = pwd1
           dialog.destroy()

       def cancel():
           dialog.destroy()

       ttk.Button(button_frame, text="OK", command=confirm_password).pack(side=tk.LEFT, padx=10)
       ttk.Button(button_frame, text="Cancel", command=cancel).pack(side=tk.LEFT, padx=10)

       password_entry.focus()
       dialog.wait_window()

       return result['password']

   def lock_script_with_password(self):
       password = self.show_password_creation_dialog()
       if password:
           self.ensure_script_saved()
           if hasattr(self, 'current_file'):
               script_data = {
                   "actions": self.actions,
                   "locked": True,
                   "password_hash": self.hash_password(password)
               }
               with open(self.current_file, 'w') as f:
                   json.dump(script_data, f, indent=2)
               self.update_ui_editable_state()
               messagebox.showinfo("Locked", "Script has been locked with password protection.")

   def lock_script_without_password(self):
       self.ensure_script_saved()
       if hasattr(self, 'current_file'):
           script_data = {
               "actions": self.actions,
               "locked": True
           }
           with open(self.current_file, 'w') as f:
               json.dump(script_data, f, indent=2)
           self.update_ui_editable_state()
           messagebox.showinfo("Locked", "Script has been locked.")

   def ensure_script_saved(self):
       if not hasattr(self, 'current_file'):
           self.save_as_script()

   def unlock_script(self):
       if not hasattr(self, 'current_file'):
           self.update_ui_editable_state()
           return
       try:
           with open(self.current_file, 'r') as f:
               data = json.load(f)
           if isinstance(data, dict) and "actions" in data:
               if data.get("locked"):
                   if "password_hash" in data:
                       dialog = tk.Toplevel(self.root)
                       dialog.title("Unlock Script")
                       dialog.minsize(250, 100)
                       dialog.resizable(False, False)
                       dialog.transient(self.root)
                       dialog.grab_set()
                       dialog.update_idletasks()
                       x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_reqwidth() // 2)
                       y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_reqheight() // 2)
                       dialog.geometry("+{}+{}".format(x, y))
                       ttk.Label(dialog, text="Enter password to unlock script:").pack(pady=10)
                       password_var = tk.StringVar()
                       password_entry = ttk.Entry(dialog, textvariable=password_var, show='*')
                       password_entry.pack(pady=5)
                       def on_ok():
                           nonlocal password
                           password = password_var.get()
                           dialog.destroy()
                       ttk.Button(dialog, text="OK", command=on_ok).pack(pady=10)
                       dialog.bind('<Return>', lambda e: on_ok())
                       password = None
                       dialog.wait_window()
                       if password and self.hash_password(password) == data["password_hash"]:
                           script_data = {
                               "actions": data["actions"],
                               "locked": False
                           }
                           with open(self.current_file, 'w') as f:
                               json.dump(script_data, f, indent=2)
                           self.update_ui_editable_state()
                           messagebox.showinfo("Unlocked", "Script has been unlocked.")
                       else:
                           messagebox.showerror("Access Denied", "Incorrect password." if password else "Unlock cancelled.")
                           self.editable_var.set(False)
                           return
                   else:
                       script_data = {
                           "actions": data["actions"],
                           "locked": False
                       }
                       with open(self.current_file, 'w') as f:
                           json.dump(script_data, f, indent=2)
                       self.update_ui_editable_state()
                       messagebox.showinfo("Unlocked", "Script has been unlocked.")
               else:
                   self.update_ui_editable_state()
       except Exception as e:
           messagebox.showerror("Error", f"Failed to unlock script: {str(e)}")
           self.editable_var.set(False)
   def copy_action(self):
        if self.check_no_actions_loaded():
            return
        if not self.editable_var.get():
            return
        selected = self.tree.selection()
        if not selected:
            self.missing_selection()
            return
        
        item = selected[0]
        index = int(self.tree.item(item, "text")) - 1
        action_to_copy = self.actions[index].copy()
        
        self.actions.insert(index + 1, action_to_copy)
        self.unsaved_changes = True
        self.update_tree()
        self.update_window_title()
        
        new_item = self.tree.get_children()[index + 1]
        self.tree.selection_set(new_item)
        self.tree.see(new_item)
    
   def update_ui_editable_state(self):
        self.editable = self.editable_var.get()
        state = tk.NORMAL if self.editable else tk.DISABLED
        self.copy_btn.configure(state=state)
        self.add_btn.configure(state=state)
        if not self.recording:
            self.record_btn.configure(state=state)
        self.delete_btn.configure(state=state)
        self.clear_btn.configure(state=state)
        self.move_up_btn.configure(state=state)
        self.move_down_btn.configure(state=state)
        self.delay_adjust_btn.configure(state=state)

   def setup_recording(self):
       self.recorded_actions = []
       self.last_time = 0

   def start_recording(self):
       if not self.editable_var.get():
           return

       self.recording_cancelled = False

       dialog = tk.Toplevel(self.root)
       dialog.title("Recording Setup")
       self.center_window(dialog, 400, 200)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()
       ttk.Label(dialog, text="Press any keyboard button to start recording.\n \nWhen finished, press the same button again\nto save your actions.\n", justify=tk.CENTER).pack(pady=20)
       def cancel_recording():
           self.recording_cancelled = True
           dialog.destroy()

       ttk.Button(dialog, text="CANCEL", command=cancel_recording).pack(pady=10)

       def wait_for_key():
           def on_any_key(key):
               if not self.recording_cancelled:
                   self.record_key = key
                   dialog.destroy()
                   self.begin_recording()
               return False
           key_listener = pynput_keyboard.Listener(on_press=on_any_key)
           key_listener.start()

       threading.Thread(target=wait_for_key, daemon=True).start()

   def begin_recording(self):
        self.recording = True
        self.recorded_actions.clear()
        self.last_time = time.time()
        self.record_btn.configure(state=tk.DISABLED)
        self.stop_record_btn.configure(state=tk.NORMAL)
        messagebox.showinfo("Recording", f"Recording started! Press {self.record_key} again to stop.")
        self.mouse_listener = mouse.Listener(on_click=self.on_record_click, on_scroll=self.on_record_scroll)
        self.key_listener = pynput_keyboard.Listener(on_press=self.on_record_key)
        self.mouse_listener.start()
        self.key_listener.start()

   def on_record_click(self, x, y, button, pressed):
        if self.recording and pressed:
            current_time = time.time()
            delay = int((current_time - self.last_time) * 1000) if self.last_time > 0 else 0
            button_name = str(button).replace('Button.', '').lower()
            self.recorded_actions.append({"type": "click", "x": x, "y": y, "button": button_name, "delay": delay})
            self.last_time = current_time


   def on_record_scroll(self, x, y, dx, dy):
        if self.recording:
            current_time = time.time()
            delay = int((current_time - self.last_time) * 1000) if self.last_time > 0 else 0
            
            if dy > 0:
                self.recorded_actions.append({"type": "click", "x": x, "y": y, "button": "scroll_up", "delay": delay, "scroll_amount": abs(dy)})
            elif dy < 0:
                self.recorded_actions.append({"type": "click", "x": x, "y": y, "button": "scroll_down", "delay": delay, "scroll_amount": abs(dy)})
            
            self.last_time = current_time

   def on_record_key(self, key):
       if self.recording:
           if key == self.record_key:
               self.stop_recording()
               return
           current_time = time.time()
           delay = int((current_time - self.last_time) * 1000) if self.last_time > 0 else 0
           try:
               if hasattr(key, 'char') and key.char:
                   if ord(key.char) < 32:
                       ctrl_map = {
                           '\u0001': 'ctrl+a',
                           '\u0002': 'ctrl+b',
                           '\u0003': 'ctrl+c',
                           '\u0004': 'ctrl+d',
                           '\u0005': 'ctrl+e',
                           '\u0006': 'ctrl+f',
                           '\u0007': 'ctrl+g',
                           '\u0008': 'backspace',
                           '\u0009': 'tab',
                           '\u000A': 'enter',
                           '\u000D': 'enter',
                           '\u0011': 'ctrl+q',
                           '\u0012': 'ctrl+r',
                           '\u0013': 'ctrl+s',
                           '\u0014': 'ctrl+t',
                           '\u0015': 'ctrl+u',
                           '\u0016': 'ctrl+v',
                           '\u0017': 'ctrl+w',
                           '\u0018': 'ctrl+x',
                           '\u0019': 'ctrl+y',
                           '\u001A': 'ctrl+z',
                           '\u001B': 'esc',
                       }
                       mapped_key = ctrl_map.get(key.char, key.char)
                       self.recorded_actions.append({"type": "key", "key": mapped_key, "delay": delay})
                   else:
                       self.recorded_actions.append({"type": "key", "key": key.char, "delay": delay})
               else:
                   key_name = str(key).replace('Key.', '').lower()
                   self.recorded_actions.append({"type": "key", "key": key_name, "delay": delay})
           except AttributeError:
               key_name = str(key).replace('Key.', '').lower()
               self.recorded_actions.append({"type": "key", "key": key_name, "delay": delay})
           self.last_time = current_time

   def stop_recording(self):
       if not self.recording:
           return
       self.recording = False
       if hasattr(self, 'mouse_listener'):
           self.mouse_listener.stop()
       if hasattr(self, 'key_listener'):
           self.key_listener.stop()
       self.record_btn.configure(state=tk.NORMAL if self.editable_var.get() else tk.DISABLED)
       self.stop_record_btn.configure(state=tk.DISABLED)
       self.actions.extend(self.recorded_actions)
       if self.recorded_actions:
           self.unsaved_changes = True
       self.update_tree(scroll_to_bottom=True)
       self.update_window_title()
       messagebox.showinfo("Recording Complete", f"Recorded {len(self.recorded_actions)} actions")

   def show_completion_notification(self):
       import winsound
       winsound.MessageBeep()
       messagebox.showinfo("Complete", "Script finished executing")

   def run_script(self):
      if not self.actions:
          messagebox.showwarning("Warning", "No actions to execute")
          return
      if self.script_running:
          return
      try:
          reps = int(self.repetitions.get())
      except ValueError:
          reps = 1
      self.show_hotkey_dialog(reps)

   def show_hotkey_dialog(self, reps):
       dialog = tk.Toplevel(self.root)
       dialog.title("Script Execution Setup")
       self.center_window(dialog, 500, 300)
       dialog.resizable(False, False)
       dialog.transient(self.root)
       dialog.grab_set()
       frame = ttk.Frame(dialog)
       frame.pack(expand=True)
       msg_text = f"You are initiating {reps} execute{'s' if reps != 1 else ''} of your script.\n\nPress any key within 10 seconds to set your abort script hotkey.\n\nStandard hotkey: F12.\n\nCurrently selected abort hotkey:"
       ttk.Label(frame, text=msg_text, justify="center", anchor="center").pack(pady=10)
       self.selected_hotkey = "F12"
       self.hotkey_label = ttk.Label(frame, text=self.selected_hotkey, font=("Arial", 12, "bold"), anchor="center")
       self.hotkey_label.pack()
       button_frame = ttk.Frame(dialog)
       button_frame.pack(pady=20)
       self.countdown_var = tk.StringVar(value="START NOW (10)")
       self.start_btn = ttk.Button(button_frame, textvariable=self.countdown_var, command=lambda: self.start_script_execution(dialog, reps))
       self.start_btn.pack(side=tk.LEFT, padx=10)
       ttk.Button(button_frame, text="CANCEL", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
       self.countdown = 10
       self.countdown_active = True
       def update_countdown():
           if self.countdown_active and self.countdown > 0:
               self.countdown_var.set(f"START NOW ({self.countdown})")
               self.countdown -= 1
               dialog.after(1000, update_countdown)
           elif self.countdown_active:
               self.start_script_execution(dialog, reps)
       def on_key_press(event):
           key_name = event.keysym
           if key_name in ['Control_L', 'Control_R', 'Alt_L', 'Alt_R', 'Shift_L', 'Shift_R']:
               return
           if key_name == 'Return':
               key_name = 'Enter'
           elif key_name == 'BackSpace':
               key_name = 'Backspace'
           elif key_name.startswith('F') and key_name[1:].isdigit():
               key_name = key_name.upper()
           elif len(key_name) == 1:
               key_name = key_name.upper()
           self.selected_hotkey = key_name
           self.hotkey_label.config(text=self.selected_hotkey)
       dialog.bind('<KeyPress>', on_key_press)
       dialog.focus_set()
       update_countdown()

   def start_script_execution(self, dialog, reps):
        self.countdown_active = False
        dialog.destroy()
        self.script_running = True
        self.update_window_title()
        self.stop_script = False
        self.run_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        
        def execute():
            def monitor_hotkey():
                while self.script_running:
                    if keyboard.is_pressed(self.selected_hotkey.lower()):
                        self.stop_script = True
                        break
                    time.sleep(0.1)
            
            hotkey_thread = threading.Thread(target=monitor_hotkey, daemon=True)
            hotkey_thread.start()
            
            rep_count = 0
            while (reps == 0 or rep_count < reps) and not self.stop_script:
                for action in self.actions:
                    if self.stop_script:
                        break
                    if action["type"] == "click":
                        time.sleep(action["delay"] / 1000.0)
                        button = action.get("button", "left")
                        if button == "double":
                            pyautogui.doubleClick(action["x"], action["y"])
                        elif button == "scroll_up":
                            scroll_amount = action.get("scroll_amount", 1)
                            pyautogui.scroll(scroll_amount, x=action["x"], y=action["y"])
                        elif button == "scroll_down":
                            scroll_amount = action.get("scroll_amount", 1)
                            pyautogui.scroll(-scroll_amount, x=action["x"], y=action["y"])
                        else:
                            pyautogui.click(action["x"], action["y"], button=button)
                    elif action["type"] == "key":
                        time.sleep(action["delay"] / 1000.0)
                        key = action["key"]
                        if len(key) == 1 and ord(key) < 32:
                            ctrl_keys = {
                                '\u0003': ['ctrl', 'c'],
                                '\u0016': ['ctrl', 'v'],
                                '\u0001': ['ctrl', 'a'],
                                '\u0018': ['ctrl', 'x'],
                                '\u001A': ['ctrl', 'z'],
                            }
                            if key in ctrl_keys:
                                pyautogui.hotkey(*ctrl_keys[key])
                            else:
                                pyautogui.press('enter' if key == '\n' else key)
                        elif "+" in key:
                            keys = [k.strip().lower() for k in key.split("+")]
                            pyautogui.hotkey(*keys)
                        else:
                            pyautogui.press(key.lower())
                    elif action["type"] == "delay":
                        time.sleep(action["delay"] / 1000.0)
                rep_count += 1
            
            self.script_running = False
            self.update_window_title()
            self.run_btn.configure(state=tk.NORMAL)
            self.stop_btn.configure(state=tk.DISABLED)
            
            if self.stop_script:
                dialog = tk.Toplevel(self.root)
                dialog.title("Script Stopped")
                dialog.resizable(False, False)
                self.center_window(dialog, 300, 150)
                dialog.transient(self.root)
                dialog.grab_set()
                dialog.focus_force()
                frame = ttk.Frame(dialog)
                frame.pack(expand=True)
                ttk.Label(frame, text="Script execution stopped", font=("Arial", 12), justify="center").pack(pady=20)
                ttk.Button(frame, text="OK", command=dialog.destroy).pack(pady=10)
            
            if not self.stop_script:
                self.show_completion_notification()
        
        threading.Thread(target=execute, daemon=True).start()
    
   def stop_script_execution(self):
       self.stop_script = True

   def save_script(self):
       if hasattr(self, 'current_file'):
           with open(self.current_file, 'w') as f:
               json.dump(self.actions, f, indent=2)
           messagebox.showinfo("Saved", f"Script saved to {self.current_file}")
           self.unsaved_changes = False
           self.update_window_title()
       else:
           self.save_as_script()

   def save_as_script(self):
       filename = filedialog.asksaveasfilename(
           defaultextension=".json",
           filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
       )
       if filename:
           with open(filename, 'w') as f:
               json.dump(self.actions, f, indent=2)
           self.current_file = filename
           self.add_to_recent(filename)
           messagebox.showinfo("Saved", f"Script saved to {filename}")
           self.unsaved_changes = False
           self.update_window_title()

   def prompt_unsaved_changes(self):
       if self.unsaved_changes:
           result = messagebox.askyesnocancel(
               "Unsaved Changes",
               "You have unsaved changes. Would you like to save them?"
           )
           if result is None:
               return False
           if result:
               if hasattr(self, 'current_file'):
                   self.save_script()
               else:
                   self.save_as_script()
               return True
           else:
               return True
       return True

   def run(self):
       self.root.mainloop()

if __name__ == "__main__":
   app = AutomationGUI()
   app.run()
