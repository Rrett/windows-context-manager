import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes, Structure
import win32gui
import win32con
import win32process
import win32api
from comtypes import GUID
import psutil
from collections import OrderedDict

# Windows API constants
DWMWA_CLOAKED = 14
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000

# Property key for device friendly name
class PROPERTYKEY(Structure):
    _fields_ = [
        ('fmtid', GUID),
        ('pid', wintypes.DWORD),
    ]

PKEY_Device_FriendlyName = PROPERTYKEY()
PKEY_Device_FriendlyName.fmtid = GUID('{a45c254e-df1c-4efd-8020-67d146a850e0}')
PKEY_Device_FriendlyName.pid = 14


class WindowManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Window Manager")
        
        # Smaller, sleeker window
        self.root.geometry("420x550")
        self.root.minsize(380, 400)
        self.root.resizable(True, True)
        
        # Dark theme colors
        self.colors = {
            'bg': '#1a1a1a',
            'fg': '#e0e0e0',
            'accent': '#3d5afe',
            'accent_hover': '#536dfe',
            'card': '#252525',
            'card_hover': '#2d2d2d',
            'card_selected': '#1a2a4a',
            'border': '#333333',
            'success': '#4caf50',
            'muted': '#888888',
            'checkbox_checked': '#3d5afe',
            'checkbox_unchecked': '#444444'
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Pin to top variable
        self.pin_to_top = tk.BooleanVar(value=False)
        
        # Track selected windows in order
        self.selected_windows = OrderedDict()
        self.window_checkboxes = {}
        self.window_cards = {}  # Store card references for styling
        self.windows_list = []
        
        # Get monitors
        self.monitors = self.get_monitors()
        
        self.setup_styles()
        self.setup_ui()
        self.refresh_windows()
    
    def temporarily_pin(self):
        """Temporarily pin window to top during actions"""
        was_pinned = self.pin_to_top.get()
        if not was_pinned:
            self.root.attributes('-topmost', True)
        return was_pinned
    
    def restore_pin_state(self, was_pinned):
        """Restore pin state after action"""
        if not was_pinned:
            self.root.after(500, lambda: self.root.attributes('-topmost', self.pin_to_top.get()))
    
    def is_window_maximized(self, hwnd):
        """Check if a window is maximized"""
        try:
            placement = win32gui.GetWindowPlacement(hwnd)
            return placement[1] == win32con.SW_SHOWMAXIMIZED
        except:
            return False
    
    def is_window_minimized(self, hwnd):
        """Check if a window is minimized"""
        try:
            return win32gui.IsIconic(hwnd)
        except:
            return False
        
    def setup_styles(self):
        """Configure ttk styles for sleek appearance"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Frame styles
        style.configure('TFrame', background=self.colors['bg'])
        style.configure('Card.TFrame', background=self.colors['card'])
        
        # Label styles
        style.configure('TLabel', 
                       background=self.colors['bg'], 
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 9))
        style.configure('Title.TLabel',
                       font=('Segoe UI', 14, 'bold'),
                       foreground=self.colors['fg'])
        style.configure('Muted.TLabel',
                       foreground=self.colors['muted'],
                       font=('Segoe UI', 8))
        style.configure('Card.TLabel',
                       background=self.colors['card'])
        
        # Button styles
        style.configure('TButton',
                       padding=(12, 6),
                       font=('Segoe UI', 9),
                       background=self.colors['card'],
                       foreground=self.colors['fg'])
        style.map('TButton',
                 background=[('active', self.colors['card_hover'])])
        
        style.configure('Accent.TButton',
                       background=self.colors['accent'],
                       foreground='white')
        style.map('Accent.TButton',
                 background=[('active', self.colors['accent_hover'])])
        
        style.configure('Small.TButton',
                       padding=(6, 3),
                       font=('Segoe UI', 8))
        
        # Checkbutton styles
        style.configure('TCheckbutton',
                       background=self.colors['bg'],
                       foreground=self.colors['fg'],
                       font=('Segoe UI', 9))
        style.configure('Card.TCheckbutton',
                       background=self.colors['card'])
        style.map('TCheckbutton',
                 background=[('active', self.colors['bg'])])
        style.map('Card.TCheckbutton',
                 background=[('active', self.colors['card'])])
        
        # Combobox styles
        style.configure('TCombobox',
                       padding=5,
                       font=('Segoe UI', 9))
        
        # Separator
        style.configure('TSeparator', background=self.colors['border'])
        
    def get_monitors(self):
        """Get all connected monitors with their info"""
        monitors = []
        
        monitors_enum = win32api.EnumDisplayMonitors(None, None)
        
        for hMonitor, hdcMonitor, pyRect in monitors_enum:
            info = win32api.GetMonitorInfo(hMonitor)
            work_area = info['Work']
            monitor_area = info['Monitor']
            is_primary = info['Flags'] & 1
            
            monitors.append({
                'handle': hMonitor,
                'work_area': work_area,
                'monitor_area': monitor_area,
                'is_primary': is_primary,
                'name': f"Monitor {len(monitors) + 1}" + (" (Primary)" if is_primary else ""),
                'resolution': f"{monitor_area[2] - monitor_area[0]}x{monitor_area[3] - monitor_area[1]}"
            })
        
        monitors.sort(key=lambda m: m['monitor_area'][0])
        
        for i, mon in enumerate(monitors):
            primary_tag = " ★" if mon['is_primary'] else ""
            mon['name'] = f"Display {i + 1}{primary_tag} ({mon['resolution']})"
            
        return monitors

    def get_window_monitor(self, hwnd):
        """Get which monitor a window is on"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            center_x = (rect[0] + rect[2]) // 2
            center_y = (rect[1] + rect[3]) // 2
            
            for mon in self.monitors:
                ma = mon['monitor_area']
                if ma[0] <= center_x < ma[2] and ma[1] <= center_y < ma[3]:
                    return mon['name']
        except:
            pass
        return None

    def get_audio_devices(self):
        """Get all audio output devices with proper names"""
        devices = {}
        
        try:
            import winreg
            
            audio_key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Render"
            
            try:
                audio_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, audio_key_path)
                
                i = 0
                while True:
                    try:
                        device_guid = winreg.EnumKey(audio_key, i)
                        device_key = winreg.OpenKey(audio_key, f"{device_guid}\\Properties")
                        
                        try:
                            name_value = winreg.QueryValueEx(device_key, "{a45c254e-df1c-4efd-8020-67d146a850e0},2")
                            device_name = name_value[0]
                        except:
                            try:
                                name_value = winreg.QueryValueEx(device_key, "{b3f8fa53-0004-438e-9003-51a46e139bfc},6")
                                device_name = name_value[0]
                            except:
                                device_name = f"Audio Device ({device_guid[:8]})"
                        
                        try:
                            state_key = winreg.OpenKey(audio_key, device_guid)
                            state = winreg.QueryValueEx(state_key, "DeviceState")[0]
                            if state == 1:
                                devices[device_name] = device_guid
                            winreg.CloseKey(state_key)
                        except:
                            devices[device_name] = device_guid
                        
                        winreg.CloseKey(device_key)
                        i += 1
                    except WindowsError:
                        break
                        
                winreg.CloseKey(audio_key)
                
            except Exception as e:
                print(f"Registry method failed: {e}")
                
            if not devices or all("Audio Device" in name for name in devices.keys()):
                import subprocess
                try:
                    result = subprocess.run(
                        ['powershell', '-Command', 
                         'Get-AudioDevice -List | Where-Object {$_.Type -eq "Playback"} | Select-Object -Property Name,ID | ConvertTo-Json'],
                        capture_output=True, text=True, timeout=5,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        import json
                        data = json.loads(result.stdout)
                        if isinstance(data, list):
                            devices = {d['Name']: d['ID'] for d in data}
                        elif isinstance(data, dict):
                            devices = {data['Name']: data['ID']}
                except:
                    pass
                    
            if not devices or all("Audio Device" in name for name in devices.keys()):
                try:
                    import subprocess
                    result = subprocess.run(
                        ['powershell', '-Command',
                         'Get-WmiObject Win32_SoundDevice | Where-Object {$_.Status -eq "OK"} | Select-Object Name | ConvertTo-Json'],
                        capture_output=True, text=True, timeout=5,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if result.returncode == 0 and result.stdout.strip():
                        import json
                        data = json.loads(result.stdout)
                        if isinstance(data, list):
                            devices = {d['Name']: str(idx) for idx, d in enumerate(data)}
                        elif isinstance(data, dict):
                            devices = {data['Name']: "0"}
                except:
                    pass
                    
        except Exception as e:
            print(f"Error getting audio devices: {e}")
            
        if not devices:
            devices["No audio devices found"] = None
            
        return devices
        
    def setup_ui(self):
        """Setup the user interface"""
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, text="Window Manager", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        # Pin to top checkbox
        pin_cb = ttk.Checkbutton(header_frame, text="📌 Pin", 
                                  variable=self.pin_to_top,
                                  command=self.toggle_pin)
        pin_cb.pack(side=tk.RIGHT)
        
        # Refresh button
        refresh_btn = ttk.Button(header_frame, text="↻", width=3,
                                  command=self.refresh_windows, style='Small.TButton')
        refresh_btn.pack(side=tk.RIGHT, padx=(0, 10))
        
        # Monitor selection section
        monitor_frame = ttk.Frame(main_frame)
        monitor_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(monitor_frame, text="Move to:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        self.monitor_var = tk.StringVar()
        if self.monitors:
            self.monitor_var.set(self.monitors[0]['name'])
            
        monitor_combo = ttk.Combobox(monitor_frame, textvariable=self.monitor_var,
                                      values=[m['name'] for m in self.monitors],
                                      width=25, state='readonly')
        monitor_combo.pack(side=tk.LEFT, padx=(8, 8))
        
        move_btn = ttk.Button(monitor_frame, text="Move", 
                               command=self.move_to_monitor, style='Accent.TButton')
        move_btn.pack(side=tk.LEFT)
        
        # Split buttons
        split_frame = ttk.Frame(main_frame)
        split_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(split_frame, text="◧ Split H", width=10,
                   command=self.split_vertical).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(split_frame, text="⬒ Split V", width=10,
                   command=self.split_horizontal).pack(side=tk.LEFT, padx=(0, 5))
        
        self.selection_label = ttk.Label(split_frame, text="0 selected", style='Muted.TLabel')
        self.selection_label.pack(side=tk.RIGHT)
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Selection controls - minimal row
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(select_frame, text="Select:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        # Minimal text buttons
        btn_style_mini = {'font': ('Segoe UI', 8), 'padding': (4, 2)}
        
        ttk.Button(select_frame, text="All", width=4, style='Small.TButton',
                   command=self.select_all).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Button(select_frame, text="None", width=5, style='Small.TButton',
                   command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="Monitor", width=7, style='Small.TButton',
                   command=self.select_monitor).pack(side=tk.LEFT, padx=2)
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Audio section
        audio_frame = ttk.Frame(main_frame)
        audio_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(audio_frame, text="Audio:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        self.audio_devices = self.get_audio_devices()
        self.audio_var = tk.StringVar()
        
        if self.audio_devices:
            self.audio_var.set(list(self.audio_devices.keys())[0])
            
        self.audio_combo = ttk.Combobox(audio_frame, textvariable=self.audio_var,
                                         values=list(self.audio_devices.keys()),
                                         width=28, state='readonly')
        self.audio_combo.pack(side=tk.LEFT, padx=(8, 8))
        
        ttk.Button(audio_frame, text="Apply", style='Small.TButton',
                   command=self.apply_audio_device).pack(side=tk.LEFT)
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Windows list
        list_container = ttk.Frame(main_frame)
        list_container.pack(fill=tk.BOTH, expand=True)
        
        self.canvas = tk.Canvas(list_container, bg=self.colors['bg'], 
                                 highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", 
                                   command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind("<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, 
                                                        anchor="nw", width=self.canvas.winfo_width())
        
        self.canvas.configure(yscrollcommand=scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.canvas.bind('<Configure>', self.on_canvas_configure)
        
        self.canvas.bind_all("<MouseWheel>", 
                             lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                               style='Muted.TLabel')
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
    def on_canvas_configure(self, event):
        """Handle canvas resize"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        
    def toggle_pin(self):
        """Toggle always on top"""
        self.root.attributes('-topmost', self.pin_to_top.get())
        status = "pinned" if self.pin_to_top.get() else "unpinned"
        self.status_var.set(f"Window {status}")
    
    def select_all(self):
        """Select all windows"""
        for hwnd, var in self.window_checkboxes.items():
            var.set(True)
            self.selected_windows[hwnd] = True
            self.update_card_style(hwnd, True)
        self.update_selection_label()
        self.status_var.set("Selected all windows")
    
    def deselect_all(self):
        """Deselect all windows"""
        for hwnd, var in self.window_checkboxes.items():
            var.set(False)
            self.update_card_style(hwnd, False)
        self.selected_windows.clear()
        self.update_selection_label()
        self.status_var.set("Cleared selection")
    
    def select_monitor(self):
        """Select all windows on the currently selected monitor"""
        target_monitor = self.monitor_var.get()
        count = 0
        
        for hwnd, var in self.window_checkboxes.items():
            window_monitor = self.get_window_monitor(hwnd)
            if window_monitor == target_monitor:
                var.set(True)
                self.selected_windows[hwnd] = True
                self.update_card_style(hwnd, True)
                count += 1
            else:
                var.set(False)
                if hwnd in self.selected_windows:
                    del self.selected_windows[hwnd]
                self.update_card_style(hwnd, False)
        
        self.update_selection_label()
        self.status_var.set(f"Selected {count} windows on {target_monitor.split(' (')[0]}")
        
    def is_real_window(self, hwnd):
        """Check if a window is a real, visible application window"""
        if not win32gui.IsWindowVisible(hwnd):
            return False
            
        if not win32gui.GetWindowText(hwnd):
            return False
            
        try:
            cloaked = ctypes.c_int(0)
            ctypes.windll.dwmapi.DwmGetWindowAttribute(
                hwnd, DWMWA_CLOAKED, ctypes.byref(cloaked), ctypes.sizeof(cloaked))
            if cloaked.value:
                return False
        except:
            pass
            
        ex_style = win32gui.GetWindowLong(hwnd, GWL_EXSTYLE)
        if ex_style & WS_EX_TOOLWINDOW:
            return False
            
        try:
            rect = win32gui.GetWindowRect(hwnd)
            if rect[2] - rect[0] <= 0 or rect[3] - rect[1] <= 0:
                return False
        except:
            return False
            
        return True
    
    def get_process_name(self, hwnd):
        """Get the process name for a window"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            process = psutil.Process(pid)
            return process.name().replace('.exe', '')
        except:
            return "Unknown"
    
    def refresh_windows(self):
        """Refresh the list of windows"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.window_checkboxes.clear()
        self.window_cards.clear()
        self.windows_list.clear()
        
        old_selection = list(self.selected_windows.keys())
        self.selected_windows.clear()
        
        self.monitors = self.get_monitors()
        
        def enum_callback(hwnd, windows):
            if self.is_real_window(hwnd):
                title = win32gui.GetWindowText(hwnd)
                process = self.get_process_name(hwnd)
                windows.append((hwnd, title, process))
            return True
            
        windows = []
        win32gui.EnumWindows(enum_callback, windows)
        self.windows_list = windows
        
        our_hwnd = self.root.winfo_id()
        windows = [(h, t, p) for h, t, p in windows if h != our_hwnd]
        
        for i, (hwnd, title, process) in enumerate(windows):
            self.create_window_card(hwnd, title, process, hwnd in old_selection)
                
        self.update_selection_label()
        self.status_var.set(f"{len(windows)} windows")
    
    def update_card_style(self, hwnd, selected):
        """Update card visual style based on selection state"""
        if hwnd not in self.window_cards:
            return
            
        card_data = self.window_cards[hwnd]
        card = card_data['card']
        inner = card_data['inner']
        left = card_data['left']
        info = card_data['info']
        actions = card_data['actions']
        cb = card_data['checkbox']
        indicator = card_data['indicator']
        
        if selected:
            bg_color = self.colors['card_selected']
            indicator.configure(bg=self.colors['checkbox_checked'])
        else:
            bg_color = self.colors['card']
            indicator.configure(bg=self.colors['checkbox_unchecked'])
        
        card.configure(bg=bg_color)
        inner.configure(bg=bg_color)
        left.configure(bg=bg_color)
        info.configure(bg=bg_color)
        actions.configure(bg=bg_color)
        cb.configure(bg=bg_color, activebackground=bg_color)
        
        for widget in info.winfo_children():
            widget.configure(bg=bg_color)
        
        card_data['base_bg'] = bg_color
        
    def create_window_card(self, hwnd, title, process, was_selected=False):
        """Create a sleek window card"""
        base_bg = self.colors['card_selected'] if was_selected else self.colors['card']
        
        # Card frame
        card = tk.Frame(self.scrollable_frame, bg=base_bg, highlightthickness=0)
        card.pack(fill=tk.X, pady=2, padx=2)
        
        # Inner padding frame
        inner = tk.Frame(card, bg=base_bg)
        inner.pack(fill=tk.X, padx=10, pady=8)
        
        # Left side: indicator + checkbox + info
        left = tk.Frame(inner, bg=base_bg)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Selection indicator (colored bar)
        indicator_color = self.colors['checkbox_checked'] if was_selected else self.colors['checkbox_unchecked']
        indicator = tk.Frame(left, bg=indicator_color, width=4, height=32)
        indicator.pack(side=tk.LEFT, padx=(0, 8))
        indicator.pack_propagate(False)
        
        # Checkbox
        var = tk.BooleanVar(value=was_selected)
        if was_selected:
            self.selected_windows[hwnd] = True
            
        cb = tk.Checkbutton(left, variable=var, bg=base_bg,
                            activebackground=base_bg,
                            selectcolor=self.colors['bg'],
                            command=lambda h=hwnd, v=var: self.on_checkbox_changed(h, v))
        cb.pack(side=tk.LEFT)
        
        self.window_checkboxes[hwnd] = var
        
        # Info container
        info = tk.Frame(left, bg=base_bg)
        info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        # Process name (bold)
        proc_label = tk.Label(info, text=process, bg=base_bg,
                              fg=self.colors['fg'], font=('Segoe UI', 9, 'bold'),
                              anchor='w')
        proc_label.pack(fill=tk.X)
        
        # Window title (muted, truncated)
        display_title = title[:35] + "…" if len(title) > 35 else title
        title_label = tk.Label(info, text=display_title, bg=base_bg,
                               fg=self.colors['muted'], font=('Segoe UI', 8),
                               anchor='w')
        title_label.pack(fill=tk.X)
        
        # Right side: action buttons
        actions = tk.Frame(inner, bg=base_bg)
        actions.pack(side=tk.RIGHT)
        
        # Minimal icon buttons
        btn_style = {'bg': base_bg, 'fg': self.colors['muted'],
                     'font': ('Segoe UI', 10), 'bd': 0, 'padx': 6, 'pady': 2,
                     'activebackground': self.colors['card_hover'],
                     'activeforeground': self.colors['fg'], 'cursor': 'hand2'}
        
        focus_btn = tk.Button(actions, text="◉", command=lambda h=hwnd: self.focus_window(h), **btn_style)
        focus_btn.pack(side=tk.LEFT)
        min_btn = tk.Button(actions, text="−", command=lambda h=hwnd: self.minimize_window(h), **btn_style)
        min_btn.pack(side=tk.LEFT)
        max_btn = tk.Button(actions, text="□", command=lambda h=hwnd: self.maximize_window(h), **btn_style)
        max_btn.pack(side=tk.LEFT)
        
        # Store card references
        self.window_cards[hwnd] = {
            'card': card,
            'inner': inner,
            'left': left,
            'info': info,
            'actions': actions,
            'checkbox': cb,
            'indicator': indicator,
            'buttons': [focus_btn, min_btn, max_btn],
            'base_bg': base_bg
        }
        
        # Hover effects
        def on_enter(e):
            current_base = self.window_cards[hwnd]['base_bg']
            hover_bg = self.colors['card_hover'] if current_base == self.colors['card'] else '#243454'
            
            card.configure(bg=hover_bg)
            inner.configure(bg=hover_bg)
            left.configure(bg=hover_bg)
            info.configure(bg=hover_bg)
            actions.configure(bg=hover_bg)
            for widget in info.winfo_children():
                widget.configure(bg=hover_bg)
            cb.configure(bg=hover_bg, activebackground=hover_bg)
            for btn in self.window_cards[hwnd]['buttons']:
                btn.configure(bg=hover_bg)
                
        def on_leave(e):
            current_base = self.window_cards[hwnd]['base_bg']
            
            card.configure(bg=current_base)
            inner.configure(bg=current_base)
            left.configure(bg=current_base)
            info.configure(bg=current_base)
            actions.configure(bg=current_base)
            for widget in info.winfo_children():
                widget.configure(bg=current_base)
            cb.configure(bg=current_base, activebackground=current_base)
            for btn in self.window_cards[hwnd]['buttons']:
                btn.configure(bg=current_base)
                
        card.bind('<Enter>', on_enter)
        card.bind('<Leave>', on_leave)
        
    def on_checkbox_changed(self, hwnd, var):
        """Handle checkbox state change"""
        if var.get():
            self.selected_windows[hwnd] = True
        else:
            if hwnd in self.selected_windows:
                del self.selected_windows[hwnd]
        self.update_card_style(hwnd, var.get())
        self.update_selection_label()
        
    def update_selection_label(self):
        """Update selection count"""
        count = len(self.selected_windows)
        self.selection_label.config(text=f"{count} selected")
    
    def get_selected_windows(self):
        """Get selected window handles in order"""
        return list(self.selected_windows.keys())
    
    def focus_window(self, hwnd):
        """Focus a window"""
        was_pinned = self.temporarily_pin()
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            self.status_var.set(f"Error: {e}")
        self.restore_pin_state(was_pinned)
            
    def minimize_window(self, hwnd):
        """Minimize a window"""
        was_pinned = self.temporarily_pin()
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        except Exception as e:
            self.status_var.set(f"Error: {e}")
        self.restore_pin_state(was_pinned)
            
    def maximize_window(self, hwnd):
        """Maximize a window"""
        was_pinned = self.temporarily_pin()
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        except Exception as e:
            self.status_var.set(f"Error: {e}")
        self.restore_pin_state(was_pinned)
            
    def move_to_monitor(self):
        """Move selected windows to chosen monitor"""
        was_pinned = self.temporarily_pin()
        selected = self.get_selected_windows()
        
        if not selected:
            self.status_var.set("No windows selected")
            self.restore_pin_state(was_pinned)
            return
            
        monitor_name = self.monitor_var.get()
        target_monitor = None
        
        for mon in self.monitors:
            if mon['name'] == monitor_name:
                target_monitor = mon
                break
                
        if not target_monitor:
            self.status_var.set("Monitor not found")
            self.restore_pin_state(was_pinned)
            return
            
        work_area = target_monitor['work_area']
        mon_x, mon_y = work_area[0], work_area[1]
        mon_width = work_area[2] - work_area[0]
        mon_height = work_area[3] - work_area[1]
        
        moved_count = 0
        
        for hwnd in selected:
            try:
                rect = win32gui.GetWindowRect(hwnd)
                win_width = rect[2] - rect[0]
                win_height = rect[3] - rect[1]
                
                if self.is_window_maximized(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    rect = win32gui.GetWindowRect(hwnd)
                    win_width = rect[2] - rect[0]
                    win_height = rect[3] - rect[1]
                
                if self.is_window_minimized(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    
                win_width = min(win_width, mon_width)
                win_height = min(win_height, mon_height)
                
                new_x = mon_x + (mon_width - win_width) // 2
                new_y = mon_y + (mon_height - win_height) // 2
                
                win32gui.SetWindowPos(hwnd, win32con.HWND_TOP,
                                       new_x, new_y, win_width, win_height,
                                       win32con.SWP_SHOWWINDOW)
                moved_count += 1
                
            except Exception as e:
                print(f"Error moving window {hwnd}: {e}")
                
        self.status_var.set(f"Moved {moved_count} window(s)")
        self.restore_pin_state(was_pinned)
    
    def get_target_monitor(self):
        """Get the work area for the selected monitor"""
        monitor_name = self.monitor_var.get()
        
        for mon in self.monitors:
            if mon['name'] == monitor_name:
                return mon['work_area']
                
        return self.monitors[0]['work_area'] if self.monitors else (0, 0, 1920, 1080)
    
    def split_vertical(self):
        """Split first two windows side by side"""
        was_pinned = self.temporarily_pin()
        selected = self.get_selected_windows()
        
        if len(selected) < 2:
            self.status_var.set("Select at least 2 windows")
            self.restore_pin_state(was_pinned)
            return
            
        hwnd1, hwnd2 = selected[0], selected[1]
        work_area = self.get_target_monitor()
        
        work_x, work_y = work_area[0], work_area[1]
        work_width = work_area[2] - work_area[0]
        work_height = work_area[3] - work_area[1]
        half_width = work_width // 2
        
        try:
            win32gui.ShowWindow(hwnd1, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd2, win32con.SW_RESTORE)
            
            win32gui.SetWindowPos(hwnd1, win32con.HWND_TOP,
                                   work_x, work_y, half_width, work_height,
                                   win32con.SWP_SHOWWINDOW)
            win32gui.SetWindowPos(hwnd2, win32con.HWND_TOP,
                                   work_x + half_width, work_y, half_width, work_height,
                                   win32con.SWP_SHOWWINDOW)
                                   
            self.status_var.set("Split side by side")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
        self.restore_pin_state(was_pinned)
    
    def split_horizontal(self):
        """Split first two windows top/bottom"""
        was_pinned = self.temporarily_pin()
        selected = self.get_selected_windows()
        
        if len(selected) < 2:
            self.status_var.set("Select at least 2 windows")
            self.restore_pin_state(was_pinned)
            return
            
        hwnd1, hwnd2 = selected[0], selected[1]
        work_area = self.get_target_monitor()
        
        work_x, work_y = work_area[0], work_area[1]
        work_width = work_area[2] - work_area[0]
        work_height = work_area[3] - work_area[1]
        half_height = work_height // 2
        
        try:
            win32gui.ShowWindow(hwnd1, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd2, win32con.SW_RESTORE)
            
            win32gui.SetWindowPos(hwnd1, win32con.HWND_TOP,
                                   work_x, work_y, work_width, half_height,
                                   win32con.SWP_SHOWWINDOW)
            win32gui.SetWindowPos(hwnd2, win32con.HWND_TOP,
                                   work_x, work_y + half_height, work_width, half_height,
                                   win32con.SWP_SHOWWINDOW)
                                   
            self.status_var.set("Split top/bottom")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
        self.restore_pin_state(was_pinned)
    
    def apply_audio_device(self):
        """Apply audio device (shows info about limitations)"""
        selected = self.get_selected_windows()
        device_name = self.audio_var.get()
        
        if not selected:
            self.status_var.set("No windows selected")
            return
            
        self.status_var.set("Audio routing requires system API - see Windows Sound settings")
        
        try:
            import subprocess
            subprocess.Popen('ms-settings:apps-volume', shell=True)
        except:
            pass
    
    def refresh_audio_devices(self):
        """Refresh audio devices"""
        self.audio_devices = self.get_audio_devices()
        self.audio_combo['values'] = list(self.audio_devices.keys())
        if self.audio_devices:
            self.audio_var.set(list(self.audio_devices.keys())[0])
            
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    required_modules = ['win32gui', 'win32con', 'win32process', 'win32api', 'psutil']
    missing = []
    
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            missing.append(module)
    
    if missing:
        print("Missing required modules. Install with:")
        print("pip install pywin32 psutil")
        return
        
    app = WindowManager()
    app.run()


if __name__ == "__main__":
    main()