import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes
import win32gui
import win32con
import win32process
import win32api
import psutil
from collections import OrderedDict

# Windows API constants
DWMWA_CLOAKED = 14
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000


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
            'checkbox_unchecked': '#444444',
            'slider_bg': '#333333',
            'slider_fg': '#3d5afe',
            'muted_icon': '#ff6b6b',
            'unmuted_icon': '#4caf50'
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        # Pin to top variable - explicitly set to False initially
        self.pin_to_top = tk.BooleanVar(value=False)
        
        # Explicitly set topmost to False on startup
        self.root.attributes('-topmost', False)
        
        # Track selected windows in order
        self.selected_windows = OrderedDict()
        self.window_checkboxes = {}
        self.window_cards = {}
        self.windows_list = []
        self.window_pids = {}  # Map hwnd to pid for audio control
        
        # Volume slider popup
        self.volume_slider_window = None
        
        # Get monitors
        self.monitors = self.get_monitors()
        
        # Initialize audio
        self.init_audio()
        
        self.setup_styles()
        self.setup_ui()
        self.refresh_windows()
    
    def init_audio(self):
        """Initialize audio control via pycaw"""
        self.audio_available = False
        try:
            from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
            from ctypes import cast, POINTER
            from comtypes import CLSCTX_ALL
            
            self.AudioUtilities = AudioUtilities
            self.ISimpleAudioVolume = ISimpleAudioVolume
            self.cast = cast
            self.POINTER = POINTER
            self.CLSCTX_ALL = CLSCTX_ALL
            
            # Test that we can get the endpoint volume interface
            from pycaw.pycaw import IAudioEndpointVolume
            self.IAudioEndpointVolume = IAudioEndpointVolume
            
            self.audio_available = True
            print("Audio control initialized successfully")
        except ImportError as e:
            print(f"pycaw not available - audio control disabled: {e}")
            print("Install with: pip install pycaw comtypes")
        except Exception as e:
            print(f"Error initializing audio: {e}")
    
    def get_audio_session_for_pid(self, pid):
        """Get audio session for a specific process ID"""
        if not self.audio_available:
            return None
        try:
            sessions = self.AudioUtilities.GetAllSessions()
            for session in sessions:
                if session.Process and session.Process.pid == pid:
                    return session
        except Exception as e:
            pass
        return None
    
    def get_app_volume(self, pid):
        """Get volume level for an app (0.0 to 1.0)"""
        if not self.audio_available:
            return None
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                return volume_interface.GetMasterVolume()
            except Exception as e:
                pass
        return None
    
    def set_app_volume(self, pid, level):
        """Set volume level for an app (0.0 to 1.0)"""
        if not self.audio_available:
            return False
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                volume_interface.SetMasterVolume(float(level), None)
                return True
            except Exception as e:
                pass
        return False
    
    def get_app_mute(self, pid):
        """Get mute state for an app"""
        if not self.audio_available:
            return None
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                return bool(volume_interface.GetMute())
            except Exception as e:
                pass
        return None
    
    def set_app_mute(self, pid, mute):
        """Set mute state for an app"""
        if not self.audio_available:
            return False
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                volume_interface.SetMute(int(mute), None)
                return True
            except Exception as e:
                pass
        return False
    
    def get_system_volume(self):
        """Get system master volume (0.0 to 1.0)"""
        if not self.audio_available:
            return 1.0
        try:
            devices = self.AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                self.IAudioEndpointVolume._iid_, self.CLSCTX_ALL, None)
            volume = self.cast(interface, self.POINTER(self.IAudioEndpointVolume))
            return volume.GetMasterVolumeLevelScalar()
        except Exception as e:
            print(f"Error getting system volume: {e}")
        return 1.0
    
    def set_system_volume(self, level):
        """Set system master volume (0.0 to 1.0)"""
        if not self.audio_available:
            return False
        try:
            devices = self.AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                self.IAudioEndpointVolume._iid_, self.CLSCTX_ALL, None)
            volume = self.cast(interface, self.POINTER(self.IAudioEndpointVolume))
            volume.SetMasterVolumeLevelScalar(float(level), None)
            return True
        except Exception as e:
            print(f"Error setting system volume: {e}")
        return False
    
    def get_system_mute(self):
        """Get system mute state"""
        if not self.audio_available:
            return False
        try:
            devices = self.AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                self.IAudioEndpointVolume._iid_, self.CLSCTX_ALL, None)
            volume = self.cast(interface, self.POINTER(self.IAudioEndpointVolume))
            return bool(volume.GetMute())
        except Exception as e:
            pass
        return False
    
    def set_system_mute(self, mute):
        """Set system mute state"""
        if not self.audio_available:
            return False
        try:
            devices = self.AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                self.IAudioEndpointVolume._iid_, self.CLSCTX_ALL, None)
            volume = self.cast(interface, self.POINTER(self.IAudioEndpointVolume))
            volume.SetMute(int(mute), None)
            return True
        except Exception as e:
            pass
        return False
    
    def toggle_pin(self):
        """Toggle always on top"""
        is_pinned = self.pin_to_top.get()
        self.root.attributes('-topmost', is_pinned)
        status = "pinned" if is_pinned else "unpinned"
        self.status_var.set(f"Window {status}")
    
    def ensure_topmost_during_action(self):
        """Temporarily ensure window is on top during an action"""
        was_pinned = self.pin_to_top.get()
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.update_idletasks()
        
        if not was_pinned:
            if hasattr(self, '_restore_job') and self._restore_job:
                self.root.after_cancel(self._restore_job)
            self._restore_job = self.root.after(200, lambda: self.root.attributes('-topmost', False))
    
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
        
        style.configure('TFrame', background=self.colors['bg'])
        style.configure('Card.TFrame', background=self.colors['card'])
        
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
        
        style.configure('TCombobox',
                       padding=5,
                       font=('Segoe UI', 9))
        
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
        
    def setup_ui(self):
        """Setup the user interface"""
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, text="Window Manager", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        pin_cb = ttk.Checkbutton(header_frame, text="📌 Pin", 
                                  variable=self.pin_to_top,
                                  command=self.toggle_pin)
        pin_cb.pack(side=tk.RIGHT)
        
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
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Selection controls
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(select_frame, text="Select:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        ttk.Button(select_frame, text="All", width=4, style='Small.TButton',
                   command=self.select_all).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Button(select_frame, text="None", width=5, style='Small.TButton',
                   command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="Monitor", width=7, style='Small.TButton',
                   command=self.select_monitor).pack(side=tk.LEFT, padx=2)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Audio bulk controls
        audio_frame = ttk.Frame(main_frame)
        audio_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(audio_frame, text="Audio:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        self.bulk_mute_btn = tk.Button(audio_frame, text="🔇 Mute", 
                                        bg=self.colors['card'], fg=self.colors['fg'],
                                        font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                        activebackground=self.colors['card_hover'],
                                        command=self.bulk_mute)
        self.bulk_mute_btn.pack(side=tk.LEFT, padx=(8, 4))
        
        self.bulk_unmute_btn = tk.Button(audio_frame, text="🔊 Unmute",
                                          bg=self.colors['card'], fg=self.colors['fg'],
                                          font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                          activebackground=self.colors['card_hover'],
                                          command=self.bulk_unmute)
        self.bulk_unmute_btn.pack(side=tk.LEFT, padx=(0, 4))
        
        # Volume button with right-click-hold for slider
        self.bulk_volume_btn = tk.Button(audio_frame, text="🎚️ Vol",
                                          bg=self.colors['card'], fg=self.colors['fg'],
                                          font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                          activebackground=self.colors['card_hover'])
        self.bulk_volume_btn.pack(side=tk.LEFT)
        self.bulk_volume_btn.bind('<ButtonPress-3>', self.on_bulk_volume_press)
        self.bulk_volume_btn.bind('<ButtonRelease-3>', self.on_volume_release)
        
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
    
    def get_process_pid(self, hwnd):
        """Get the process ID for a window"""
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            return pid
        except:
            return None
    
    def refresh_windows(self):
        """Refresh the list of windows"""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.window_checkboxes.clear()
        self.window_cards.clear()
        self.windows_list.clear()
        self.window_pids.clear()
        
        old_selection = list(self.selected_windows.keys())
        self.selected_windows.clear()
        
        self.monitors = self.get_monitors()
        
        def enum_callback(hwnd, windows):
            if self.is_real_window(hwnd):
                title = win32gui.GetWindowText(hwnd)
                process = self.get_process_name(hwnd)
                pid = self.get_process_pid(hwnd)
                windows.append((hwnd, title, process, pid))
            return True
            
        windows = []
        win32gui.EnumWindows(enum_callback, windows)
        self.windows_list = windows
        
        our_hwnd = self.root.winfo_id()
        windows = [(h, t, p, pid) for h, t, p, pid in windows if h != our_hwnd]
        
        for i, (hwnd, title, process, pid) in enumerate(windows):
            self.window_pids[hwnd] = pid
            self.create_window_card(hwnd, title, process, pid, hwnd in old_selection)
                
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
        
    def create_window_card(self, hwnd, title, process, pid, was_selected=False):
        """Create a sleek window card"""
        base_bg = self.colors['card_selected'] if was_selected else self.colors['card']
        
        card = tk.Frame(self.scrollable_frame, bg=base_bg, highlightthickness=0)
        card.pack(fill=tk.X, pady=2, padx=2)
        
        inner = tk.Frame(card, bg=base_bg)
        inner.pack(fill=tk.X, padx=10, pady=8)
        
        left = tk.Frame(inner, bg=base_bg)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        indicator_color = self.colors['checkbox_checked'] if was_selected else self.colors['checkbox_unchecked']
        indicator = tk.Frame(left, bg=indicator_color, width=4, height=32)
        indicator.pack(side=tk.LEFT, padx=(0, 8))
        indicator.pack_propagate(False)
        
        var = tk.BooleanVar(value=was_selected)
        if was_selected:
            self.selected_windows[hwnd] = True
            
        cb = tk.Checkbutton(left, variable=var, bg=base_bg,
                            activebackground=base_bg,
                            selectcolor=self.colors['bg'],
                            command=lambda h=hwnd, v=var: self.on_checkbox_changed(h, v))
        cb.pack(side=tk.LEFT)
        
        self.window_checkboxes[hwnd] = var
        
        info = tk.Frame(left, bg=base_bg)
        info.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        proc_label = tk.Label(info, text=process, bg=base_bg,
                              fg=self.colors['fg'], font=('Segoe UI', 9, 'bold'),
                              anchor='w')
        proc_label.pack(fill=tk.X)
        
        display_title = title[:35] + "…" if len(title) > 35 else title
        title_label = tk.Label(info, text=display_title, bg=base_bg,
                               fg=self.colors['muted'], font=('Segoe UI', 8),
                               anchor='w')
        title_label.pack(fill=tk.X)
        
        actions = tk.Frame(inner, bg=base_bg)
        actions.pack(side=tk.RIGHT)
        
        btn_style = {'bg': base_bg, 'fg': self.colors['muted'],
                     'font': ('Segoe UI', 10), 'bd': 0, 'padx': 6, 'pady': 2,
                     'activebackground': self.colors['card_hover'],
                     'activeforeground': self.colors['fg'], 'cursor': 'hand2'}
        
        # Focus button
        focus_btn = tk.Button(actions, text="◉", command=lambda h=hwnd: self.focus_window(h), **btn_style)
        focus_btn.pack(side=tk.LEFT)
        
        # Min/Max toggle button
        minmax_btn = tk.Button(actions, text="□", command=lambda h=hwnd: self.toggle_minmax(h), **btn_style)
        minmax_btn.pack(side=tk.LEFT)
        
        # Audio mute toggle button with right-click-hold for slider
        is_muted = self.get_app_mute(pid) if pid else False
        audio_icon = "🔇" if is_muted else "🔊"
        audio_btn = tk.Button(actions, text=audio_icon, **btn_style)
        audio_btn.configure(command=lambda h=hwnd, p=pid, b=audio_btn: self.toggle_app_mute(h, p, b))
        audio_btn.bind('<ButtonPress-3>', lambda e, h=hwnd, p=pid, b=audio_btn: self.on_app_volume_press(e, h, p, b))
        audio_btn.bind('<ButtonRelease-3>', self.on_volume_release)
        audio_btn.pack(side=tk.LEFT)
        
        self.window_cards[hwnd] = {
            'card': card,
            'inner': inner,
            'left': left,
            'info': info,
            'actions': actions,
            'checkbox': cb,
            'indicator': indicator,
            'buttons': [focus_btn, minmax_btn, audio_btn],
            'audio_btn': audio_btn,
            'base_bg': base_bg,
            'pid': pid
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
    
    def toggle_minmax(self, hwnd):
        """Toggle between minimized and maximized states"""
        self.ensure_topmost_during_action()
        try:
            if self.is_window_minimized(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                self.status_var.set("Window maximized")
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                self.status_var.set("Window minimized")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
    
    def toggle_app_mute(self, hwnd, pid, btn):
        """Toggle mute state for an app"""
        if not pid or not self.audio_available:
            self.status_var.set("No audio session for this window")
            return
            
        current_mute = self.get_app_mute(pid)
        if current_mute is None:
            self.status_var.set("No audio session for this window")
            return
            
        new_mute = not current_mute
        if self.set_app_mute(pid, new_mute):
            btn.configure(text="🔇" if new_mute else "🔊")
            self.status_var.set("Muted" if new_mute else "Unmuted")
        else:
            self.status_var.set("Failed to toggle mute")
    
    def on_app_volume_press(self, event, hwnd, pid, btn):
        """Handle right-click press on app audio button - show slider"""
        if not pid or not self.audio_available:
            self.status_var.set("No audio session for this window")
            return
            
        current_volume = self.get_app_volume(pid)
        if current_volume is None:
            self.status_var.set("No audio session for this window")
            return
        
        def on_change(v):
            self.set_app_volume(pid, v)
        
        def on_close():
            self.update_audio_btn(hwnd, pid, btn)
        
        self.show_volume_slider(event, current_volume, on_change, on_close, "App Volume")
    
    def on_bulk_volume_press(self, event):
        """Handle right-click press on bulk volume button"""
        selected = self.get_selected_windows()
        
        if not selected:
            # No selection - control system volume
            current_volume = self.get_system_volume()
            
            def on_change(v):
                self.set_system_volume(v)
            
            self.show_volume_slider(event, current_volume, on_change, None, "System")
        else:
            # Control selected windows - start at 100%
            def on_change(v):
                self.set_selected_volumes(v)
            
            def on_close():
                self.update_all_audio_btns()
            
            self.show_volume_slider(event, 1.0, on_change, on_close, f"{len(selected)} Apps")
    
    def on_volume_release(self, event):
        """Handle right-click release - close slider"""
        self.close_volume_slider()
    
    def show_volume_slider(self, event, initial_volume, on_change, on_close=None, title="Volume"):
        """Create a floating volume slider window that stays open while right-click is held"""
        # Close any existing slider
        self.close_volume_slider()
        
        # Create popup window
        slider_win = tk.Toplevel(self.root)
        slider_win.overrideredirect(True)
        slider_win.attributes('-topmost', True)
        
        # Calculate position near the button
        x = event.x_root - 25
        y = event.y_root - 180  # Position above the button
        
        # Slider height is approximately twice a card height
        slider_height = 160
        slider_width = 50
        
        slider_win.geometry(f"{slider_width}x{slider_height}+{x}+{y}")
        slider_win.configure(bg=self.colors['card'])
        
        # Store reference and callbacks
        self.volume_slider_window = slider_win
        self.volume_slider_on_close = on_close
        self.volume_slider_on_change = on_change
        
        # Title label
        title_lbl = tk.Label(slider_win, text=title, bg=self.colors['card'],
                             fg=self.colors['muted'], font=('Segoe UI', 7))
        title_lbl.pack(pady=(5, 0))
        
        # Volume percentage label
        self.vol_var = tk.StringVar(value=f"{int(initial_volume * 100)}%")
        vol_label = tk.Label(slider_win, textvariable=self.vol_var, bg=self.colors['card'],
                             fg=self.colors['fg'], font=('Segoe UI', 9, 'bold'))
        vol_label.pack(pady=(2, 5))
        
        # Create canvas for custom slider
        canvas_height = slider_height - 60
        self.slider_canvas = tk.Canvas(slider_win, width=30, height=canvas_height,
                                        bg=self.colors['slider_bg'], highlightthickness=0)
        self.slider_canvas.pack(pady=(0, 5))
        
        # Slider parameters
        self.track_x = 15
        self.track_top = 5
        self.track_bottom = canvas_height - 5
        self.track_height = self.track_bottom - self.track_top
        
        # Track background
        self.slider_canvas.create_rectangle(self.track_x - 3, self.track_top, 
                                             self.track_x + 3, self.track_bottom,
                                             fill=self.colors['slider_bg'], 
                                             outline=self.colors['border'])
        
        # Calculate initial handle position based on actual volume
        handle_y = self.track_bottom - (initial_volume * self.track_height)
        
        # Filled portion
        self.fill_rect = self.slider_canvas.create_rectangle(
            self.track_x - 3, handle_y, self.track_x + 3, self.track_bottom,
            fill=self.colors['slider_fg'], outline='')
        
        # Handle
        self.handle = self.slider_canvas.create_oval(
            self.track_x - 8, handle_y - 8, self.track_x + 8, handle_y + 8,
            fill=self.colors['accent'], outline=self.colors['fg'])
        
        # Bind motion events to root for tracking mouse while right-click held
        self.root.bind('<Motion>', self.on_slider_motion)
        slider_win.bind('<Motion>', self.on_slider_motion)
    
    def on_slider_motion(self, event):
        """Handle mouse motion while slider is open"""
        if not self.volume_slider_window or not self.slider_canvas:
            return
        
        try:
            # Get mouse position relative to canvas
            canvas_x = self.slider_canvas.winfo_rootx()
            canvas_y = self.slider_canvas.winfo_rooty()
            
            # Calculate y position relative to canvas
            y_pos = event.y_root - canvas_y
            
            # Clamp y position to track bounds
            y_pos = max(self.track_top, min(self.track_bottom, y_pos))
            
            # Calculate volume (inverted - top is high, bottom is low)
            volume = (self.track_bottom - y_pos) / self.track_height
            volume = max(0.0, min(1.0, volume))
            
            # Update visual
            self.slider_canvas.coords(self.handle, 
                                       self.track_x - 8, y_pos - 8, 
                                       self.track_x + 8, y_pos + 8)
            self.slider_canvas.coords(self.fill_rect, 
                                       self.track_x - 3, y_pos, 
                                       self.track_x + 3, self.track_bottom)
            
            # Update label
            self.vol_var.set(f"{int(volume * 100)}%")
            
            # Apply volume change
            if self.volume_slider_on_change:
                self.volume_slider_on_change(volume)
                
        except Exception as e:
            pass
    
    def close_volume_slider(self):
        """Close the volume slider"""
        # Unbind motion event
        try:
            self.root.unbind('<Motion>')
        except:
            pass
        
        # Call on_close callback if set
        if hasattr(self, 'volume_slider_on_close') and self.volume_slider_on_close:
            try:
                self.volume_slider_on_close()
            except:
                pass
            self.volume_slider_on_close = None
        
        # Destroy slider window
        if self.volume_slider_window:
            try:
                self.volume_slider_window.destroy()
            except:
                pass
            self.volume_slider_window = None
        
        # Clear references
        self.slider_canvas = None
        self.volume_slider_on_change = None
    
    def update_audio_btn(self, hwnd, pid, btn):
        """Update audio button icon based on mute state"""
        if hwnd in self.window_cards:
            is_muted = self.get_app_mute(pid)
            if is_muted is not None:
                btn.configure(text="🔇" if is_muted else "🔊")
    
    def update_all_audio_btns(self):
        """Update all audio button icons"""
        for hwnd, card_data in self.window_cards.items():
            pid = card_data.get('pid')
            if pid:
                is_muted = self.get_app_mute(pid)
                if is_muted is not None:
                    card_data['audio_btn'].configure(text="🔇" if is_muted else "🔊")
    
    def set_selected_volumes(self, volume):
        """Set volume for all selected windows"""
        selected = self.get_selected_windows()
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid:
                self.set_app_volume(pid, volume)
    
    def bulk_mute(self):
        """Mute selected windows or system"""
        selected = self.get_selected_windows()
        
        if not selected:
            if self.set_system_mute(True):
                self.status_var.set("System muted")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid and self.set_app_mute(pid, True):
                count += 1
                if hwnd in self.window_cards:
                    self.window_cards[hwnd]['audio_btn'].configure(text="🔇")
        
        self.status_var.set(f"Muted {count} window(s)")
    
    def bulk_unmute(self):
        """Unmute selected windows or system"""
        selected = self.get_selected_windows()
        
        if not selected:
            if self.set_system_mute(False):
                self.status_var.set("System unmuted")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid and self.set_app_mute(pid, False):
                count += 1
                if hwnd in self.window_cards:
                    self.window_cards[hwnd]['audio_btn'].configure(text="🔊")
        
        self.status_var.set(f"Unmuted {count} window(s)")
        
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
        self.ensure_topmost_during_action()
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            
    def move_to_monitor(self):
        """Move selected windows to chosen monitor"""
        self.ensure_topmost_during_action()
        selected = self.get_selected_windows()
        
        if not selected:
            self.status_var.set("No windows selected")
            return
            
        monitor_name = self.monitor_var.get()
        target_monitor = None
        
        for mon in self.monitors:
            if mon['name'] == monitor_name:
                target_monitor = mon
                break
                
        if not target_monitor:
            self.status_var.set("Monitor not found")
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
    
    def get_target_monitor(self):
        """Get the work area for the selected monitor"""
        monitor_name = self.monitor_var.get()
        
        for mon in self.monitors:
            if mon['name'] == monitor_name:
                return mon['work_area']
                
        return self.monitors[0]['work_area'] if self.monitors else (0, 0, 1920, 1080)
    
    def split_vertical(self):
        """Split first two windows side by side"""
        self.ensure_topmost_during_action()
        selected = self.get_selected_windows()
        
        if len(selected) < 2:
            self.status_var.set("Select at least 2 windows")
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
    
    def split_horizontal(self):
        """Split first two windows top/bottom"""
        self.ensure_topmost_during_action()
        selected = self.get_selected_windows()
        
        if len(selected) < 2:
            self.status_var.set("Select at least 2 windows")
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
    
    try:
        import pycaw
        print("pycaw found - audio control enabled")
    except ImportError:
        print("Note: pycaw not installed. Audio control will be limited.")
        print("Install with: pip install pycaw comtypes")
        
    app = WindowManager()
    app.run()


if __name__ == "__main__":
    main()