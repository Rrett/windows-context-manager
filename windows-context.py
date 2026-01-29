import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes, Structure, POINTER, cast, HRESULT
import win32gui
import win32con
import win32process
import win32api
from comtypes import GUID, IUnknown, COMMETHOD, CoCreateInstance, CLSCTX_ALL
import psutil
from collections import OrderedDict

# Windows API constants
DWMWA_CLOAKED = 14
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000

# Audio API interfaces
class IAudioEndpointVolume(IUnknown):
    _iid_ = GUID('{5CDF2C82-841E-4546-9722-0CF74078229A}')
    _methods_ = [
        COMMETHOD([], HRESULT, 'RegisterControlChangeNotify', (['in'], POINTER(IUnknown), 'pNotify')),
        COMMETHOD([], HRESULT, 'UnregisterControlChangeNotify', (['in'], POINTER(IUnknown), 'pNotify')),
        COMMETHOD([], HRESULT, 'GetChannelCount', (['out', 'retval'], POINTER(ctypes.c_uint), 'pnChannelCount')),
        COMMETHOD([], HRESULT, 'SetMasterVolumeLevel', (['in'], ctypes.c_float, 'fLevelDB'), (['in'], POINTER(GUID), 'pguidEventContext')),
        COMMETHOD([], HRESULT, 'SetMasterVolumeLevelScalar', (['in'], ctypes.c_float, 'fLevel'), (['in'], POINTER(GUID), 'pguidEventContext')),
        COMMETHOD([], HRESULT, 'GetMasterVolumeLevel', (['out', 'retval'], POINTER(ctypes.c_float), 'pfLevelDB')),
        COMMETHOD([], HRESULT, 'GetMasterVolumeLevelScalar', (['out', 'retval'], POINTER(ctypes.c_float), 'pfLevel')),
        COMMETHOD([], HRESULT, 'SetChannelVolumeLevel', (['in'], ctypes.c_uint, 'nChannel'), (['in'], ctypes.c_float, 'fLevelDB'), (['in'], POINTER(GUID), 'pguidEventContext')),
        COMMETHOD([], HRESULT, 'SetChannelVolumeLevelScalar', (['in'], ctypes.c_uint, 'nChannel'), (['in'], ctypes.c_float, 'fLevel'), (['in'], POINTER(GUID), 'pguidEventContext')),
        COMMETHOD([], HRESULT, 'GetChannelVolumeLevel', (['in'], ctypes.c_uint, 'nChannel'), (['out', 'retval'], POINTER(ctypes.c_float), 'pfLevelDB')),
        COMMETHOD([], HRESULT, 'GetChannelVolumeLevelScalar', (['in'], ctypes.c_uint, 'nChannel'), (['out', 'retval'], POINTER(ctypes.c_float), 'pfLevel')),
        COMMETHOD([], HRESULT, 'SetMute', (['in'], ctypes.c_int, 'bMute'), (['in'], POINTER(GUID), 'pguidEventContext')),
        COMMETHOD([], HRESULT, 'GetMute', (['out', 'retval'], POINTER(ctypes.c_int), 'pbMute')),
    ]


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
        
        # Audio session manager
        self.audio_sessions = {}
        self.volume_slider_window = None
        
        # Get monitors
        self.monitors = self.get_monitors()
        
        # Initialize pycaw
        self.init_audio()
        
        self.setup_styles()
        self.setup_ui()
        self.refresh_windows()
    
    def init_audio(self):
        """Initialize audio control via pycaw"""
        try:
            from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
            self.AudioUtilities = AudioUtilities
            self.ISimpleAudioVolume = ISimpleAudioVolume
            self.audio_available = True
        except ImportError:
            print("pycaw not available - audio control disabled")
            print("Install with: pip install pycaw")
            self.audio_available = False
    
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
            print(f"Error getting audio session: {e}")
        return None
    
    def get_app_volume(self, pid):
        """Get volume level for an app (0.0 to 1.0)"""
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                return volume.GetMasterVolume()
            except:
                pass
        return None
    
    def set_app_volume(self, pid, level):
        """Set volume level for an app (0.0 to 1.0)"""
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                volume.SetMasterVolume(level, None)
                return True
            except:
                pass
        return False
    
    def get_app_mute(self, pid):
        """Get mute state for an app"""
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                return volume.GetMute()
            except:
                pass
        return None
    
    def set_app_mute(self, pid, mute):
        """Set mute state for an app"""
        session = self.get_audio_session_for_pid(pid)
        if session:
            try:
                volume = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                volume.SetMute(mute, None)
                return True
            except:
                pass
        return False
    
    def get_system_volume(self):
        """Get system master volume (0.0 to 1.0)"""
        try:
            from pycaw.pycaw import AudioUtilities
            from pycaw.pycaw import IAudioEndpointVolume
            from comtypes import CLSCTX_ALL
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            return volume.GetMasterVolumeLevelScalar()
        except Exception as e:
            print(f"Error getting system volume: {e}")
        return 0.5
    
    def set_system_volume(self, level):
        """Set system master volume (0.0 to 1.0)"""
        try:
            from pycaw.pycaw import AudioUtilities
            from pycaw.pycaw import IAudioEndpointVolume
            from comtypes import CLSCTX_ALL
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevelScalar(level, None)
            return True
        except Exception as e:
            print(f"Error setting system volume: {e}")
        return False
    
    def get_system_mute(self):
        """Get system mute state"""
        try:
            from pycaw.pycaw import AudioUtilities
            from pycaw.pycaw import IAudioEndpointVolume
            from comtypes import CLSCTX_ALL
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            return volume.GetMute()
        except:
            pass
        return False
    
    def set_system_mute(self, mute):
        """Set system mute state"""
        try:
            from pycaw.pycaw import AudioUtilities
            from pycaw.pycaw import IAudioEndpointVolume
            from comtypes import CLSCTX_ALL
            
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMute(1 if mute else 0, None)
            return True
        except:
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
        # Store current pin state
        was_pinned = self.pin_to_top.get()
        
        # Force to top
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.update_idletasks()
        
        # If not pinned, schedule restore
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
        
        ttk.Button(select_frame, text="All", width=4, style='Small.TButton',
                   command=self.select_all).pack(side=tk.LEFT, padx=(8, 2))
        ttk.Button(select_frame, text="None", width=5, style='Small.TButton',
                   command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="Monitor", width=7, style='Small.TButton',
                   command=self.select_monitor).pack(side=tk.LEFT, padx=2)
        
        # Separator
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
        # Audio bulk controls
        audio_frame = ttk.Frame(main_frame)
        audio_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(audio_frame, text="Audio:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        # Mute/Unmute buttons for bulk operations
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
        
        # Volume button with right-click for slider
        self.bulk_volume_btn = tk.Button(audio_frame, text="🎚️ Vol",
                                          bg=self.colors['card'], fg=self.colors['fg'],
                                          font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                          activebackground=self.colors['card_hover'])
        self.bulk_volume_btn.pack(side=tk.LEFT)
        self.bulk_volume_btn.bind('<Button-3>', self.show_bulk_volume_slider)
        
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
        
        # Focus button
        focus_btn = tk.Button(actions, text="◉", command=lambda h=hwnd: self.focus_window(h), **btn_style)
        focus_btn.pack(side=tk.LEFT)
        
        # Min/Max toggle button
        minmax_btn = tk.Button(actions, text="□", command=lambda h=hwnd: self.toggle_minmax(h), **btn_style)
        minmax_btn.pack(side=tk.LEFT)
        
        # Audio mute toggle button
        is_muted = self.get_app_mute(pid) if pid else False
        audio_icon = "🔇" if is_muted else "🔊"
        audio_btn = tk.Button(actions, text=audio_icon, **btn_style)
        audio_btn.configure(command=lambda h=hwnd, p=pid, b=audio_btn: self.toggle_app_mute(h, p, b))
        audio_btn.bind('<Button-3>', lambda e, h=hwnd, p=pid, b=audio_btn: self.show_app_volume_slider(e, h, p, b))
        audio_btn.pack(side=tk.LEFT)
        
        # Store card references
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
    
    def show_app_volume_slider(self, event, hwnd, pid, btn):
        """Show volume slider for a specific app"""
        if not pid or not self.audio_available:
            self.status_var.set("No audio session for this window")
            return
            
        current_volume = self.get_app_volume(pid)
        if current_volume is None:
            self.status_var.set("No audio session for this window")
            return
        
        self.create_volume_slider(event, current_volume, 
                                   lambda v: self.set_app_volume(pid, v),
                                   lambda: self.update_audio_btn(hwnd, pid, btn))
    
    def show_bulk_volume_slider(self, event):
        """Show volume slider for bulk operation or system volume"""
        selected = self.get_selected_windows()
        
        if not selected:
            # No selection - control system volume
            current_volume = self.get_system_volume()
            self.create_volume_slider(event, current_volume,
                                       lambda v: self.set_system_volume(v),
                                       None, "System Volume")
        else:
            # Control selected windows
            # Get average volume of selected windows
            volumes = []
            for hwnd in selected:
                pid = self.window_pids.get(hwnd)
                if pid:
                    vol = self.get_app_volume(pid)
                    if vol is not None:
                        volumes.append(vol)
            
            current_volume = sum(volumes) / len(volumes) if volumes else 0.5
            self.create_volume_slider(event, current_volume,
                                       lambda v: self.set_selected_volumes(v),
                                       self.update_all_audio_btns, f"{len(selected)} Windows")
    
    def create_volume_slider(self, event, initial_volume, on_change, on_release=None, title="Volume"):
        """Create a floating volume slider window"""
        # Close any existing slider
        if self.volume_slider_window:
            self.volume_slider_window.destroy()
        
        # Create popup window
        slider_win = tk.Toplevel(self.root)
        slider_win.overrideredirect(True)
        slider_win.attributes('-topmost', True)
        
        # Calculate position near the button
        x = event.x_root - 20
        y = event.y_root + 10
        
        # Slider height is approximately twice a card height (80-100 pixels)
        slider_height = 160
        slider_width = 50
        
        slider_win.geometry(f"{slider_width}x{slider_height}+{x}+{y}")
        slider_win.configure(bg=self.colors['card'])
        
        self.volume_slider_window = slider_win
        
        # Title label
        title_lbl = tk.Label(slider_win, text=title, bg=self.colors['card'],
                             fg=self.colors['muted'], font=('Segoe UI', 7))
        title_lbl.pack(pady=(5, 0))
        
        # Volume percentage label
        vol_var = tk.StringVar(value=f"{int(initial_volume * 100)}%")
        vol_label = tk.Label(slider_win, textvariable=vol_var, bg=self.colors['card'],
                             fg=self.colors['fg'], font=('Segoe UI', 9, 'bold'))
        vol_label.pack(pady=(2, 5))
        
        # Create canvas for custom slider
        canvas_height = slider_height - 60
        canvas = tk.Canvas(slider_win, width=30, height=canvas_height,
                           bg=self.colors['slider_bg'], highlightthickness=0)
        canvas.pack(pady=(0, 5))
        
        # Draw slider track
        track_x = 15
        track_top = 5
        track_bottom = canvas_height - 5
        track_height = track_bottom - track_top
        
        # Track background
        canvas.create_rectangle(track_x - 3, track_top, track_x + 3, track_bottom,
                                fill=self.colors['slider_bg'], outline=self.colors['border'])
        
        # Calculate initial handle position
        handle_y = track_bottom - (initial_volume * track_height)
        
        # Filled portion
        fill_rect = canvas.create_rectangle(track_x - 3, handle_y, track_x + 3, track_bottom,
                                            fill=self.colors['slider_fg'], outline='')
        
        # Handle
        handle = canvas.create_oval(track_x - 8, handle_y - 8, track_x + 8, handle_y + 8,
                                    fill=self.colors['accent'], outline=self.colors['fg'])
        
        # Dragging state
        dragging = {'active': False}
        
        def update_slider(y_pos):
            # Clamp y position
            y_pos = max(track_top, min(track_bottom, y_pos))
            
            # Calculate volume (inverted - top is high, bottom is low)
            volume = (track_bottom - y_pos) / track_height
            volume = max(0.0, min(1.0, volume))
            
            # Update visual
            canvas.coords(handle, track_x - 8, y_pos - 8, track_x + 8, y_pos + 8)
            canvas.coords(fill_rect, track_x - 3, y_pos, track_x + 3, track_bottom)
            
            # Update label
            vol_var.set(f"{int(volume * 100)}%")
            
            # Apply volume change
            if on_change:
                on_change(volume)
        
        def on_press(e):
            dragging['active'] = True
            update_slider(e.y)
        
        def on_drag(e):
            if dragging['active']:
                update_slider(e.y)
        
        def on_release(e):
            dragging['active'] = False
            if on_release:
                on_release()
        
        canvas.bind('<Button-1>', on_press)
        canvas.bind('<B1-Motion>', on_drag)
        canvas.bind('<ButtonRelease-1>', on_release)
        
        # Start dragging immediately from current position
        # Simulate initial press at handle position
        dragging['active'] = True
        
        # Close when clicking outside
        def close_on_outside(e):
            # Check if click is outside the slider window
            try:
                if slider_win.winfo_exists():
                    x, y = e.x_root, e.y_root
                    wx = slider_win.winfo_rootx()
                    wy = slider_win.winfo_rooty()
                    ww = slider_win.winfo_width()
                    wh = slider_win.winfo_height()
                    
                    if not (wx <= x <= wx + ww and wy <= y <= wy + wh):
                        if on_release:
                            on_release()
                        slider_win.destroy()
                        self.volume_slider_window = None
            except:
                pass
        
        # Bind to root window
        self.root.bind('<Button-1>', close_on_outside, add='+')
        
        # Also close on escape
        def close_on_escape(e):
            if on_release:
                on_release()
            slider_win.destroy()
            self.volume_slider_window = None
        
        slider_win.bind('<Escape>', close_on_escape)
        slider_win.focus_set()
    
    def update_audio_btn(self, hwnd, pid, btn):
        """Update audio button icon based on mute state"""
        if hwnd in self.window_cards:
            is_muted = self.get_app_mute(pid)
            if is_muted is not None:
                btn.configure(text="🔇" if is_muted else "🔊")
    
    def update_all_audio_btns(self):
        """Update all audio button icons"""
        for hwnd, card_data in self.window_cards.items():
            pid = self.window_pids.get(hwnd)
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
            # Mute system
            if self.set_system_mute(True):
                self.status_var.set("System muted")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid and self.set_app_mute(pid, True):
                count += 1
                # Update button
                if hwnd in self.window_cards:
                    self.window_cards[hwnd]['audio_btn'].configure(text="🔇")
        
        self.status_var.set(f"Muted {count} window(s)")
    
    def bulk_unmute(self):
        """Unmute selected windows or system"""
        selected = self.get_selected_windows()
        
        if not selected:
            # Unmute system
            if self.set_system_mute(False):
                self.status_var.set("System unmuted")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid and self.set_app_mute(pid, False):
                count += 1
                # Update button
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
    
    # Check for pycaw (optional but recommended)
    try:
        import pycaw
    except ImportError:
        print("Note: pycaw not installed. Audio control will be limited.")
        print("Install with: pip install pycaw")
        
    app = WindowManager()
    app.run()


if __name__ == "__main__":
    main()