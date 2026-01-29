import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes
import win32gui
import win32con
import win32process
import win32api
from pycaw.pycaw import AudioUtilities, IMMDeviceEnumerator, EDataFlow, DEVICE_STATE
from pycaw.constants import CLSID_MMDeviceEnumerator
from comtypes import CLSCTX_ALL, CoCreateInstance
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
        self.root.geometry("800x600")
        self.root.attributes('-topmost', True)
        self.root.configure(bg='#2b2b2b')
        
        # Track selected windows in order
        self.selected_windows = OrderedDict()
        self.window_checkboxes = {}
        self.windows_list = []
        
        # Get screen dimensions
        self.screen_width = win32api.GetSystemMetrics(0)
        self.screen_height = win32api.GetSystemMetrics(1)
        
        self.setup_ui()
        self.refresh_windows()
        
    def setup_ui(self):
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#2b2b2b')
        style.configure('TLabel', background='#2b2b2b', foreground='white')
        style.configure('TButton', padding=6)
        style.configure('TCheckbutton', background='#2b2b2b', foreground='white')
        
        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = ttk.Label(main_frame, text="Window Manager", font=('Segoe UI', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Control buttons frame
        control_frame = ttk.Frame(main_frame)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        refresh_btn = ttk.Button(control_frame, text="🔄 Refresh", command=self.refresh_windows)
        refresh_btn.pack(side=tk.LEFT, padx=5)
        
        # Split options frame
        split_frame = ttk.LabelFrame(main_frame, text="Split Screen (First 2 Selected Windows)")
        split_frame.pack(fill=tk.X, pady=(0, 10))
        
        split_v_btn = ttk.Button(split_frame, text="↔ Split Vertical (Side by Side)", 
                                  command=self.split_vertical)
        split_v_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        split_h_btn = ttk.Button(split_frame, text="↕ Split Horizontal (Top/Bottom)", 
                                  command=self.split_horizontal)
        split_h_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        # Selection info
        self.selection_label = ttk.Label(split_frame, text="Selected: 0 windows", font=('Segoe UI', 9))
        self.selection_label.pack(side=tk.RIGHT, padx=10)
        
        # Audio device selection frame
        audio_frame = ttk.LabelFrame(main_frame, text="Audio Output Device")
        audio_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.audio_devices = self.get_audio_devices()
        self.audio_var = tk.StringVar()
        
        if self.audio_devices:
            self.audio_var.set(list(self.audio_devices.keys())[0])
            
        self.audio_combo = ttk.Combobox(audio_frame, textvariable=self.audio_var, 
                                         values=list(self.audio_devices.keys()), width=50)
        self.audio_combo.pack(side=tk.LEFT, padx=10, pady=5)
        
        apply_audio_btn = ttk.Button(audio_frame, text="Apply to Selected", 
                                      command=self.apply_audio_device)
        apply_audio_btn.pack(side=tk.LEFT, padx=10, pady=5)
        
        refresh_audio_btn = ttk.Button(audio_frame, text="🔄", width=3,
                                        command=self.refresh_audio_devices)
        refresh_audio_btn.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Windows list frame
        list_frame = ttk.LabelFrame(main_frame, text="Open Windows")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas for scrollable window list
        canvas = tk.Canvas(list_frame, bg='#3c3c3c', highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Enable mousewheel scrolling
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
    def get_audio_devices(self):
        """Get all audio output devices"""
        devices = {}
        try:
            # Get audio endpoints
            deviceEnumerator = CoCreateInstance(
                CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                CLSCTX_ALL
            )
            
            # Enumerate render devices (speakers/output)
            collection = deviceEnumerator.EnumAudioEndpoints(EDataFlow.eRender.value, DEVICE_STATE.ACTIVE.value)
            count = collection.GetCount()
            
            for i in range(count):
                device = collection.Item(i)
                device_id = device.GetId()
                props = device.OpenPropertyStore(0)
                
                # Get friendly name (PKEY_Device_FriendlyName)
                try:
                    from pycaw.pycaw import PKEY_Device_FriendlyName
                    name = props.GetValue(PKEY_Device_FriendlyName).GetValue()
                except:
                    name = f"Audio Device {i+1}"
                    
                devices[name] = device_id
                
        except Exception as e:
            print(f"Error getting audio devices: {e}")
            devices["Default Audio Device"] = None
            
        return devices
    
    def refresh_audio_devices(self):
        """Refresh the list of audio devices"""
        self.audio_devices = self.get_audio_devices()
        self.audio_combo['values'] = list(self.audio_devices.keys())
        if self.audio_devices:
            self.audio_var.set(list(self.audio_devices.keys())[0])
        self.status_var.set("Audio devices refreshed")
        
    def is_real_window(self, hwnd):
        """Check if a window is a real, visible application window"""
        if not win32gui.IsWindowVisible(hwnd):
            return False
            
        if not win32gui.GetWindowText(hwnd):
            return False
            
        # Check if window is cloaked (hidden by DWM)
        try:
            cloaked = ctypes.c_int(0)
            ctypes.windll.dwmapi.DwmGetWindowAttribute(
                hwnd, DWMWA_CLOAKED, ctypes.byref(cloaked), ctypes.sizeof(cloaked))
            if cloaked.value:
                return False
        except:
            pass
            
        # Check window style
        ex_style = win32gui.GetWindowLong(hwnd, GWL_EXSTYLE)
        if ex_style & WS_EX_TOOLWINDOW:
            return False
            
        # Get window rect
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
            return process.name()
        except:
            return "Unknown"
    
    def refresh_windows(self):
        """Refresh the list of windows"""
        # Clear existing widgets
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.window_checkboxes.clear()
        self.windows_list.clear()
        
        # Preserve selection order for windows that still exist
        old_selection = list(self.selected_windows.keys())
        self.selected_windows.clear()
        
        def enum_callback(hwnd, windows):
            if self.is_real_window(hwnd):
                title = win32gui.GetWindowText(hwnd)
                process = self.get_process_name(hwnd)
                windows.append((hwnd, title, process))
            return True
            
        windows = []
        win32gui.EnumWindows(enum_callback, windows)
        self.windows_list = windows
        
        # Exclude our own window
        our_hwnd = self.root.winfo_id()
        windows = [(h, t, p) for h, t, p in windows if h != our_hwnd]
        
        # Create window entries
        for i, (hwnd, title, process) in enumerate(windows):
            frame = ttk.Frame(self.scrollable_frame)
            frame.pack(fill=tk.X, padx=5, pady=2)
            
            # Checkbox variable
            var = tk.BooleanVar(value=hwnd in old_selection)
            
            # If was previously selected, re-add to selection
            if hwnd in old_selection:
                self.selected_windows[hwnd] = True
            
            # Checkbox
            cb = ttk.Checkbutton(frame, variable=var, 
                                  command=lambda h=hwnd, v=var: self.on_checkbox_changed(h, v))
            cb.pack(side=tk.LEFT)
            
            self.window_checkboxes[hwnd] = var
            
            # Window info
            display_title = title[:50] + "..." if len(title) > 50 else title
            info_text = f"{display_title} ({process})"
            
            label = ttk.Label(frame, text=info_text, width=60, anchor='w')
            label.pack(side=tk.LEFT, padx=5)
            
            # Quick action buttons
            focus_btn = ttk.Button(frame, text="Focus", width=6,
                                    command=lambda h=hwnd: self.focus_window(h))
            focus_btn.pack(side=tk.RIGHT, padx=2)
            
            minimize_btn = ttk.Button(frame, text="Min", width=4,
                                       command=lambda h=hwnd: self.minimize_window(h))
            minimize_btn.pack(side=tk.RIGHT, padx=2)
            
            maximize_btn = ttk.Button(frame, text="Max", width=4,
                                       command=lambda h=hwnd: self.maximize_window(h))
            maximize_btn.pack(side=tk.RIGHT, padx=2)
            
        self.update_selection_label()
        self.status_var.set(f"Found {len(windows)} windows")
        
    def on_checkbox_changed(self, hwnd, var):
        """Handle checkbox state change to track selection order"""
        if var.get():
            # Add to selection (will be added at the end, maintaining order)
            self.selected_windows[hwnd] = True
        else:
            # Remove from selection
            if hwnd in self.selected_windows:
                del self.selected_windows[hwnd]
                
        self.update_selection_label()
        
    def update_selection_label(self):
        """Update the selection count label"""
        count = len(self.selected_windows)
        if count >= 2:
            # Show which windows will be used for split
            selected_list = list(self.selected_windows.keys())
            first_title = win32gui.GetWindowText(selected_list[0])[:20]
            second_title = win32gui.GetWindowText(selected_list[1])[:20]
            self.selection_label.config(
                text=f"Selected: {count} | Split: '{first_title}...' + '{second_title}...'"
            )
        else:
            self.selection_label.config(text=f"Selected: {count} windows")
    
    def get_selected_windows(self):
        """Get list of selected window handles in selection order"""
        return list(self.selected_windows.keys())
    
    def focus_window(self, hwnd):
        """Bring a window to focus"""
        try:
            # Restore if minimized
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            self.status_var.set(f"Focused: {win32gui.GetWindowText(hwnd)[:30]}")
        except Exception as e:
            self.status_var.set(f"Error focusing window: {e}")
            
    def minimize_window(self, hwnd):
        """Minimize a window"""
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            self.status_var.set(f"Minimized: {win32gui.GetWindowText(hwnd)[:30]}")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            
    def maximize_window(self, hwnd):
        """Maximize a window"""
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
            self.status_var.set(f"Maximized: {win32gui.GetWindowText(hwnd)[:30]}")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
    
    def split_vertical(self):
        """Split first two selected windows side by side (vertical split)"""
        selected = self.get_selected_windows()
        
        if len(selected) < 2:
            self.status_var.set("Please select at least 2 windows for split")
            return
            
        # Get first two selected windows
        hwnd1, hwnd2 = selected[0], selected[1]
        
        # Calculate dimensions (account for taskbar)
        work_area = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0)))['Work']
        work_width = work_area[2] - work_area[0]
        work_height = work_area[3] - work_area[1]
        work_left = work_area[0]
        work_top = work_area[1]
        
        half_width = work_width // 2
        
        try:
            # Restore windows if maximized/minimized
            win32gui.ShowWindow(hwnd1, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd2, win32con.SW_RESTORE)
            
            # First window (left)
            win32gui.SetWindowPos(hwnd1, win32con.HWND_TOP, 
                                   work_left, work_top, 
                                   half_width, work_height, 
                                   win32con.SWP_SHOWWINDOW)
            
            # Second window (right)
            win32gui.SetWindowPos(hwnd2, win32con.HWND_TOP, 
                                   work_left + half_width, work_top, 
                                   half_width, work_height, 
                                   win32con.SWP_SHOWWINDOW)
            
            title1 = win32gui.GetWindowText(hwnd1)[:20]
            title2 = win32gui.GetWindowText(hwnd2)[:20]
            self.status_var.set(f"Split vertical: '{title1}' (left) | '{title2}' (right)")
            
        except Exception as e:
            self.status_var.set(f"Error splitting windows: {e}")
    
    def split_horizontal(self):
        """Split first two selected windows top/bottom (horizontal split)"""
        selected = self.get_selected_windows()
        
        if len(selected) < 2:
            self.status_var.set("Please select at least 2 windows for split")
            return
            
        # Get first two selected windows
        hwnd1, hwnd2 = selected[0], selected[1]
        
        # Calculate dimensions (account for taskbar)
        work_area = win32api.GetMonitorInfo(win32api.MonitorFromPoint((0, 0)))['Work']
        work_width = work_area[2] - work_area[0]
        work_height = work_area[3] - work_area[1]
        work_left = work_area[0]
        work_top = work_area[1]
        
        half_height = work_height // 2
        
        try:
            # Restore windows if maximized/minimized
            win32gui.ShowWindow(hwnd1, win32con.SW_RESTORE)
            win32gui.ShowWindow(hwnd2, win32con.SW_RESTORE)
            
            # First window (top)
            win32gui.SetWindowPos(hwnd1, win32con.HWND_TOP, 
                                   work_left, work_top, 
                                   work_width, half_height, 
                                   win32con.SWP_SHOWWINDOW)
            
            # Second window (bottom)
            win32gui.SetWindowPos(hwnd2, win32con.HWND_TOP, 
                                   work_left, work_top + half_height, 
                                   work_width, half_height, 
                                   win32con.SWP_SHOWWINDOW)
            
            title1 = win32gui.GetWindowText(hwnd1)[:20]
            title2 = win32gui.GetWindowText(hwnd2)[:20]
            self.status_var.set(f"Split horizontal: '{title1}' (top) | '{title2}' (bottom)")
            
        except Exception as e:
            self.status_var.set(f"Error splitting windows: {e}")
    
    def apply_audio_device(self):
        """Apply selected audio device to selected windows' processes"""
        selected = self.get_selected_windows()
        
        if not selected:
            self.status_var.set("No windows selected")
            return
            
        device_name = self.audio_var.get()
        
        if not device_name:
            self.status_var.set("No audio device selected")
            return
        
        # Note: Per-application audio routing requires Windows 10 1803+
        # and is complex to implement. This shows the concept.
        try:
            sessions = AudioUtilities.GetAllSessions()
            affected_count = 0
            
            for hwnd in selected:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    
                    # Find audio session for this process
                    for session in sessions:
                        if session.Process and session.Process.pid == pid:
                            # Get the session control
                            volume = session.SimpleAudioVolume
                            if volume:
                                # We can control volume, but device routing
                                # requires different API (AudioPolicyConfig)
                                affected_count += 1
                                
                except Exception as e:
                    print(f"Error processing window {hwnd}: {e}")
                    
            self.status_var.set(
                f"Audio device '{device_name}' - Found {affected_count} audio sessions. "
                "Note: Full per-app audio routing requires Windows Audio Policy API."
            )
            
        except Exception as e:
            self.status_var.set(f"Error applying audio device: {e}")
            
        # Show info about audio routing limitation
        self.show_audio_info()
    
    def show_audio_info(self):
        """Show information about audio routing capabilities"""
        info_window = tk.Toplevel(self.root)
        info_window.title("Audio Routing Info")
        info_window.geometry("500x200")
        info_window.attributes('-topmost', True)
        info_window.configure(bg='#2b2b2b')
        
        info_text = """
Per-application audio device routing in Windows requires:

1. Windows 10 version 1803 or later
2. Access to the undocumented AudioPolicyConfig API
3. Or use of third-party tools like:
   - EarTrumpet (free, open-source)
   - Audio Router
   - VoiceMeeter

The current implementation can detect audio sessions but 
full device routing requires additional system integration.

For quick audio device switching, you can use:
Win + Ctrl + V (Windows 11) or the Sound settings.
        """
        
        label = ttk.Label(info_window, text=info_text, justify=tk.LEFT,
                          background='#2b2b2b', foreground='white')
        label.pack(padx=20, pady=20)
        
        close_btn = ttk.Button(info_window, text="Close", command=info_window.destroy)
        close_btn.pack(pady=10)
    
    def run(self):
        """Run the application"""
        self.root.mainloop()


def main():
    """Main entry point"""
    # Check for required modules
    required_modules = ['win32gui', 'win32con', 'win32process', 'win32api', 'pycaw', 'psutil']
    missing = []
    
    for module in required_modules:
        try:
            __import__(module)
        except ImportError:
            missing.append(module)
    
    if missing:
        print("Missing required modules. Please install them:")
        print(f"pip install pywin32 pycaw psutil")
        print(f"\nMissing: {', '.join(missing)}")
        return
        
    app = WindowManager()
    app.run()


if __name__ == "__main__":
    main()