import tkinter as tk
from tkinter import ttk
import ctypes
from ctypes import wintypes, cast, POINTER
import win32gui
import win32con
import win32process
import win32api
import psutil
import os
import time
import traceback
import atexit
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
        self.root.geometry("420x580")
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
        
        # Debugging
        self.debug_mode = tk.BooleanVar(value=False)
        self.debug_log = []
        self.debug_dir = "windows-manager-debugging"
        
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
        self.slider_start_y = None
        self.slider_start_volume = None
        self.current_slider_volume = None
        self.volume_slider_on_change = None
        self.volume_slider_on_close = None
        self.slider_canvas = None
        self.slider_is_dragging = False  # Track if we're actively dragging
        
        # Get monitors
        self.monitors = self.get_monitors()
        
        # Initialize audio
        self.init_audio()
        
        self.setup_styles()
        self.setup_ui()
        self.refresh_windows()
        
        # Register cleanup on exit
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        atexit.register(self.export_debug_log_on_exit)
    
    def on_closing(self):
        """Handle window close - export debug log if debug mode was used"""
        if self.debug_log:  # Only export if there's something to log
            self.log_debug("Application closing")
            self.export_debug_log()
        self.root.destroy()
    
    def export_debug_log_on_exit(self):
        """Export debug log on exit (via atexit)"""
        # This is a backup in case on_closing isn't called
        pass
    
    def log_debug(self, message, level="INFO"):
        """Add a message to the debug log"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] [{level}] {message}"
        self.debug_log.append(entry)
        
        # Keep log from growing too large
        if len(self.debug_log) > 1000:
            self.debug_log = self.debug_log[-500:]
        
        # Print to console if debug mode is on
        if self.debug_mode.get():
            print(entry)
    
    def export_debug_log(self):
        """Export debug log to a file"""
        if not self.debug_log:
            return None
            
        # Create debug directory if it doesn't exist
        if not os.path.exists(self.debug_dir):
            os.makedirs(self.debug_dir)
        
        # Generate filename with unix timestamp
        unix_time = int(time.time())
        filename = os.path.join(self.debug_dir, f"debug_{unix_time}.log")
        
        # Gather system info
        system_info = self.gather_system_info()
        
        # Write log file
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write("=" * 60 + "\n")
                f.write("WINDOW MANAGER DEBUG LOG\n")
                f.write(f"Exported: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Unix Timestamp: {unix_time}\n")
                f.write("=" * 60 + "\n\n")
                
                f.write("SYSTEM INFORMATION\n")
                f.write("-" * 40 + "\n")
                for key, value in system_info.items():
                    f.write(f"{key}: {value}\n")
                f.write("\n")
                
                f.write("DEBUG LOG\n")
                f.write("-" * 40 + "\n")
                for entry in self.debug_log:
                    f.write(entry + "\n")
            
            if hasattr(self, 'status_var'):
                self.status_var.set(f"Debug log exported: {filename}")
            return filename
        except Exception as e:
            print(f"Failed to export debug log: {e}")
            return None
    
    def gather_system_info(self):
        """Gather system information for debugging"""
        info = {}
        
        # Python version
        import sys
        info["Python Version"] = sys.version
        
        # OS info
        import platform
        info["OS"] = platform.system()
        info["OS Version"] = platform.version()
        info["OS Release"] = platform.release()
        info["Machine"] = platform.machine()
        
        # Module versions
        try:
            import pycaw
            info["pycaw Version"] = getattr(pycaw, '__version__', 'unknown')
        except ImportError:
            info["pycaw Version"] = "NOT INSTALLED"
        
        try:
            import comtypes
            info["comtypes Version"] = getattr(comtypes, '__version__', 'unknown')
        except ImportError:
            info["comtypes Version"] = "NOT INSTALLED"
        
        try:
            info["psutil Version"] = psutil.__version__
        except:
            info["psutil Version"] = "unknown"
        
        try:
            import win32api
            info["pywin32"] = "installed"
        except:
            info["pywin32"] = "NOT INSTALLED"
        
        # Audio status
        info["Audio Available"] = self.audio_available
        info["Volume Interface Type"] = str(type(self.volume_interface)) if self.volume_interface else "None"
        info["Audio Device Count"] = len(self.audio_devices) if hasattr(self, 'audio_devices') else 0
        
        # List audio devices
        if hasattr(self, 'audio_devices'):
            for i, (device_id, device_name) in enumerate(self.audio_devices.items()):
                info[f"Audio Device {i}"] = device_name
        
        # Get fresh audio session info for debugging
        if self.AudioUtilities:
            try:
                sessions = self.get_all_audio_sessions_all_devices()
                info["Audio Session Count"] = len(sessions) if sessions else 0
                for i, (pid, session_info) in enumerate(list(sessions.items())[:10]):  # Limit to 10
                    info[f"Session {i}"] = f"PID={pid}, Name={session_info.get('name', 'unknown')}"
            except Exception as e:
                info["Audio Session Error"] = str(e)
        
        # Monitor info
        info["Monitor Count"] = len(self.monitors)
        for i, mon in enumerate(self.monitors):
            info[f"Monitor {i+1}"] = f"{mon['name']} - {mon['resolution']}"
        
        # Window count
        info["Tracked Windows"] = len(self.windows_list)
        
        return info
    
    def init_audio(self):
        """Initialize audio control via pycaw"""
        self.audio_available = False
        self.volume_interface = None
        self.AudioUtilities = None
        self.ISimpleAudioVolume = None
        self.IAudioEndpointVolume = None
        self._speakers = None  # Keep reference to speakers device
        self.audio_devices = {}  # Store all audio devices
        self._all_endpoints = []  # Store all audio endpoints
        
        self.log_debug("Initializing audio...")
        
        try:
            # Import pycaw modules
            try:
                from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume
                self.AudioUtilities = AudioUtilities
                self.ISimpleAudioVolume = ISimpleAudioVolume
                self.IAudioEndpointVolume = IAudioEndpointVolume
                self.log_debug("Imported from pycaw.pycaw")
            except ImportError:
                try:
                    from pycaw import AudioUtilities, ISimpleAudioVolume, IAudioEndpointVolume
                    self.AudioUtilities = AudioUtilities
                    self.ISimpleAudioVolume = ISimpleAudioVolume
                    self.IAudioEndpointVolume = IAudioEndpointVolume
                    self.log_debug("Imported from pycaw")
                except ImportError as e:
                    self.log_debug(f"Failed to import AudioUtilities: {e}", "ERROR")
                    return
            
            # Import comtypes for CLSCTX_ALL
            try:
                from comtypes import CLSCTX_ALL
                self.CLSCTX_ALL = CLSCTX_ALL
            except ImportError:
                self.CLSCTX_ALL = 0x17  # Fallback value
            
            # Get all audio devices
            self._enumerate_audio_devices()
            
            # Get default speakers device - keep reference
            self._speakers = self.AudioUtilities.GetSpeakers()
            self.log_debug(f"GetSpeakers() returned: {type(self._speakers)}")
            
            if self._speakers is None:
                self.log_debug("GetSpeakers() returned None", "ERROR")
                return
            
            # Try EndpointVolume property (new pycaw API)
            if hasattr(self._speakers, 'EndpointVolume'):
                self.log_debug("Found EndpointVolume property (new pycaw API)")
                try:
                    self.volume_interface = self._speakers.EndpointVolume
                    test_vol = self.volume_interface.GetMasterVolumeLevelScalar()
                    self.log_debug(f"Test read volume: {test_vol}")
                    self.audio_available = True
                    self.log_debug("Audio initialized via EndpointVolume property", "SUCCESS")
                    return
                except Exception as e:
                    self.log_debug(f"EndpointVolume property failed: {e}", "ERROR")
            
            # Try _volume attribute
            if hasattr(self._speakers, '_volume'):
                self.log_debug("Found _volume attribute")
                try:
                    self.volume_interface = self._speakers._volume
                    test_vol = self.volume_interface.GetMasterVolumeLevelScalar()
                    self.log_debug(f"Test read volume via _volume: {test_vol}")
                    self.audio_available = True
                    self.log_debug("Audio initialized via _volume attribute", "SUCCESS")
                    return
                except Exception as e:
                    self.log_debug(f"_volume attribute failed: {e}", "ERROR")
            
            # Fallback: Try older Activate method
            if hasattr(self._speakers, 'Activate'):
                self.log_debug("Trying Activate() method (old API)")
                try:
                    interface = self._speakers.Activate(
                        self.IAudioEndpointVolume._iid_, self.CLSCTX_ALL, None)
                    self.volume_interface = cast(interface, POINTER(self.IAudioEndpointVolume))
                    test_vol = self.volume_interface.GetMasterVolumeLevelScalar()
                    self.log_debug(f"Test read volume via Activate: {test_vol}")
                    self.audio_available = True
                    self.log_debug("Audio initialized via Activate()", "SUCCESS")
                    return
                except Exception as e:
                    self.log_debug(f"Activate() failed: {e}", "ERROR")
            
            self.log_debug("All audio initialization methods exhausted", "ERROR")
            
        except ImportError as e:
            self.log_debug(f"Import error: {e}", "ERROR")
        except Exception as e:
            self.log_debug(f"Unexpected error initializing audio: {e}", "ERROR")
            self.log_debug(traceback.format_exc(), "ERROR")
    
    def _enumerate_audio_devices(self):
        """Enumerate all audio output devices"""
        self.audio_devices = {}
        self._all_endpoints = []
        
        try:
            # Try to get all audio devices
            if hasattr(self.AudioUtilities, 'GetAllDevices'):
                devices = self.AudioUtilities.GetAllDevices()
                for device in devices:
                    try:
                        device_id = str(device.id) if hasattr(device, 'id') else str(id(device))
                        device_name = device.FriendlyName if hasattr(device, 'FriendlyName') else "Unknown"
                        self.audio_devices[device_id] = device_name
                        self._all_endpoints.append(device)
                        self.log_debug(f"Found audio device: {device_name}")
                    except:
                        pass
            
            # Also try GetSpeakers for each device if possible
            try:
                from pycaw.pycaw import AudioUtilities as AU
                # Get all render devices (speakers/headphones)
                from comtypes import CLSCTX_ALL
                from pycaw.pycaw import EDataFlow, DEVICE_STATE
                
                if hasattr(AU, 'GetAllSessions'):
                    # This approach gets sessions from all devices
                    pass
                    
            except Exception as e:
                self.log_debug(f"Extended device enumeration failed: {e}", "WARNING")
                
        except Exception as e:
            self.log_debug(f"Error enumerating audio devices: {e}", "ERROR")
    
    def get_all_audio_sessions_all_devices(self):
        """Get audio sessions from ALL audio devices, not just the default"""
        sessions_by_pid = {}
        
        if not self.AudioUtilities:
            return sessions_by_pid
        
        try:
            # Method 1: Standard GetAllSessions (default device)
            sessions = self.AudioUtilities.GetAllSessions()
            if sessions:
                for session in sessions:
                    proc = session.Process
                    if proc:
                        sessions_by_pid[proc.pid] = {
                            'session': session,
                            'name': proc.name(),
                            'device': 'default'
                        }
            
            # Method 2: Try to get sessions from all devices using lower-level API
            try:
                from comtypes import CLSCTX_ALL, CoCreateInstance, GUID
                from pycaw.pycaw import IMMDeviceEnumerator, EDataFlow, DEVICE_STATE
                
                # Create device enumerator
                CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')
                IID_IMMDeviceEnumerator = GUID('{A95664D2-9614-4F35-A746-DE8DB63617E6}')
                
                try:
                    enumerator = CoCreateInstance(
                        CLSID_MMDeviceEnumerator,
                        IMMDeviceEnumerator,
                        CLSCTX_ALL
                    )
                    
                    # Get all render (output) devices
                    if hasattr(enumerator, 'EnumAudioEndpoints'):
                        collection = enumerator.EnumAudioEndpoints(
                            EDataFlow.eRender.value if hasattr(EDataFlow, 'eRender') else 0,
                            DEVICE_STATE.ACTIVE.value if hasattr(DEVICE_STATE, 'ACTIVE') else 1
                        )
                        
                        if collection:
                            count = collection.GetCount()
                            for i in range(count):
                                try:
                                    device = collection.Item(i)
                                    # Get sessions for this device
                                    if hasattr(device, 'Activate'):
                                        from pycaw.pycaw import IAudioSessionManager2
                                        mgr = device.Activate(
                                            IAudioSessionManager2._iid_,
                                            CLSCTX_ALL,
                                            None
                                        )
                                        if mgr:
                                            session_enum = mgr.GetSessionEnumerator()
                                            if session_enum:
                                                sess_count = session_enum.GetCount()
                                                for j in range(sess_count):
                                                    try:
                                                        sess_ctrl = session_enum.GetSession(j)
                                                        if sess_ctrl:
                                                            pid = sess_ctrl.GetProcessId()
                                                            if pid and pid > 0 and pid not in sessions_by_pid:
                                                                try:
                                                                    proc = psutil.Process(pid)
                                                                    sessions_by_pid[pid] = {
                                                                        'session_ctrl': sess_ctrl,
                                                                        'name': proc.name(),
                                                                        'device': f'device_{i}'
                                                                    }
                                                                except:
                                                                    pass
                                                    except:
                                                        pass
                                except:
                                    pass
                except Exception as e:
                    # This is expected to fail on some systems
                    pass
                    
            except ImportError:
                pass
            except Exception as e:
                pass
                
        except Exception as e:
            self.log_debug(f"Error getting audio sessions from all devices: {e}", "ERROR")
        
        return sessions_by_pid
    
    def get_fresh_audio_sessions(self, log_sessions=True):
        """Get fresh audio sessions from all devices - don't cache COM objects"""
        if not self.AudioUtilities:
            return {}
        
        sessions_by_pid = {}
        try:
            # Get sessions from all devices
            all_sessions = self.get_all_audio_sessions_all_devices()
            
            for pid, session_info in all_sessions.items():
                if 'session' in session_info:
                    sessions_by_pid[pid] = session_info['session']
                elif 'session_ctrl' in session_info:
                    # Wrap the session control in a compatible format
                    sessions_by_pid[pid] = session_info
            
            # Also get standard sessions as fallback
            sessions = self.AudioUtilities.GetAllSessions()
            if sessions:
                for session in sessions:
                    proc = session.Process
                    if proc and proc.pid not in sessions_by_pid:
                        sessions_by_pid[proc.pid] = session
                        
        except Exception as e:
            if log_sessions:
                self.log_debug(f"Error getting audio sessions: {e}", "ERROR")
        
        return sessions_by_pid
    
    def get_audio_session_for_pid(self, pid, log_lookup=True):
        """Get audio session for a specific process ID - always fetch fresh"""
        if not self.AudioUtilities or not pid:
            return None
        
        sessions = self.get_fresh_audio_sessions(log_sessions=log_lookup)
        
        # Direct PID match
        if pid in sessions:
            return sessions[pid]
        
        # Try to find by process tree (child processes)
        try:
            proc = psutil.Process(pid)
            # Check parent
            parent = proc.parent()
            if parent and parent.pid in sessions:
                if log_lookup:
                    self.log_debug(f"Found audio via parent PID {parent.pid}")
                return sessions[parent.pid]
            
            # Check children
            for child in proc.children(recursive=True):
                if child.pid in sessions:
                    if log_lookup:
                        self.log_debug(f"Found audio via child PID {child.pid}")
                    return sessions[child.pid]
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
        
        return None
    
    def get_app_volume(self, pid, log_lookup=True):
        """Get volume level for an app (0.0 to 1.0)"""
        session = self.get_audio_session_for_pid(pid, log_lookup=log_lookup)
        if not session:
            return None
        
        try:
            # Handle dict-wrapped session from multi-device enumeration
            if isinstance(session, dict):
                if 'session_ctrl' in session:
                    sess_ctrl = session['session_ctrl']
                    if hasattr(sess_ctrl, 'QueryInterface') and self.ISimpleAudioVolume:
                        volume_interface = sess_ctrl.QueryInterface(self.ISimpleAudioVolume)
                        return volume_interface.GetMasterVolume()
                return None
            
            # Try SimpleAudioVolume interface
            if hasattr(session, 'SimpleAudioVolume') and session.SimpleAudioVolume:
                return session.SimpleAudioVolume.GetMasterVolume()
            
            # Try _ctl.QueryInterface
            if hasattr(session, '_ctl') and self.ISimpleAudioVolume:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                return volume_interface.GetMasterVolume()
        except Exception as e:
            if log_lookup:
                self.log_debug(f"Error getting app volume for PID {pid}: {e}", "WARNING")
        return None
    
    def set_app_volume(self, pid, level, log_lookup=True):
        """Set volume level for an app (0.0 to 1.0)"""
        session = self.get_audio_session_for_pid(pid, log_lookup=log_lookup)
        if not session:
            return False
        
        try:
            level = max(0.0, min(1.0, float(level)))
            
            # Handle dict-wrapped session from multi-device enumeration
            if isinstance(session, dict):
                if 'session_ctrl' in session:
                    sess_ctrl = session['session_ctrl']
                    if hasattr(sess_ctrl, 'QueryInterface') and self.ISimpleAudioVolume:
                        volume_interface = sess_ctrl.QueryInterface(self.ISimpleAudioVolume)
                        volume_interface.SetMasterVolume(level, None)
                        return True
                return False
            
            # Try SimpleAudioVolume interface
            if hasattr(session, 'SimpleAudioVolume') and session.SimpleAudioVolume:
                session.SimpleAudioVolume.SetMasterVolume(level, None)
                return True
            
            # Try _ctl.QueryInterface
            if hasattr(session, '_ctl') and self.ISimpleAudioVolume:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                volume_interface.SetMasterVolume(level, None)
                return True
        except Exception as e:
            if log_lookup:
                self.log_debug(f"Error setting app volume for PID {pid}: {e}", "WARNING")
        return False
    
    def get_app_mute(self, pid, log_lookup=True):
        """Get mute state for an app"""
        session = self.get_audio_session_for_pid(pid, log_lookup=log_lookup)
        if not session:
            return None
        
        try:
            # Handle dict-wrapped session from multi-device enumeration
            if isinstance(session, dict):
                if 'session_ctrl' in session:
                    sess_ctrl = session['session_ctrl']
                    if hasattr(sess_ctrl, 'QueryInterface') and self.ISimpleAudioVolume:
                        volume_interface = sess_ctrl.QueryInterface(self.ISimpleAudioVolume)
                        return bool(volume_interface.GetMute())
                return None
            
            # Try SimpleAudioVolume interface
            if hasattr(session, 'SimpleAudioVolume') and session.SimpleAudioVolume:
                return bool(session.SimpleAudioVolume.GetMute())
            
            # Try _ctl.QueryInterface
            if hasattr(session, '_ctl') and self.ISimpleAudioVolume:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                return bool(volume_interface.GetMute())
        except Exception as e:
            if log_lookup:
                self.log_debug(f"Error getting app mute for PID {pid}: {e}", "WARNING")
        return None
    
    def set_app_mute(self, pid, mute, log_lookup=True):
        """Set mute state for an app"""
        session = self.get_audio_session_for_pid(pid, log_lookup=log_lookup)
        if not session:
            return False
        
        try:
            # Handle dict-wrapped session from multi-device enumeration
            if isinstance(session, dict):
                if 'session_ctrl' in session:
                    sess_ctrl = session['session_ctrl']
                    if hasattr(sess_ctrl, 'QueryInterface') and self.ISimpleAudioVolume:
                        volume_interface = sess_ctrl.QueryInterface(self.ISimpleAudioVolume)
                        volume_interface.SetMute(int(mute), None)
                        return True
                return False
            
            # Try SimpleAudioVolume interface
            if hasattr(session, 'SimpleAudioVolume') and session.SimpleAudioVolume:
                session.SimpleAudioVolume.SetMute(int(mute), None)
                return True
            
            # Try _ctl.QueryInterface
            if hasattr(session, '_ctl') and self.ISimpleAudioVolume:
                volume_interface = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                volume_interface.SetMute(int(mute), None)
                return True
        except Exception as e:
            if log_lookup:
                self.log_debug(f"Error setting app mute for PID {pid}: {e}", "WARNING")
        return False
    
    def get_system_volume(self):
        """Get system master volume (0.0 to 1.0)"""
        if not self.audio_available or not self.volume_interface:
            self.log_debug("System volume unavailable - no interface", "WARNING")
            return 1.0
        try:
            vol = self.volume_interface.GetMasterVolumeLevelScalar()
            return vol
        except Exception as e:
            self.log_debug(f"Error getting system volume: {e}", "ERROR")
            # Try to reinitialize
            self.log_debug("Attempting to reinitialize audio...", "INFO")
            self.init_audio()
            if self.volume_interface:
                try:
                    return self.volume_interface.GetMasterVolumeLevelScalar()
                except:
                    pass
        return 1.0
    
    def set_system_volume(self, level):
        """Set system master volume (0.0 to 1.0)"""
        if not self.audio_available or not self.volume_interface:
            self.log_debug("Cannot set system volume - no interface", "WARNING")
            return False
        try:
            level = max(0.0, min(1.0, float(level)))
            self.volume_interface.SetMasterVolumeLevelScalar(level, None)
            return True
        except Exception as e:
            self.log_debug(f"Error setting system volume: {e}", "ERROR")
            # Try to reinitialize
            self.init_audio()
            if self.volume_interface:
                try:
                    self.volume_interface.SetMasterVolumeLevelScalar(level, None)
                    return True
                except:
                    pass
        return False
    
    def get_system_mute(self):
        """Get system mute state"""
        if not self.audio_available or not self.volume_interface:
            return False
        try:
            return bool(self.volume_interface.GetMute())
        except Exception as e:
            self.log_debug(f"Error getting system mute: {e}", "ERROR")
        return False
    
    def set_system_mute(self, mute):
        """Set system mute state"""
        if not self.audio_available or not self.volume_interface:
            return False
        try:
            self.volume_interface.SetMute(int(mute), None)
            return True
        except Exception as e:
            self.log_debug(f"Error setting system mute: {e}", "ERROR")
        return False
    
    def toggle_pin(self):
        """Toggle always on top"""
        is_pinned = self.pin_to_top.get()
        self.root.attributes('-topmost', is_pinned)
        status = "pinned" if is_pinned else "unpinned"
        self.status_var.set(f"Window {status}")
    
    def toggle_debug(self):
        """Toggle debug mode and export log"""
        if self.debug_mode.get():
            self.log_debug("Debug mode enabled")
            self.status_var.set("Debug mode ON - actions will be logged")
        else:
            # Export log when turning off debug mode
            self.log_debug("Debug mode disabled - exporting log")
            self.export_debug_log()
    
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
            mon['short_name'] = f"Display {i + 1}{primary_tag}"
            
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
    
    def get_window_monitor_index(self, hwnd):
        """Get the index of the monitor a window is on"""
        try:
            rect = win32gui.GetWindowRect(hwnd)
            center_x = (rect[0] + rect[2]) // 2
            center_y = (rect[1] + rect[3]) // 2
            
            for i, mon in enumerate(self.monitors):
                ma = mon['monitor_area']
                if ma[0] <= center_x < ma[2] and ma[1] <= center_y < ma[3]:
                    return i
        except:
            pass
        return 0
        
    def setup_ui(self):
        """Setup the user interface"""
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, text="Window Manager", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
        # Debug checkbox
        debug_cb = ttk.Checkbutton(header_frame, text="🐛", 
                                    variable=self.debug_mode,
                                    command=self.toggle_debug)
        debug_cb.pack(side=tk.RIGHT, padx=(5, 0))
        
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
        
        # Audio status indicator
        audio_status = "🔊" if self.audio_available else "🔇❌"
        ttk.Label(audio_frame, text=f"Audio {audio_status}:", style='Muted.TLabel').pack(side=tk.LEFT)
        
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
        self.status_var = tk.StringVar(value="Ready" + (" (Audio OK)" if self.audio_available else " (No Audio)"))
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
        self.log_debug(f"Selected all {len(self.window_checkboxes)} windows")
    
    def deselect_all(self):
        """Deselect all windows"""
        for hwnd, var in self.window_checkboxes.items():
            var.set(False)
            self.update_card_style(hwnd, False)
        self.selected_windows.clear()
        self.update_selection_label()
        self.status_var.set("Cleared selection")
        self.log_debug("Deselected all windows")
    
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
        self.log_debug(f"Selected {count} windows on {target_monitor}")
        
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
        self.log_debug("Refreshing window list...")
        
        # Refresh monitors too
        self.monitors = self.get_monitors()
        
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.window_checkboxes.clear()
        self.window_cards.clear()
        self.windows_list.clear()
        self.window_pids.clear()
        
        old_selection = list(self.selected_windows.keys())
        self.selected_windows.clear()
        
        # Get current audio sessions for display - log this one
        audio_sessions = self.get_fresh_audio_sessions(log_sessions=True)
        audio_pids = set()
        
        # Extract PIDs from sessions (handle both regular sessions and dict-wrapped ones)
        for pid, session in audio_sessions.items():
            audio_pids.add(pid)
        
        session_names = []
        for pid in audio_pids:
            try:
                proc = psutil.Process(pid)
                session_names.append(f"{proc.name()}(PID:{pid})")
            except:
                session_names.append(f"PID:{pid}")
        if session_names:
            self.log_debug(f"Audio sessions: {', '.join(session_names)}")
        
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
        
        # Sort windows: audio apps first, then by process name
        def sort_key(item):
            hwnd, title, process, pid = item
            has_audio = pid in audio_pids
            # Also check if any child process has audio
            if not has_audio and pid:
                try:
                    proc = psutil.Process(pid)
                    for child in proc.children(recursive=True):
                        if child.pid in audio_pids:
                            has_audio = True
                            break
                except:
                    pass
            # Sort by: has_audio (True first = 0), then process name
            return (0 if has_audio else 1, process.lower())
        
        windows.sort(key=sort_key)
        
        for i, (hwnd, title, process, pid) in enumerate(windows):
            self.window_pids[hwnd] = pid
            has_audio = pid in audio_pids
            # Also check children for audio
            if not has_audio and pid:
                try:
                    proc = psutil.Process(pid)
                    for child in proc.children(recursive=True):
                        if child.pid in audio_pids:
                            has_audio = True
                            break
                except:
                    pass
            self.create_window_card(hwnd, title, process, pid, hwnd in old_selection, has_audio)
                
        self.update_selection_label()
        self.status_var.set(f"{len(windows)} windows")
        self.log_debug(f"Found {len(windows)} windows")
    
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
    
    def show_monitor_menu(self, event, hwnd):
        """Show dropdown menu for quick monitor switching"""
        menu = tk.Menu(self.root, tearoff=0, bg=self.colors['card'], fg=self.colors['fg'],
                       activebackground=self.colors['accent'], activeforeground='white',
                       font=('Segoe UI', 9))
        
        current_monitor_idx = self.get_window_monitor_index(hwnd)
        
        for i, mon in enumerate(self.monitors):
            if i == current_monitor_idx:
                # Current monitor - greyed out with star
                menu.add_command(label=f"★ {mon['short_name']}", state='disabled')
            else:
                # Other monitors - clickable
                menu.add_command(label=f"  {mon['short_name']}", 
                               command=lambda h=hwnd, m=mon: self.quick_move_to_monitor(h, m))
        
        # Show menu at button position
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()
    
    def quick_move_to_monitor(self, hwnd, target_monitor):
        """Quickly move a single window to a specific monitor"""
        self.ensure_topmost_during_action()
        self.log_debug(f"Quick moving window {hwnd} to {target_monitor['short_name']}")
        
        work_area = target_monitor['work_area']
        mon_x, mon_y = work_area[0], work_area[1]
        mon_width = work_area[2] - work_area[0]
        mon_height = work_area[3] - work_area[1]
        
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
            
            self.status_var.set(f"Moved to {target_monitor['short_name']}")
            self.log_debug(f"Moved window to {target_monitor['short_name']}")
            
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            self.log_debug(f"Error quick moving window {hwnd}: {e}", "ERROR")
        
    def create_window_card(self, hwnd, title, process, pid, was_selected=False, has_audio=False):
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
        
        # Audio indicator
        audio_indicator = " 🔊" if has_audio else ""
        
        proc_label = tk.Label(info, text=process + audio_indicator, bg=base_bg,
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
        
        # Monitor swap button (◀ with dropdown)
        monitor_btn = tk.Button(actions, text="◀", **btn_style)
        monitor_btn.configure(command=lambda e=None, h=hwnd, b=monitor_btn: self.show_monitor_menu_btn(h, b))
        monitor_btn.pack(side=tk.LEFT)
        
        # Focus button
        focus_btn = tk.Button(actions, text="◉", command=lambda h=hwnd: self.focus_window(h), **btn_style)
        focus_btn.pack(side=tk.LEFT)
        
        # Min/Max toggle button
        minmax_btn = tk.Button(actions, text="□", command=lambda h=hwnd: self.toggle_minmax(h), **btn_style)
        minmax_btn.pack(side=tk.LEFT)
        
        # Audio mute toggle button with right-click-hold for slider
        is_muted = self.get_app_mute(pid, log_lookup=False) if pid and has_audio else False
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
            'buttons': [monitor_btn, focus_btn, minmax_btn, audio_btn],
            'audio_btn': audio_btn,
            'monitor_btn': monitor_btn,
            'base_bg': base_bg,
            'pid': pid,
            'has_audio': has_audio
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
    
    def show_monitor_menu_btn(self, hwnd, btn):
        """Show monitor menu positioned relative to button"""
        # Create a fake event with the button's position
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height()
        
        menu = tk.Menu(self.root, tearoff=0, bg=self.colors['card'], fg=self.colors['fg'],
                       activebackground=self.colors['accent'], activeforeground='white',
                       font=('Segoe UI', 9))
        
        current_monitor_idx = self.get_window_monitor_index(hwnd)
        
        for i, mon in enumerate(self.monitors):
            if i == current_monitor_idx:
                # Current monitor - greyed out with star
                menu.add_command(label=f"★ {mon['short_name']}", state='disabled')
            else:
                # Other monitors - clickable
                menu.add_command(label=f"  {mon['short_name']}", 
                               command=lambda h=hwnd, m=mon: self.quick_move_to_monitor(h, m))
        
        try:
            menu.tk_popup(x, y)
        finally:
            menu.grab_release()
    
    def toggle_minmax(self, hwnd):
        """Toggle between minimized and maximized states"""
        self.ensure_topmost_during_action()
        try:
            if self.is_window_minimized(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                self.status_var.set("Window maximized")
                self.log_debug(f"Maximized window {hwnd}")
            else:
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                self.status_var.set("Window minimized")
                self.log_debug(f"Minimized window {hwnd}")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            self.log_debug(f"Error toggling minmax for {hwnd}: {e}", "ERROR")
    
    def toggle_app_mute(self, hwnd, pid, btn):
        """Toggle mute state for an app"""
        self.log_debug(f"Toggle mute for PID {pid}")
        
        if not pid:
            self.status_var.set("No PID for this window")
            return
            
        current_mute = self.get_app_mute(pid)
        if current_mute is None:
            self.status_var.set("No audio session (app may not be playing audio)")
            self.log_debug(f"No audio session for PID {pid}", "WARNING")
            return
            
        new_mute = not current_mute
        if self.set_app_mute(pid, new_mute):
            btn.configure(text="🔇" if new_mute else "🔊")
            self.status_var.set("Muted" if new_mute else "Unmuted")
            self.log_debug(f"Set mute={new_mute} for PID {pid}")
        else:
            self.status_var.set("Failed to toggle mute")
            self.log_debug(f"Failed to set mute for PID {pid}", "ERROR")
    
    def on_app_volume_press(self, event, hwnd, pid, btn):
        """Handle right-click press on app audio button - show slider"""
        if not pid:
            self.status_var.set("No audio session for this app")
            return
            
        current_volume = self.get_app_volume(pid)
        if current_volume is None:
            self.status_var.set("No audio session (app may not be playing audio)")
            return
        
        self.log_debug(f"Opening volume slider for PID {pid}, current volume: {current_volume:.0%}")
        
        def on_change(v):
            self.set_app_volume(pid, v, log_lookup=False)
        
        def on_close(final_vol):
            self.log_debug(f"Volume slider closed for PID {pid}, final volume: {final_vol:.0%}")
            self.update_audio_btn(hwnd, pid, btn)
        
        self.show_volume_slider(event, current_volume, on_change, on_close, "App Volume")
    
    def on_bulk_volume_press(self, event):
        """Handle right-click press on bulk volume button"""
        selected = self.get_selected_windows()
        
        if not selected:
            # No selection - control system volume
            current_volume = self.get_system_volume()
            self.log_debug(f"Opening system volume slider, current: {current_volume:.0%}")
            
            def on_change(v):
                self.set_system_volume(v)
            
            def on_close(final_vol):
                self.log_debug(f"System volume slider closed, final volume: {final_vol:.0%}")
            
            self.show_volume_slider(event, current_volume, on_change, on_close, "System")
        else:
            # Control selected windows
            self.log_debug(f"Opening bulk volume slider for {len(selected)} apps")
            
            def on_change(v):
                self.set_selected_volumes(v)
            
            def on_close(final_vol):
                self.log_debug(f"Bulk volume slider closed for {len(selected)} apps, final volume: {final_vol:.0%}")
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
        x = event.x_root - 30
        y = event.y_root - 180  # Position above the button
        
        # Slider height is approximately twice a card height
        slider_height = 170
        slider_width = 60
        
        slider_win.geometry(f"{slider_width}x{slider_height}+{x}+{y}")
        slider_win.configure(bg=self.colors['card'])
        
        # Store reference and callbacks
        self.volume_slider_window = slider_win
        self.volume_slider_on_close = on_close
        self.volume_slider_on_change = on_change
        
        # Store initial mouse position and volume for relative movement
        self.slider_start_y = event.y_root
        self.slider_start_volume = initial_volume
        self.current_slider_volume = initial_volume
        self.slider_is_dragging = True  # Mark that we're dragging
        
        # Title label
        title_lbl = tk.Label(slider_win, text=title, bg=self.colors['card'],
                             fg=self.colors['muted'], font=('Segoe UI', 7))
        title_lbl.pack(pady=(8, 0))
        
        # Volume percentage label
        self.vol_var = tk.StringVar(value=f"{int(initial_volume * 100)}%")
        vol_label = tk.Label(slider_win, textvariable=self.vol_var, bg=self.colors['card'],
                             fg=self.colors['fg'], font=('Segoe UI', 11, 'bold'))
        vol_label.pack(pady=(2, 8))
        
        # Create canvas for custom slider visualization
        canvas_height = slider_height - 70
        self.slider_canvas = tk.Canvas(slider_win, width=40, height=canvas_height,
                                        bg=self.colors['slider_bg'], highlightthickness=0)
        self.slider_canvas.pack(pady=(0, 8))
        
        # Slider visual parameters
        self.track_x = 20
        self.track_top = 5
        self.track_bottom = canvas_height - 5
        self.track_height = self.track_bottom - self.track_top
        
        # Track background
        self.slider_canvas.create_rectangle(self.track_x - 4, self.track_top, 
                                             self.track_x + 4, self.track_bottom,
                                             fill=self.colors['border'], 
                                             outline='')
        
        # Calculate initial handle position based on actual volume
        handle_y = self.track_bottom - (initial_volume * self.track_height)
        
        # Filled portion (from handle to bottom)
        self.fill_rect = self.slider_canvas.create_rectangle(
            self.track_x - 4, handle_y, self.track_x + 4, self.track_bottom,
            fill=self.colors['slider_fg'], outline='')
        
        # Handle
        self.handle = self.slider_canvas.create_oval(
            self.track_x - 10, handle_y - 10, self.track_x + 10, handle_y + 10,
            fill=self.colors['accent'], outline=self.colors['fg'], width=2)
        
        # Sensitivity: pixels of mouse movement for full volume range
        self.slider_sensitivity = 150  # 150 pixels = 0% to 100%
        
        # Bind mouse motion globally for tracking during right-click hold
        self.root.bind('<Motion>', self.on_slider_motion)
        slider_win.bind('<Motion>', self.on_slider_motion)
    
    def on_slider_motion(self, event):
        """Handle mouse motion while slider is open - relative movement"""
        if not self.volume_slider_window or not self.slider_canvas:
            return
        
        if self.slider_start_y is None or self.slider_start_volume is None:
            return
        
        if not self.slider_is_dragging:
            return
        
        try:
            # Calculate relative movement from start position
            # Moving UP (negative delta) = increase volume
            # Moving DOWN (positive delta) = decrease volume
            delta_y = self.slider_start_y - event.y_root
            
            # Convert pixel movement to volume change
            volume_change = delta_y / self.slider_sensitivity
            
            # Calculate new volume
            new_volume = self.slider_start_volume + volume_change
            new_volume = max(0.0, min(1.0, new_volume))
            
            self.current_slider_volume = new_volume
            
            # Update visual - calculate handle position
            handle_y = self.track_bottom - (new_volume * self.track_height)
            
            # Update handle position
            self.slider_canvas.coords(self.handle, 
                                       self.track_x - 10, handle_y - 10, 
                                       self.track_x + 10, handle_y + 10)
            
            # Update fill rectangle
            self.slider_canvas.coords(self.fill_rect, 
                                       self.track_x - 4, handle_y, 
                                       self.track_x + 4, self.track_bottom)
            
            # Update label
            self.vol_var.set(f"{int(new_volume * 100)}%")
            
            # Apply volume change (no logging during drag)
            if self.volume_slider_on_change:
                self.volume_slider_on_change(new_volume)
                
        except Exception as e:
            # Don't log during slider motion to avoid spam
            pass
    
    def close_volume_slider(self):
        """Close the volume slider"""
        # Mark that we're no longer dragging
        self.slider_is_dragging = False
        
        # Unbind motion event
        try:
            self.root.unbind('<Motion>')
        except:
            pass
        
        # Get final volume for logging
        final_volume = self.current_slider_volume if self.current_slider_volume is not None else 1.0
        
        # Call on_close callback if set (with final volume)
        if self.volume_slider_on_close:
            try:
                self.volume_slider_on_close(final_volume)
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
        self.slider_start_y = None
        self.slider_start_volume = None
        self.current_slider_volume = None
    
    def update_audio_btn(self, hwnd, pid, btn):
        """Update audio button icon based on mute state"""
        if hwnd in self.window_cards:
            is_muted = self.get_app_mute(pid, log_lookup=False)
            if is_muted is not None:
                btn.configure(text="🔇" if is_muted else "🔊")
    
    def update_all_audio_btns(self):
        """Update all audio button icons"""
        for hwnd, card_data in self.window_cards.items():
            pid = card_data.get('pid')
            if pid:
                is_muted = self.get_app_mute(pid, log_lookup=False)
                if is_muted is not None:
                    card_data['audio_btn'].configure(text="🔇" if is_muted else "🔊")
    
    def set_selected_volumes(self, volume):
        """Set volume for all selected windows"""
        selected = self.get_selected_windows()
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid:
                self.set_app_volume(pid, volume, log_lookup=False)
    
    def bulk_mute(self):
        """Mute selected windows or system"""
        selected = self.get_selected_windows()
        self.log_debug(f"Bulk mute, {len(selected)} selected")
        
        if not selected:
            if self.set_system_mute(True):
                self.status_var.set("System muted")
                self.log_debug("System muted")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid and self.set_app_mute(pid, True):
                count += 1
                if hwnd in self.window_cards:
                    self.window_cards[hwnd]['audio_btn'].configure(text="🔇")
        
        self.status_var.set(f"Muted {count} window(s)")
        self.log_debug(f"Muted {count} windows")
    
    def bulk_unmute(self):
        """Unmute selected windows or system"""
        selected = self.get_selected_windows()
        self.log_debug(f"Bulk unmute, {len(selected)} selected")
        
        if not selected:
            if self.set_system_mute(False):
                self.status_var.set("System unmuted")
                self.log_debug("System unmuted")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid and self.set_app_mute(pid, False):
                count += 1
                if hwnd in self.window_cards:
                    self.window_cards[hwnd]['audio_btn'].configure(text="🔊")
        
        self.status_var.set(f"Unmuted {count} window(s)")
        self.log_debug(f"Unmuted {count} windows")
        
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
        self.log_debug(f"Focusing window {hwnd}")
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            self.log_debug(f"Error focusing window {hwnd}: {e}", "ERROR")
            
    def move_to_monitor(self):
        """Move selected windows to chosen monitor"""
        self.ensure_topmost_during_action()
        selected = self.get_selected_windows()
        
        if not selected:
            self.status_var.set("No windows selected")
            return
        
        self.log_debug(f"Moving {len(selected)} windows to monitor")
            
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
                self.log_debug(f"Error moving window {hwnd}: {e}", "ERROR")
                
        self.status_var.set(f"Moved {moved_count} window(s)")
        self.log_debug(f"Moved {moved_count} windows to {monitor_name}")
    
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
            self.log_debug("Split windows horizontally")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            self.log_debug(f"Error splitting windows: {e}", "ERROR")
    
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
            self.log_debug("Split windows vertically")
        except Exception as e:
            self.status_var.set(f"Error: {e}")
            self.log_debug(f"Error splitting windows: {e}", "ERROR")
    
    def run(self):
        """Run the application"""
        self.log_debug("Application started")
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