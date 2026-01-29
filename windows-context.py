import sys
if sys.platform == 'win32':
    import ctypes
import tkinter as tk

from tkinter import ttk
import ctypes
from ctypes import wintypes
import win32gui
import win32con
import win32process
import win32api
import psutil
import os
import time
import traceback
import atexit
import json
import warnings
from collections import OrderedDict

# Suppress pycaw COM warnings
warnings.filterwarnings('ignore', category=UserWarning, module='pycaw')

# Windows API constants
DWMWA_CLOAKED = 14
GWL_EXSTYLE = -20
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_APPWINDOW = 0x00040000

SETTINGS_FILE = "window_manager_settings.json"


class WindowManager:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Window Manager")
        
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
        
        # Load settings
        self.settings = self.load_settings()

        # Pinned windows tracking
        self.pinned_windows = set()  # Current hwnds that are pinned
        self.pinned_identifiers = self.settings.get('pinned_identifiers', set())
        if not isinstance(self.pinned_identifiers, set):
            self.pinned_identifiers = set()
        
        # Debugging
        self.debug_mode = tk.BooleanVar(value=self.settings.get('debug_enabled', False))
        self.verbose_logging = self.settings.get('verbose_logging', False)
        self.debug_log = []
        self.debug_dir = "windows-manager-debugging"
        
        self.pin_to_top = tk.BooleanVar(value=False)
        self.root.attributes('-topmost', False)
        
        # Track selected windows in order
        self.selected_windows = OrderedDict()
        self.window_checkboxes = {}
        self.window_cards = {}
        self.windows_list = []
        self.window_pids = {}
        # Pinned windows (by hwnd, but we'll also track by title+process for persistence)
        self.pinned_windows = set()  # Current hwnds that are pinned
        self.pinned_identifiers = set()  # (process, title) tuples for persistence across refresh
        
        # Audio state tracking
        self.muted_pids = set()
        self.volume_levels = {}
        
        # Volume slider popup
        self.volume_slider_window = None
        self.slider_start_y = None
        self.slider_start_volume = None
        self.current_slider_volume = None
        self.volume_slider_on_change = None
        self.volume_slider_on_close = None
        self.slider_canvas = None
        self.slider_is_dragging = False
        
        # Get monitors
        self.monitors = self.get_monitors()
        
        # Initialize audio
        self.init_audio()
        
        self.setup_styles()
        self.setup_ui()
        self.refresh_windows()
        
        # Start background audio monitor
        self.audio_monitor_running = True
        self.start_audio_monitor()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        atexit.register(self.export_debug_log_on_exit)
    
    def load_settings(self):
        """Load settings from file"""
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
                    # Convert pinned_identifiers from list of lists to set of tuples
                    if 'pinned_identifiers' in settings:
                        settings['pinned_identifiers'] = set(tuple(x) for x in settings['pinned_identifiers'])
                    else:
                        settings['pinned_identifiers'] = set()
                    return settings
        except Exception as e:
            print(f"Error loading settings: {e}")
        return {'debug_enabled': False, 'verbose_logging': False, 'pinned_identifiers': set()}
    
    def save_settings(self):
        """Save settings to file"""
        try:
            settings = {
                'debug_enabled': self.debug_mode.get(),
                'verbose_logging': self.verbose_logging,
                'pinned_identifiers': list(list(x) for x in self.pinned_identifiers)
            }
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            print(f"Error saving settings: {e}")
    
    def on_closing(self):
        """Handle window close"""
        self.audio_monitor_running = False
        self.save_settings()
        if self.debug_log:
            self.log_debug("Application closing")
            self.export_debug_log()
        self.root.destroy()
    
    def export_debug_log_on_exit(self):
        pass
    
    def log_debug(self, message, level="INFO"):
        """Add a message to the debug log"""
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        entry = f"[{timestamp}] [{level}] {message}"
        self.debug_log.append(entry)
        
        if len(self.debug_log) > 1000:
            self.debug_log = self.debug_log[-500:]
        
        if self.debug_mode.get():
            print(entry)
    
    def log_verbose(self, message, level="INFO"):
        """Log verbose messages only if verbose logging is enabled"""
        if self.verbose_logging:
            self.log_debug(message, level)
    
    def export_debug_log(self):
        """Export debug log to a file"""
        if not self.debug_log:
            return None
            
        if not os.path.exists(self.debug_dir):
            os.makedirs(self.debug_dir)
        
        unix_time = int(time.time())
        filename = os.path.join(self.debug_dir, f"debug_{unix_time}.log")
        
        system_info = self.gather_system_info()
        
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
        
        import sys
        info["Python Version"] = sys.version
        
        import platform
        info["OS"] = platform.system()
        info["OS Version"] = platform.version()
        info["OS Release"] = platform.release()
        info["Machine"] = platform.machine()
        
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
        
        info["Audio Available"] = self.audio_available
        info["Multi-Device Support"] = getattr(self, 'multi_device_support', False)
        info["Verbose Logging"] = self.verbose_logging
        
        # Count active render devices
        try:
            devices = self._get_active_render_devices()
            info["Active Render Devices"] = len(devices)
            for i, dev in enumerate(devices):
                info[f"  Device {i+1}"] = dev.get('name', 'Unknown')
        except:
            info["Active Render Devices"] = "Error"
        
        info["Muted PIDs"] = str(list(self.muted_pids))
        info["Volume Levels"] = str(dict(self.volume_levels))
        
        info["Monitor Count"] = len(self.monitors)
        for i, mon in enumerate(self.monitors):
            info[f"Monitor {i+1}"] = f"{mon['name']} - {mon['resolution']}"
        
        info["Tracked Windows"] = len(self.windows_list)
        
        return info
    def start_audio_monitor(self):
        """Start background audio session monitoring"""
        def monitor():
            if not self.audio_monitor_running:
                return
            
#            self.enforce_audio_states()
            
            self.root.after(500, monitor)
        
        self.root.after(1000, monitor)
    
    def init_audio(self):      
        """Initialize audio control via pycaw"""
        self.audio_available = False
        self.volume_interface = None
        self.AudioUtilities = None
        self.multi_device_support = False
        
        self.log_debug("Initializing audio...")
        
        try:
            from pycaw.pycaw import (
                AudioUtilities, 
                ISimpleAudioVolume, 
                IAudioEndpointVolume,
                IMMDeviceEnumerator,
                IAudioSessionManager2,
                IAudioSessionEnumerator,
                IAudioSessionControl2,
                IMMDeviceCollection,
            )
            from comtypes import CLSCTX_ALL, GUID, CoCreateInstance, COMObject
            from ctypes import POINTER, cast
            import comtypes
            
            self.AudioUtilities = AudioUtilities
            self.ISimpleAudioVolume = ISimpleAudioVolume
            self.IAudioEndpointVolume = IAudioEndpointVolume
            self.IMMDeviceEnumerator = IMMDeviceEnumerator
            self.IAudioSessionManager2 = IAudioSessionManager2
            self.IAudioSessionEnumerator = IAudioSessionEnumerator
            self.IAudioSessionControl2 = IAudioSessionControl2
            self.IMMDeviceCollection = IMMDeviceCollection
            self.CLSCTX_ALL = CLSCTX_ALL
            self.comtypes = comtypes
            self.cast = cast
            self.POINTER = POINTER
            self.CoCreateInstance = CoCreateInstance
            self.GUID = GUID
            
            # Define the MMDeviceEnumerator CLSID (Windows constant)
            self.CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')
            
            # Constants for device enumeration
            self.eRender = 0  # Output devices
            self.eCapture = 1  # Input devices  
            self.eAll = 2
            self.DEVICE_STATE_ACTIVE = 0x1
            
            self.multi_device_support = True
            self.log_debug("Imported pycaw successfully with multi-device support")
            
            # Get system volume control from default device
            try:
                speakers = self.AudioUtilities.GetSpeakers()
                if speakers:
                    # AudioDevice wrapper - get the underlying IMMDevice
                    if hasattr(speakers, '_dev') and speakers._dev:
                        dev = speakers._dev
                        interface = dev.Activate(
                            self.IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                        self.volume_interface = interface.QueryInterface(self.IAudioEndpointVolume)
                        self.log_debug("System volume interface ready")
                    else:
                        self.log_debug(f"GetSpeakers returned: {type(speakers)}, attrs: {[a for a in dir(speakers) if not a.startswith('_')]}", "WARNING")
            except Exception as e:
                self.log_debug(f"Could not get system volume interface: {e}", "WARNING")
                self.log_debug(traceback.format_exc(), "WARNING")
            
            # Test multi-device enumeration
            try:
                devices = self._enumerate_all_audio_devices()
                self.log_debug(f"Found {len(devices)} audio render device(s):")
                for dev in devices:
                    self.log_debug(f"  - {dev.get('name', 'Unknown')}")
                    
                # Get sessions from all devices
                all_sessions = self._get_all_audio_sessions_all_devices()
                total = sum(len(v) for v in all_sessions.values())
                self.log_debug(f"Found {total} audio session(s) across all devices:")
                for pid, sessions in all_sessions.items():
                    for sess in sessions:
                        self.log_debug(f"  - {sess.get('name', '?')} (PID: {pid}) on {sess.get('device_name', '?')}")
                        
            except Exception as e:
                self.log_debug(f"Multi-device enumeration failed: {e}", "WARNING")
                self.log_debug(traceback.format_exc(), "WARNING")
                self.multi_device_support = False
            
            self.audio_available = True
            self.log_debug("Audio initialized successfully")
                
        except ImportError as e:
            self.log_debug(f"pycaw import error: {e}", "ERROR")
            self.log_debug(traceback.format_exc(), "ERROR")
        except Exception as e:
            self.log_debug(f"Error initializing audio: {e}", "ERROR")
            self.log_debug(traceback.format_exc(), "ERROR")


    def _enumerate_all_audio_devices(self):
        """Enumerate all active audio render devices"""
        devices = []
        
        if not self.audio_available or not self.multi_device_support:
            # Fallback
            try:
                speakers = self.AudioUtilities.GetSpeakers()
                if speakers and hasattr(speakers, '_dev'):
                    devices.append({
                        'device': speakers._dev,
                        'name': 'Default Output',
                        'index': 0
                    })
            except:
                pass
            return devices
        
        try:
            # Create the device enumerator via COM
            enumerator = self.CoCreateInstance(
                self.CLSID_MMDeviceEnumerator,
                self.IMMDeviceEnumerator,
                self.CLSCTX_ALL
            )
            
            # Get all active render (output) devices
            collection = enumerator.EnumAudioEndpoints(self.eRender, self.DEVICE_STATE_ACTIVE)
            count = collection.GetCount()
            
            self.log_verbose(f"EnumAudioEndpoints found {count} active render device(s)")
            
            for i in range(count):
                try:
                    device = collection.Item(i)
                    
                    # Get friendly name
                    device_name = f"Audio Device {i+1}"
                    try:
                        # Open property store in read mode
                        props = device.OpenPropertyStore(0)  # STGM_READ = 0
                        
                        # PKEY_Device_FriendlyName = {a45c254e-df1c-4efd-8020-67d146a850e0}, 14
                        from pycaw.pycaw import PROPERTYKEY
                        PKEY_FriendlyName = PROPERTYKEY()
                        PKEY_FriendlyName.fmtid = self.GUID('{a45c254e-df1c-4efd-8020-67d146a850e0}')
                        PKEY_FriendlyName.pid = 14
                        
                        value = props.GetValue(PKEY_FriendlyName)
                        
                        # Extract string from PROPVARIANT via union.pwszVal
                        if value and hasattr(value, 'union'):
                            union = value.union
                            if hasattr(union, 'pwszVal') and union.pwszVal:
                                device_name = union.pwszVal
                                
                    except Exception as e:
                        self.log_verbose(f"Could not get device {i} friendly name: {e}")
                    
                    devices.append({
                        'device': device,
                        'name': device_name,
                        'index': i
                    })
                    
                except Exception as e:
                    self.log_verbose(f"Error getting device {i}: {e}")
                    
        except Exception as e:
            self.log_debug(f"Error enumerating audio devices: {e}", "ERROR")
            self.log_debug(traceback.format_exc(), "ERROR")
            
            # Fallback to default device
            try:
                speakers = self.AudioUtilities.GetSpeakers()
                if speakers and hasattr(speakers, '_dev'):
                    devices.append({
                        'device': speakers._dev,
                        'name': 'Default Output',
                        'index': 0
                    })
            except:
                pass
        
        return devices


    def _get_sessions_from_device(self, device, device_name):
        """Get all audio sessions from a specific device"""
        sessions = []
        
        try:
            # Activate the session manager on this device
            mgr = device.Activate(
                self.IAudioSessionManager2._iid_,
                self.CLSCTX_ALL,
                None
            )
            
            session_mgr = mgr.QueryInterface(self.IAudioSessionManager2)
            session_enum = session_mgr.GetSessionEnumerator()
            
            count = session_enum.GetCount()
            self.log_verbose(f"  {device_name}: {count} session(s)")
            
            for i in range(count):
                try:
                    session_ctl = session_enum.GetSession(i)
                    
                    # Get extended session control for PID
                    session_ctl2 = session_ctl.QueryInterface(self.IAudioSessionControl2)
                    pid = session_ctl2.GetProcessId()
                    
                    if pid == 0:
                        # System sounds session, skip
                        continue
                    
                    # Get process name
                    proc_name = f"PID:{pid}"
                    try:
                        proc = psutil.Process(pid)
                        proc_name = proc.name()
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        pass
                    
                    # Get volume control interface
                    volume_ctl = session_ctl.QueryInterface(self.ISimpleAudioVolume)
                    
                    sessions.append({
                        'pid': pid,
                        'name': proc_name,
                        'device_name': device_name,
                        'volume': volume_ctl,
                        'session_ctl': session_ctl,
                        'session_ctl2': session_ctl2
                    })
                    
                    self.log_verbose(f"    - {proc_name} (PID: {pid})")
                    
                except Exception as e:
                    self.log_verbose(f"    Error getting session {i}: {e}")
                    
        except Exception as e:
            self.log_verbose(f"  Error getting sessions from {device_name}: {e}")
            self.log_verbose(traceback.format_exc())
        
        return sessions
    def _get_all_audio_sessions_all_devices(self):

        """Get all audio sessions from ALL audio render devices"""
        sessions_by_pid = {}
        
        if not self.audio_available:
            return sessions_by_pid
        
        try:
            # Get all devices
            devices = self._enumerate_all_audio_devices()
            
            for dev_info in devices:
                device = dev_info['device']
                device_name = dev_info['name']
                
                sessions = self._get_sessions_from_device(device, device_name)
                
                for sess in sessions:
                    pid = sess['pid']
                    if pid not in sessions_by_pid:
                        sessions_by_pid[pid] = []
                    sessions_by_pid[pid].append(sess)
                    
        except Exception as e:
            self.log_debug(f"Error in _get_all_audio_sessions_all_devices: {e}", "ERROR")
            self.log_debug(traceback.format_exc(), "ERROR")
            
            # Fallback to AudioUtilities.GetAllSessions()
            try:
                all_sessions = self.AudioUtilities.GetAllSessions()
                for session in all_sessions:
                    if session.Process:
                        pid = session.Process.pid
                        proc_name = session.Process.name()
                        
                        try:
                            volume_ctl = session._ctl.QueryInterface(self.ISimpleAudioVolume)
                            
                            if pid not in sessions_by_pid:
                                sessions_by_pid[pid] = []
                            
                            sessions_by_pid[pid].append({
                                'pid': pid,
                                'name': proc_name,
                                'device_name': 'Default Device',
                                'volume': volume_ctl,
                                'session': session
                            })
                        except:
                            pass
            except:
                pass
        
        return sessions_by_pid

    def _get_active_render_devices(self):
        """Get list of active render device info for display"""
        return self._enumerate_all_audio_devices()

    def get_audio_sessions_for_pid(self, pid, log_lookup=True):
        """Get audio sessions using fresh COM imports each time"""
        from pycaw.pycaw import ISimpleAudioVolume, IMMDeviceEnumerator, IAudioSessionManager2, IAudioSessionControl2, PROPERTYKEY
        from comtypes import CLSCTX_ALL, GUID, CoCreateInstance
        
        sessions = []
        
        if not self.audio_available:
            return sessions
        
        CLSID_MMDeviceEnumerator = GUID('{BCDE0395-E52F-467C-8E3D-C4579291692E}')
        eRender = 0
        DEVICE_STATE_ACTIVE = 0x1
        
        try:
            enumerator = CoCreateInstance(
                CLSID_MMDeviceEnumerator,
                IMMDeviceEnumerator,
                CLSCTX_ALL
            )
            
            collection = enumerator.EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE)
            count = collection.GetCount()
            
            for i in range(count):
                device = collection.Item(i)
                
                # Get device name
                device_name = f"Device {i}"
                try:
                    props = device.OpenPropertyStore(0)
                    PKEY_FriendlyName = PROPERTYKEY()
                    PKEY_FriendlyName.fmtid = GUID('{a45c254e-df1c-4efd-8020-67d146a850e0}')
                    PKEY_FriendlyName.pid = 14
                    value = props.GetValue(PKEY_FriendlyName)
                    if value and hasattr(value, 'union') and value.union.pwszVal:
                        device_name = value.union.pwszVal
                except:
                    pass
                
                # Get session manager for this device
                mgr = device.Activate(IAudioSessionManager2._iid_, CLSCTX_ALL, None)
                session_mgr = mgr.QueryInterface(IAudioSessionManager2)
                session_enum = session_mgr.GetSessionEnumerator()
                
                session_count = session_enum.GetCount()
                
                # DEBUG: Print all sessions on this device
#                print(f"DEBUG {device_name}: {session_count} sessions")
#                for j in range(session_count):
#                    try:
#                        session_ctl = session_enum.GetSession(j)
#                        session_ctl2 = session_ctl.QueryInterface(IAudioSessionControl2)
#                        sess_pid = session_ctl2.GetProcessId()
#                        print(f"  session[{j}] pid={sess_pid}")
#                    except Exception as e:
#                        print(f"  session[{j}] error: {e}")
#                # END DEBUG
                
                for j in range(session_count):
                    try:
                        session_ctl = session_enum.GetSession(j)
                        session_ctl2 = session_ctl.QueryInterface(IAudioSessionControl2)
                        
                        session_pid = session_ctl2.GetProcessId()
                        if session_pid != pid:
                            continue
                        
                        volume_ctl = session_ctl.QueryInterface(ISimpleAudioVolume)
                        
                        sessions.append({
                            'pid': pid,
                            'device_name': device_name,
                            'volume': volume_ctl
                        })
                    except:
                        continue
                        
        except Exception as e:
            self.log_debug(f"Error getting sessions for PID {pid}: {e}", "ERROR")
        
        if log_lookup:
            self.log_debug(f"Found {len(sessions)} session(s) for PID {pid}")
        
        return sessions


    def get_app_volume(self, pid, log_lookup=True):
        """Get volume level for an app"""
        sessions = self.get_audio_sessions_for_pid(pid, log_lookup=log_lookup)

        if sessions:
            volume_ctl = sessions[0].get('volume')
            if volume_ctl:
                try:
                    return volume_ctl.GetMasterVolume()
                except:
                    pass
        return None
    
    def set_app_volume(self, pid, level, log_lookup=True):
        """Set volume for an app on ALL its sessions"""
        level = max(0.0, min(1.0, float(level)))
        self.volume_levels[pid] = level
        
        sessions = self.get_audio_sessions_for_pid(pid, log_lookup=log_lookup)
        success = False
        
        for sess_info in sessions:
            volume_ctl = sess_info.get('volume')
            if volume_ctl:
                try:
                    volume_ctl.SetMasterVolume(level, None)
                    success = True
                    if log_lookup:
                        self.log_debug(f"Set volume={level:.0%} for PID {pid}")
                except Exception as e:
                    self.log_verbose(f"Error setting volume: {e}", "WARNING")
        
        return success
    
    def get_app_mute(self, pid, log_lookup=True):
        """Get mute state for an app"""
        if pid in self.muted_pids:
            return True
        
        sessions = self.get_audio_sessions_for_pid(pid, log_lookup=log_lookup)
        if sessions:
            volume_ctl = sessions[0].get('volume')
            if volume_ctl:
                try:
                    return bool(volume_ctl.GetMute())
                except:
                    pass
        return None
    
    def set_app_mute(self, pid, mute, log_lookup=True):
        """Set mute for an app on ALL its sessions"""
        if mute:
            self.muted_pids.add(pid)
            self.log_debug(f"Added PID {pid} to muted_pids: {self.muted_pids}")
        else:
            self.muted_pids.discard(pid)
            self.log_debug(f"Removed PID {pid} from muted_pids: {self.muted_pids}")
        
        sessions = self.get_audio_sessions_for_pid(pid, log_lookup=log_lookup)
        success = False
        muted_count = 0
        
        for sess_info in sessions:
            volume_ctl = sess_info.get('volume')
            device_name = sess_info.get('device_name', 'Unknown')
            if volume_ctl:
                try:
                    # Set the mute
                    volume_ctl.SetMute(int(mute), None)
                    
                    # Verify it actually took effect
                    actual_mute = volume_ctl.GetMute()
                    if actual_mute == int(mute):
                        success = True
                        muted_count += 1
                        self.log_debug(f"  ✓ {device_name}: mute={mute} (verified)")
                    else:
                        self.log_debug(f"  ✗ {device_name}: SetMute({mute}) failed, actual={actual_mute}", "WARNING")
                        
                except Exception as e:
                    self.log_debug(f"  ✗ {device_name}: Error - {e}", "WARNING")
        
        if log_lookup:
            if muted_count > 0:
                self.log_debug(f"Set mute={mute} for PID {pid} on {muted_count} session(s)")
            else:
                self.log_debug(f"No sessions found to {'mute' if mute else 'unmute'} for PID {pid}")
        
        return success
    def get_system_volume(self):
        """Get system master volume"""
        if self.volume_interface:
            try:
                return self.volume_interface.GetMasterVolumeLevelScalar()
            except:
                pass
        return 1.0
    
    def set_system_volume(self, level):
        """Set system master volume"""
        if self.volume_interface:
            try:
                level = max(0.0, min(1.0, float(level)))
                self.volume_interface.SetMasterVolumeLevelScalar(level, None)
                return True
            except:
                pass
        return False
    
    def get_system_mute(self):
        """Get system mute state"""
        if self.volume_interface:
            try:
                return bool(self.volume_interface.GetMute())
            except:
                pass
        return False
    
    def set_system_mute(self, mute):
        """Set system mute state"""
        if self.volume_interface:
            try:
                self.volume_interface.SetMute(int(mute), None)
                return True
            except:
                pass
        return False
    
    def show_audio_device_menu(self):
        """Show menu for audio device info and sessions"""
        menu_win = tk.Toplevel(self.root)
        menu_win.title("Audio Sessions")
        menu_win.geometry("450x400")
        menu_win.configure(bg=self.colors['bg'])
        menu_win.transient(self.root)
        menu_win.grab_set()
        
        menu_win.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - 450) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 400) // 2
        menu_win.geometry(f"+{x}+{y}")
        
        title_frame = tk.Frame(menu_win, bg=self.colors['bg'])
        title_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        tk.Label(title_frame, text="Active Audio Sessions",
                 bg=self.colors['bg'], fg=self.colors['fg'],
                 font=('Segoe UI', 12, 'bold')).pack(side=tk.LEFT)
        
        status_text = "✓ Audio Available" if self.audio_available else "✗ Audio Unavailable"
        status_color = self.colors['success'] if self.audio_available else self.colors['muted_icon']
        tk.Label(title_frame, text=status_text,
                 bg=self.colors['bg'], fg=status_color,
                 font=('Segoe UI', 9)).pack(side=tk.RIGHT)
        
        list_frame = tk.Frame(menu_win, bg=self.colors['bg'])
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        canvas = tk.Canvas(list_frame, bg=self.colors['bg'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=self.colors['bg'])
        
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def populate_list():
            for widget in scrollable.winfo_children():
                widget.destroy()
            
            if not self.audio_available:
                frame = tk.Frame(scrollable, bg=self.colors['card'])
                frame.pack(fill=tk.X, pady=2)
                tk.Label(frame, text="Audio control not available.\nCheck that pycaw is installed correctly.",
                        bg=self.colors['card'], fg=self.colors['muted'],
                        font=('Segoe UI', 9), justify='left').pack(padx=10, pady=8, anchor='w')
                return
            
            all_sessions = self._get_all_audio_sessions_all_devices()
            
            if not all_sessions:
                frame = tk.Frame(scrollable, bg=self.colors['card'])
                frame.pack(fill=tk.X, pady=2)
                tk.Label(frame, text="No active audio sessions found.\nStart playing audio in an application.",
                        bg=self.colors['card'], fg=self.colors['muted'],
                        font=('Segoe UI', 9), justify='left').pack(padx=10, pady=8, anchor='w')
                return
            
            for pid, sessions in all_sessions.items():
                if not sessions:
                    continue
                    
                sess = sessions[0]
                proc_name = sess.get('name', f'PID:{pid}')
                is_muted = pid in self.muted_pids
                
                frame = tk.Frame(scrollable, bg=self.colors['card'])
                frame.pack(fill=tk.X, pady=2)
                
                # Get current volume
                volume_str = "?"
                try:
                    vol = sess.get('volume')
                    if vol:
                        volume_str = f"{int(vol.GetMasterVolume() * 100)}%"
                except:
                    pass
                
                mute_icon = "🔇" if is_muted else "🔊"
                tk.Label(frame, text=f"{mute_icon} {proc_name}",
                        bg=self.colors['card'], fg=self.colors['fg'],
                        font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=10, pady=(8, 2))
                
                tk.Label(frame, text=f"  PID: {pid} | Volume: {volume_str} | Sessions: {len(sessions)}",
                        bg=self.colors['card'], fg=self.colors['muted'],
                        font=('Segoe UI', 8)).pack(anchor='w', padx=10, pady=(0, 8))
        
        populate_list()
        
        verbose_frame = tk.Frame(menu_win, bg=self.colors['bg'])
        verbose_frame.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        verbose_var = tk.BooleanVar(value=self.verbose_logging)
        verbose_cb = tk.Checkbutton(verbose_frame, text="Enable verbose logging",
                                    variable=verbose_var,
                                    bg=self.colors['bg'], fg=self.colors['muted'],
                                    selectcolor=self.colors['bg'],
                                    activebackground=self.colors['bg'],
                                    font=('Segoe UI', 9))
        verbose_cb.pack(side=tk.LEFT)
        
        btn_frame = tk.Frame(menu_win, bg=self.colors['bg'])
        btn_frame.pack(fill=tk.X, padx=15, pady=(0, 15))
        
        def refresh_list():
            populate_list()
        
        def save_and_close():
            self.verbose_logging = verbose_var.get()
            self.save_settings()
            menu_win.destroy()
        
        tk.Button(btn_frame, text="Refresh", command=refresh_list,
                 bg=self.colors['card'], fg=self.colors['fg'],
                 font=('Segoe UI', 9), bd=0, padx=15, pady=5).pack(side=tk.LEFT)
        
        tk.Button(btn_frame, text="Close", command=save_and_close,
                 bg=self.colors['accent'], fg='white',
                 font=('Segoe UI', 9, 'bold'), bd=0, padx=20, pady=5).pack(side=tk.RIGHT)
    
    def toggle_pin(self):
        """Toggle always on top"""
        is_pinned = self.pin_to_top.get()
        self.root.attributes('-topmost', is_pinned)
        status = "pinned" if is_pinned else "unpinned"
        self.status_var.set(f"Window {status}")
    
    def toggle_debug(self):
        """Toggle debug mode"""
        self.save_settings()
        if self.debug_mode.get():
            self.log_debug("Debug mode enabled")
            self.status_var.set("Debug mode ON")
        else:
            self.log_debug("Debug mode disabled - exporting log")
            self.export_debug_log()
    
    def ensure_topmost_during_action(self):
        """Temporarily ensure window is on top during actions"""
        # Cancel any pending restore job
        if hasattr(self, '_restore_job') and self._restore_job:
            self.root.after_cancel(self._restore_job)
            self._restore_job = None
    
        # Always bring to top during action
        self.root.attributes('-topmost', True)
        self.root.lift()
        self.root.update_idletasks()
    
        # After 1 second, restore based on checkbox state
        def restore_topmost():
            self._restore_job = None
            if not self.pin_to_top.get():
                self.root.attributes('-topmost', False)
    
        self._restore_job = self.root.after(1000, restore_topmost)
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
        """Configure ttk styles"""
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
        """Get all connected monitors"""
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
        
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, text="Window Manager", style='Title.TLabel')
        title_label.pack(side=tk.LEFT)
        
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
        
        split_frame = ttk.Frame(main_frame)
        split_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(split_frame, text="◧ Split H", width=10,
                   command=self.split_vertical).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(split_frame, text="⬒ Split V", width=10,
                   command=self.split_horizontal).pack(side=tk.LEFT, padx=(0, 5))
        
        self.selection_label = ttk.Label(split_frame, text="0 selected", style='Muted.TLabel')
        self.selection_label.pack(side=tk.RIGHT)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
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
        
        audio_frame = ttk.Frame(main_frame)
        audio_frame.pack(fill=tk.X, pady=(0, 10))
        
        audio_status = "🔊" if self.audio_available else "🔇❌"
        ttk.Label(audio_frame, text=f"Audio {audio_status}:", style='Muted.TLabel').pack(side=tk.LEFT)
        
        self.audio_device_btn = tk.Button(audio_frame, text="🎧 Sessions",
                                          bg=self.colors['card'], fg=self.colors['fg'],
                                          font=('Segoe UI', 9), bd=0, padx=8, pady=4,
                                          activebackground=self.colors['card_hover'],
                                          command=self.show_audio_device_menu)
        self.audio_device_btn.pack(side=tk.LEFT, padx=(4, 4))
        
        self.bulk_mute_btn = tk.Button(audio_frame, text="🔇 Mute", 
                                        bg=self.colors['card'], fg=self.colors['fg'],
                                        font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                        activebackground=self.colors['card_hover'],
                                        command=self.bulk_mute)
        self.bulk_mute_btn.pack(side=tk.LEFT, padx=(0, 4))
        
        self.bulk_unmute_btn = tk.Button(audio_frame, text="🔊 Unmute",
                                          bg=self.colors['card'], fg=self.colors['fg'],
                                          font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                          activebackground=self.colors['card_hover'],
                                          command=self.bulk_unmute)
        self.bulk_unmute_btn.pack(side=tk.LEFT, padx=(0, 4))
        
        self.bulk_volume_btn = tk.Button(audio_frame, text="🎚️ Vol",
                                          bg=self.colors['card'], fg=self.colors['fg'],
                                          font=('Segoe UI', 9), bd=0, padx=10, pady=4,
                                          activebackground=self.colors['card_hover'])
        self.bulk_volume_btn.pack(side=tk.LEFT)
        self.bulk_volume_btn.bind('<ButtonPress-3>', self.on_bulk_volume_press)
        self.bulk_volume_btn.bind('<ButtonRelease-3>', self.on_volume_release)
        
        ttk.Separator(main_frame, orient='horizontal').pack(fill=tk.X, pady=10)
        
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
        
        self.status_var = tk.StringVar(value="Ready" + (" (Audio OK)" if self.audio_available else " (No Audio)"))
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, style='Muted.TLabel')
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
        """Select all windows on selected monitor"""
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
        
        self.monitors = self.get_monitors()
        
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
            
        self.window_checkboxes.clear()
        self.window_cards.clear()
        self.windows_list.clear()
        self.window_pids.clear()
        
        old_selection = list(self.selected_windows.keys())
        self.selected_windows.clear()
        
        all_audio_sessions = self._get_all_audio_sessions_all_devices()
        audio_pids = set(all_audio_sessions.keys())
        
        if all_audio_sessions:
            session_info = []
            for pid, sess_list in all_audio_sessions.items():
                name = sess_list[0].get('name', f'PID:{pid}')
                session_info.append(f"{name}(PID:{pid})")
            self.log_debug(f"Audio sessions: {', '.join(session_info)}")
        
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
        
        def sort_key(item):
            hwnd, title, process, pid = item
    
            # Check if pinned (by identifier since hwnd may have changed)
            window_id = (process, title)
            is_pinned = window_id in self.pinned_identifiers
    
            has_audio = pid in audio_pids
            if not has_audio and pid:
                try:
                    proc = psutil.Process(pid)
                    for child in proc.children(recursive=True):
                        if child.pid in audio_pids:
                            has_audio = True
                            break
                except:
                    pass
    
            # Sort order: pinned first, then audio, then alphabetical
            return (0 if is_pinned else 1, 0 if has_audio else 1, process.lower())



        windows.sort(key=sort_key)
        
        for i, (hwnd, title, process, pid) in enumerate(windows):
            self.window_pids[hwnd] = pid
            has_audio = pid in audio_pids
            
            session_count = len(all_audio_sessions.get(pid, []))
            
            if not has_audio and pid:
                try:
                    proc = psutil.Process(pid)
                    for child in proc.children(recursive=True):
                        if child.pid in audio_pids:
                            has_audio = True
                            session_count = len(all_audio_sessions.get(child.pid, []))
                            break
                except:
                    pass
            
            self.create_window_card(hwnd, title, process, pid, hwnd in old_selection, has_audio, session_count)
                
        self.update_selection_label()
        self.status_var.set(f"{len(windows)} windows")
        self.log_debug(f"Found {len(windows)} windows")
    
    def update_card_style(self, hwnd, selected):
        """Update card visual style based on selection state"""
        if hwnd not in self.window_cards:
            return
            
        card_data = self.window_cards[hwnd]
        
        if selected:
            bg_color = self.colors['card_selected']
            card_data['indicator'].configure(bg=self.colors['checkbox_checked'])
        else:
            bg_color = self.colors['card']
            card_data['indicator'].configure(bg=self.colors['checkbox_unchecked'])
        
        # Update all widgets
        for key in ['card', 'inner', 'left', 'fixed_left', 'info', 'actions']:
            if key in card_data:
                card_data[key].configure(bg=bg_color)
        
        card_data['checkbox'].configure(bg=bg_color, activebackground=bg_color)
        card_data['pin_btn'].configure(bg=bg_color)
        card_data['proc_label'].configure(bg=bg_color)
        card_data['title_label'].configure(bg=bg_color)
        
        for btn in card_data['buttons']:
            btn.configure(bg=bg_color)
        
        card_data['base_bg'] = bg_color


    
    def show_monitor_menu(self, event, hwnd):
        """Show dropdown menu for quick monitor switching"""
        menu = tk.Menu(self.root, tearoff=0, bg=self.colors['card'], fg=self.colors['fg'],
                       activebackground=self.colors['accent'], activeforeground='white',
                       font=('Segoe UI', 9))
        
        current_monitor_idx = self.get_window_monitor_index(hwnd)
        
        for i, mon in enumerate(self.monitors):
            if i == current_monitor_idx:
                menu.add_command(label=f"★ {mon['short_name']}", state='disabled')
            else:
                menu.add_command(label=f"  {mon['short_name']}", 
                               command=lambda h=hwnd, m=mon: self.quick_move_to_monitor(h, m))
        
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
        

    def create_window_card(self, hwnd, title, process, pid, was_selected=False, has_audio=False, session_count=0):
        """Create a sleek window card"""
        # Check if this window should be pinned (from previous session)
        window_id = (process, title)
        is_pinned = window_id in self.pinned_identifiers
        if is_pinned:
            self.pinned_windows.add(hwnd)
        
        base_bg = self.colors['card_selected'] if was_selected else self.colors['card']
        
        card = tk.Frame(self.scrollable_frame, bg=base_bg, highlightthickness=0)
        card.pack(fill=tk.X, pady=2, padx=2)
        
        inner = tk.Frame(card, bg=base_bg)
        inner.pack(fill=tk.X, padx=10, pady=8)
        
        # Actions frame FIRST (pack to right) - fixed width so buttons don't get cut off
        actions = tk.Frame(inner, bg=base_bg)
        actions.pack(side=tk.RIGHT, padx=(5, 0))
        # Don't use pack_propagate(False) - let it size naturally
        
        btn_style = {'bg': base_bg, 'fg': self.colors['muted'],
                    'font': ('Segoe UI', 10), 'bd': 0, 'padx': 4, 'pady': 2,
                    'activebackground': self.colors['card_hover'],
                    'activeforeground': self.colors['fg'], 'cursor': 'hand2'}
        
        monitor_btn = tk.Button(actions, text="◀", width=2, **btn_style)
        monitor_btn.configure(command=lambda e=None, h=hwnd, b=monitor_btn: self.show_monitor_menu_btn(h, b))
        monitor_btn.pack(side=tk.LEFT)
        
        focus_btn = tk.Button(actions, text="◉", width=2, command=lambda h=hwnd: self.focus_window(h), **btn_style)
        focus_btn.pack(side=tk.LEFT)
        
        minmax_btn = tk.Button(actions, text="□", width=2, command=lambda h=hwnd: self.toggle_minmax(h), **btn_style)
        minmax_btn.pack(side=tk.LEFT)
        
        is_tracked_muted = pid in self.muted_pids
        if is_tracked_muted:
            audio_icon = "🔇"
        else:
            is_muted = self.get_app_mute(pid, log_lookup=False) if pid and has_audio else False
            audio_icon = "🔇" if is_muted else "🔊"
        
        audio_btn = tk.Button(actions, text=audio_icon, width=2, **btn_style)
        audio_btn.configure(command=lambda h=hwnd, p=pid, b=audio_btn: self.toggle_app_mute(h, p, b))
        audio_btn.bind('<ButtonPress-3>', lambda e, h=hwnd, p=pid, b=audio_btn: self.on_app_volume_press(e, h, p, b))
        audio_btn.bind('<ButtonRelease-3>', self.on_volume_release)
        audio_btn.pack(side=tk.LEFT)
        
        # Left side (will compress as needed)
        left = tk.Frame(inner, bg=base_bg)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Fixed width elements container
        fixed_left = tk.Frame(left, bg=base_bg)
        fixed_left.pack(side=tk.LEFT)
        
        indicator_color = self.colors['checkbox_checked'] if was_selected else self.colors['checkbox_unchecked']
        indicator = tk.Frame(fixed_left, bg=indicator_color, width=4, height=32)
        indicator.pack(side=tk.LEFT, padx=(0, 8))
        indicator.pack_propagate(False)
        
        var = tk.BooleanVar(value=was_selected)
        if was_selected:
            self.selected_windows[hwnd] = True
            
        cb = tk.Checkbutton(fixed_left, variable=var, bg=base_bg,
                            activebackground=base_bg,
                            selectcolor=self.colors['bg'],
                            command=lambda h=hwnd, v=var: self.on_checkbox_changed(h, v))
        cb.pack(side=tk.LEFT)
        
        self.window_checkboxes[hwnd] = var
        
        # Pin to list button
        pin_color = self.colors['accent'] if is_pinned else self.colors['muted']
        pin_btn = tk.Button(fixed_left, text="📌", bg=base_bg, fg=pin_color,
                            font=('Segoe UI', 9), bd=0, padx=4, pady=0,
                            activebackground=self.colors['card_hover'],
                            cursor='hand2')
        pin_btn.configure(command=lambda h=hwnd, p=process, t=title, b=pin_btn: self.toggle_pin_to_list(h, p, t, b))
        pin_btn.pack(side=tk.LEFT, padx=(0, 4))
        
        # Info section - this is what compresses
        info = tk.Frame(left, bg=base_bg)
        info.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        info.columnconfigure(0, weight=1)
        
        is_tracked_muted = pid in self.muted_pids
        
        if session_count > 0:
            audio_indicator = " 🔇" if is_tracked_muted else " 🔊"
        else:
            audio_indicator = ""
        
        # Process name - clips with ellipsis
        proc_label = tk.Label(info, text=process + audio_indicator, bg=base_bg,
                            fg=self.colors['fg'], font=('Segoe UI', 9, 'bold'),
                            anchor='w')
        proc_label.pack(fill=tk.X, anchor='w')
        
        # Title/subtitle - fully compressible using grid instead of pack
        display_title = title[:60] + "…" if len(title) > 60 else title
        title_label = tk.Label(info, text=display_title, bg=base_bg,
                            fg=self.colors['muted'], font=('Segoe UI', 8),
                            anchor='w')
        title_label.pack(fill=tk.X, anchor='w')
        
        # Store reference for dynamic hiding
        card_data = {
            'card': card,
            'inner': inner,
            'left': left,
            'fixed_left': fixed_left,
            'info': info,
            'actions': actions,
            'checkbox': cb,
            'indicator': indicator,
            'buttons': [monitor_btn, focus_btn, minmax_btn, audio_btn],
            'audio_btn': audio_btn,
            'monitor_btn': monitor_btn,
            'pin_btn': pin_btn,
            'proc_label': proc_label,
            'title_label': title_label,
            'base_bg': base_bg,
            'pid': pid,
            'has_audio': has_audio,
            'process': process,
            'title': title
        }
        
        self.window_cards[hwnd] = card_data
        
        def on_enter(e):
            current_base = self.window_cards[hwnd]['base_bg']
            hover_bg = self.colors['card_hover'] if current_base == self.colors['card'] else '#243454'
            
            for widget in [card, inner, left, fixed_left, info, actions]:
                widget.configure(bg=hover_bg)
            proc_label.configure(bg=hover_bg)
            title_label.configure(bg=hover_bg)
            cb.configure(bg=hover_bg, activebackground=hover_bg)
            pin_btn.configure(bg=hover_bg)
            for btn in self.window_cards[hwnd]['buttons']:
                btn.configure(bg=hover_bg)
                
        def on_leave(e):
            current_base = self.window_cards[hwnd]['base_bg']
            
            for widget in [card, inner, left, fixed_left, info, actions]:
                widget.configure(bg=current_base)
            proc_label.configure(bg=current_base)
            title_label.configure(bg=current_base)
            cb.configure(bg=current_base, activebackground=current_base)
            pin_btn.configure(bg=current_base)
            for btn in self.window_cards[hwnd]['buttons']:
                btn.configure(bg=current_base)
                
        card.bind('<Enter>', on_enter)
        card.bind('<Leave>', on_leave)
        
        # Bind resize to handle title visibility
        def on_card_resize(e):
            # Hide title if card is too narrow
            card_width = card.winfo_width()
            if card_width < 280:
                title_label.pack_forget()
            else:
                if not title_label.winfo_ismapped():
                    title_label.pack(fill=tk.X, anchor='w')
        
        card.bind('<Configure>', on_card_resize)
        #end of create window card

    def toggle_pin_to_list(self, hwnd, process, title, btn):
        """Toggle whether a window is pinned to the top of the list"""
        window_id = (process, title)
    
        if hwnd in self.pinned_windows:
            # Unpin
            self.pinned_windows.discard(hwnd)
            self.pinned_identifiers.discard(window_id)
            btn.configure(fg=self.colors['muted'])
            self.status_var.set(f"Unpinned {process}")
            self.log_debug(f"Unpinned from list: {process} - {title[:30]}")
        else:
            # Pin
            self.pinned_windows.add(hwnd)
            self.pinned_identifiers.add(window_id)
            btn.configure(fg=self.colors['accent'])
            self.status_var.set(f"Pinned {process}")
            self.log_debug(f"Pinned to list: {process} - {title[:30]}")
    
        # Re-sort the list
        self.resort_window_list()

    #add the resort method here:
    def resort_window_list(self):
        """Re-sort window cards with pinned items at top"""
        # Get all card frames in current order
        cards_info = []
        for hwnd, card_data in self.window_cards.items():
            is_pinned = hwnd in self.pinned_windows
            has_audio = card_data.get('has_audio', False)
            process = card_data.get('process', '').lower()
            cards_info.append((hwnd, card_data['card'], is_pinned, has_audio, process))
        
        # Sort: pinned first, then audio, then alphabetical
        cards_info.sort(key=lambda x: (0 if x[2] else 1, 0 if x[3] else 1, x[4]))
        
        # Repack in new order
        for hwnd, card, is_pinned, has_audio, process in cards_info:
            card.pack_forget()
        
        for hwnd, card, is_pinned, has_audio, process in cards_info:
            card.pack(fill=tk.X, pady=2, padx=2)


    def show_monitor_menu_btn(self, hwnd, btn):
        """Show monitor menu positioned relative to button"""
        x = btn.winfo_rootx()
        y = btn.winfo_rooty() + btn.winfo_height()
        
        menu = tk.Menu(self.root, tearoff=0, bg=self.colors['card'], fg=self.colors['fg'],
                       activebackground=self.colors['accent'], activeforeground='white',
                       font=('Segoe UI', 9))
        
        current_monitor_idx = self.get_window_monitor_index(hwnd)
        
        for i, mon in enumerate(self.monitors):
            if i == current_monitor_idx:
                menu.add_command(label=f"★ {mon['short_name']}", state='disabled')
            else:
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
        
        currently_muted = pid in self.muted_pids
        new_mute = not currently_muted
        
        if self.set_app_mute(pid, new_mute):
            btn.configure(text="🔇" if new_mute else "🔊")
            self.status_var.set("Muted" if new_mute else "Unmuted")
        else:
            if new_mute:
                self.muted_pids.add(pid)
            else:
                self.muted_pids.discard(pid)
            btn.configure(text="🔇" if new_mute else "🔊")
            self.status_var.set(f"{'Muted' if new_mute else 'Unmuted'} (will apply when audio starts)")
            self.log_debug(f"Queued mute={new_mute} for PID {pid}")
    
    def on_app_volume_press(self, event, hwnd, pid, btn):
        """Handle right-click press on app audio button"""
        if not pid:
            self.status_var.set("No audio session for this app")
            return
        
        current_volume = self.get_app_volume(pid)
        if current_volume is None:
            current_volume = self.volume_levels.get(pid, 1.0)
        
        self.log_debug(f"Opening volume slider for PID {pid}, current: {current_volume:.0%}")
        
        def on_change(v):
            self.set_app_volume(pid, v, log_lookup=False)
        
        def on_close(final_vol):
            self.log_debug(f"Volume slider closed for PID {pid}, final: {final_vol:.0%}")
            self.update_audio_btn(hwnd, pid, btn)
        
        self.show_volume_slider(event, current_volume, on_change, on_close, "App Volume")
    
    def on_bulk_volume_press(self, event):
        """Handle right-click press on bulk volume button"""
        selected = self.get_selected_windows()
        
        if not selected:
            current_volume = self.get_system_volume()
            self.log_debug(f"Opening system volume slider, current: {current_volume:.0%}")
            
            def on_change(v):
                self.set_system_volume(v)
            
            def on_close(final_vol):
                self.log_debug(f"System volume slider closed, final: {final_vol:.0%}")
            
            self.show_volume_slider(event, current_volume, on_change, on_close, "System")
        else:
            self.log_debug(f"Opening bulk volume slider for {len(selected)} apps")
            
            def on_change(v):
                self.set_selected_volumes(v)
            
            def on_close(final_vol):
                self.log_debug(f"Bulk volume slider closed, final: {final_vol:.0%}")
                self.update_all_audio_btns()
            
            self.show_volume_slider(event, 1.0, on_change, on_close, f"{len(selected)} Apps")
    
    def on_volume_release(self, event):
        """Handle right-click release"""
        self.close_volume_slider()
    
    def show_volume_slider(self, event, initial_volume, on_change, on_close=None, title="Volume"):
        """Create a floating volume slider window"""
        self.close_volume_slider()
        
        slider_win = tk.Toplevel(self.root)
        slider_win.overrideredirect(True)
        slider_win.attributes('-topmost', True)
        
        x = event.x_root - 30
        y = event.y_root - 180
        
        slider_height = 170
        slider_width = 60
        
        slider_win.geometry(f"{slider_width}x{slider_height}+{x}+{y}")
        slider_win.configure(bg=self.colors['card'])
        
        self.volume_slider_window = slider_win
        self.volume_slider_on_close = on_close
        self.volume_slider_on_change = on_change
        
        self.slider_start_y = event.y_root
        self.slider_start_volume = initial_volume
        self.current_slider_volume = initial_volume
        self.slider_is_dragging = True
        
        title_lbl = tk.Label(slider_win, text=title, bg=self.colors['card'],
                             fg=self.colors['muted'], font=('Segoe UI', 7))
        title_lbl.pack(pady=(8, 0))
        
        self.vol_var = tk.StringVar(value=f"{int(initial_volume * 100)}%")
        vol_label = tk.Label(slider_win, textvariable=self.vol_var, bg=self.colors['card'],
                             fg=self.colors['fg'], font=('Segoe UI', 11, 'bold'))
        vol_label.pack(pady=(2, 8))
        
        canvas_height = slider_height - 70
        self.slider_canvas = tk.Canvas(slider_win, width=40, height=canvas_height,
                                        bg=self.colors['slider_bg'], highlightthickness=0)
        self.slider_canvas.pack(pady=(0, 8))
        
        self.track_x = 20
        self.track_top = 5
        self.track_bottom = canvas_height - 5
        self.track_height = self.track_bottom - self.track_top
        
        self.slider_canvas.create_rectangle(self.track_x - 4, self.track_top, 
                                             self.track_x + 4, self.track_bottom,
                                             fill=self.colors['border'], outline='')
        
        handle_y = self.track_bottom - (initial_volume * self.track_height)
        
        self.fill_rect = self.slider_canvas.create_rectangle(
            self.track_x - 4, handle_y, self.track_x + 4, self.track_bottom,
            fill=self.colors['slider_fg'], outline='')
        
        self.handle = self.slider_canvas.create_oval(
            self.track_x - 10, handle_y - 10, self.track_x + 10, handle_y + 10,
            fill=self.colors['accent'], outline=self.colors['fg'], width=2)
        
        self.slider_sensitivity = 150
        
        self.root.bind('<Motion>', self.on_slider_motion)
        slider_win.bind('<Motion>', self.on_slider_motion)
    
    def on_slider_motion(self, event):
        """Handle mouse motion while slider is open"""
        if not self.volume_slider_window or not self.slider_canvas:
            return
        
        if self.slider_start_y is None or self.slider_start_volume is None:
            return
        
        if not self.slider_is_dragging:
            return
        
        try:
            delta_y = self.slider_start_y - event.y_root
            volume_change = delta_y / self.slider_sensitivity
            new_volume = self.slider_start_volume + volume_change
            new_volume = max(0.0, min(1.0, new_volume))
            
            self.current_slider_volume = new_volume
            
            handle_y = self.track_bottom - (new_volume * self.track_height)
            
            self.slider_canvas.coords(self.handle, 
                                       self.track_x - 10, handle_y - 10, 
                                       self.track_x + 10, handle_y + 10)
            
            self.slider_canvas.coords(self.fill_rect, 
                                       self.track_x - 4, handle_y, 
                                       self.track_x + 4, self.track_bottom)
            
            self.vol_var.set(f"{int(new_volume * 100)}%")
            
            if self.volume_slider_on_change:
                self.volume_slider_on_change(new_volume)
                
        except Exception:
            pass
    
    def close_volume_slider(self):
        """Close the volume slider"""
        self.slider_is_dragging = False
        
        try:
            self.root.unbind('<Motion>')
        except:
            pass
        
        final_volume = self.current_slider_volume if self.current_slider_volume is not None else 1.0
        
        if self.volume_slider_on_close:
            try:
                self.volume_slider_on_close(final_volume)
            except:
                pass
            self.volume_slider_on_close = None
        
        if self.volume_slider_window:
            try:
                self.volume_slider_window.destroy()
            except:
                pass
            self.volume_slider_window = None
        
        self.slider_canvas = None
        self.volume_slider_on_change = None
        self.slider_start_y = None
        self.slider_start_volume = None
        self.current_slider_volume = None
    
    def update_audio_btn(self, hwnd, pid, btn):
        """Update audio button icon"""
        if hwnd in self.window_cards:
            is_muted = pid in self.muted_pids
            btn.configure(text="🔇" if is_muted else "🔊")
    
    def update_all_audio_btns(self):
        """Update all audio button icons"""
        for hwnd, card_data in self.window_cards.items():
            pid = card_data.get('pid')
            if pid:
                is_muted = pid in self.muted_pids
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
            else:
                self.status_var.set("Failed to mute system")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid:
                self.set_app_mute(pid, True)
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
            else:
                self.status_var.set("Failed to unmute system")
            return
        
        count = 0
        for hwnd in selected:
            pid = self.window_pids.get(hwnd)
            if pid:
                self.set_app_mute(pid, False)
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