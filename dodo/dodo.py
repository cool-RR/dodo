#!/usr/bin/env python
from __future__ import annotations

import os
import sys
import time
import threading
import pathlib
from typing import Any, Optional
import re
import wx
import wx.adv
import ctypes
from ctypes import wintypes
import pyvda
import win32gui
import win32con
import win32api
import yaml
import subprocess
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

try:
    from python_toolbox.misc_tools import RotatingLogStream
    RotatingLogStream.install(pathlib.Path.home() / '.dodo' / 'log')
except ModuleNotFoundError:
    pass


class Monitor:
    """Represents a monitor with its position and size."""
    def __init__(self, index: int, handle: int, left: int, top: int, width: int, height: int):
        self.index = index
        self.handle = handle
        self.left = left
        self.top = top
        self.width = width
        self.height = height

    @property
    def right(self):
        return self.left + self.width

    @property
    def bottom(self):
        return self.top + self.height

    @staticmethod
    def get_all():
        """Get all monitors in the system."""
        monitors = []

        def enum_monitors_callback(hmonitor, hdc, rect, data):
            index = len(monitors)
            monitor = Monitor(
                index=index,
                handle=hmonitor,
                left=rect.contents.left,
                top=rect.contents.top,
                width=rect.contents.right - rect.contents.left,
                height=rect.contents.bottom - rect.contents.top
            )
            monitors.append(monitor)
            return True

        class RECT(ctypes.Structure):
            _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long),
                       ('right', ctypes.c_long), ('bottom', ctypes.c_long)]

        MonitorEnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_ulong,
                                            ctypes.c_ulong, ctypes.POINTER(RECT), ctypes.c_long)
        callback = MonitorEnumProc(enum_monitors_callback)
        ctypes.windll.user32.EnumDisplayMonitors(None, None, callback, 0)

        return monitors


class DesktopNumberOverlay(wx.Frame):
    """Single small overlay window showing desktop number."""
    def __init__(self, desktop_number: int, x: int, y: int, desktop_name: Optional[str] = None):
        super().__init__(None, style=wx.FRAME_NO_TASKBAR | wx.STAY_ON_TOP | wx.NO_BORDER)

        self.desktop_number = desktop_number
        self.desktop_name = desktop_name

        # Set font - large and bold
        font = wx.Font(72, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)

        # Create a temporary DC to measure text size
        temp_bmp = wx.Bitmap(1, 1)
        temp_dc = wx.MemoryDC(temp_bmp)
        temp_dc.SetFont(font)
        # Display "0" for desktop 10, otherwise show the desktop number
        text = "0" if desktop_number == 10 else str(desktop_number)
        text_width, text_height = temp_dc.GetTextExtent(text)

        # Measure name text if present
        name_width, name_height = 0, 0
        if desktop_name:
            name_font = wx.Font(24, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
            temp_dc.SetFont(name_font)
            name_width, name_height = temp_dc.GetTextExtent(desktop_name)

        temp_dc.SelectObject(wx.NullBitmap)

        # Add margin around the text (20px on each side)
        margin = 20
        window_width = max(text_width, name_width) + margin * 2
        window_height = text_height + margin * 2

        # Add space for name if present (with 10px spacing)
        if desktop_name:
            window_height += name_height + 10

        # Position and size the window
        self.SetSize((window_width, window_height))
        self.SetPosition((x, y))

        # Make window semi-transparent (70% opacity = 179 out of 255)
        self.SetTransparent(179)

        # Make window click-through
        hwnd = self.GetHandle()
        extended_style = ctypes.windll.user32.GetWindowLongW(hwnd, win32con.GWL_EXSTYLE)
        ctypes.windll.user32.SetWindowLongW(
            hwnd,
            win32con.GWL_EXSTYLE,
            extended_style | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_LAYERED
        )

        # Setup drawing
        self.Bind(wx.EVT_PAINT, self.on_paint)

        self.Show()

    def on_paint(self, event):
        """Draw desktop number with black background."""
        dc = wx.PaintDC(self)
        width, height = self.GetClientSize()

        # Draw black background
        dc.SetBrush(wx.Brush(wx.Colour(0, 0, 0)))
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.DrawRectangle(0, 0, width, height)

        # Set font and draw pink-white text for number
        font = wx.Font(72, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        dc.SetFont(font)
        dc.SetTextForeground(wx.Colour(255, 182, 193))  # Light pink

        # Display "0" for desktop 10, otherwise show the desktop number
        text = "0" if self.desktop_number == 10 else str(self.desktop_number)
        text_width, text_height = dc.GetTextExtent(text)

        # Calculate vertical positioning
        if self.desktop_name:
            # If there's a name, position number higher
            margin = 20
            number_y = margin
            x = (width - text_width) // 2
            dc.DrawText(text, x, number_y)

            # Draw name below number in teal-white
            name_font = wx.Font(24, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
            dc.SetFont(name_font)
            dc.SetTextForeground(wx.Colour(72, 209, 204))  # Medium turquoise
            name_width, name_height = dc.GetTextExtent(self.desktop_name)
            name_x = (width - name_width) // 2
            name_y = number_y + text_height + 10
            dc.DrawText(self.desktop_name, name_x, name_y)
        else:
            # Center the number if no name
            x = (width - text_width) // 2
            y = (height - text_height) // 2
            dc.DrawText(text, x, y)


class DesktopNumberOverlayManager:
    """Manages multiple overlay windows (one per monitor) and auto-closes them."""
    def __init__(self, desktop_number: int, desktop_name: Optional[str] = None):
        self.overlays = []
        self.timer_id = None

        # Get all monitors
        monitors = Monitor.get_all()

        # Create one overlay per monitor
        for monitor in monitors:
            # Position at top-left of each monitor with 20px padding
            overlay = DesktopNumberOverlay(desktop_number, monitor.left + 20, monitor.top + 20, desktop_name)
            self.overlays.append(overlay)

        # Setup delayed cleanup after 1.5 seconds using CallLater
        if self.overlays:
            self.timer_id = wx.CallLater(1500, self.cleanup)

    def cleanup(self):
        """Destroy all overlays immediately."""
        # Cancel timer if it hasn't fired yet
        if self.timer_id and self.timer_id.IsRunning():
            self.timer_id.Stop()
            self.timer_id = None

        # First pass: Hide all overlays to make them disappear immediately
        for overlay in self.overlays:
            try:
                if overlay:
                    overlay.Hide()
                    overlay.Show(False)
            except Exception as e:
                print(f'Error hiding overlay: {e}')

        # Force processing of pending events to ensure all hides take effect
        try:
            wx.SafeYield()
        except Exception as e:
            print(f'Error in SafeYield: {e}')

        # Second pass: Destroy all overlay windows
        for overlay in self.overlays:
            try:
                if overlay:
                    overlay.Destroy()
            except Exception as e:
                print(f'Error destroying overlay: {e}')

        self.overlays.clear()

class VirtualDesktopAccessor:
    """Access Windows Virtual Desktop functionality using pyvda library"""

    def __init__(self, frame: Optional[DodoFrame] = None) -> None:
        self.current_desktop_number: Optional[int] = None
        self.previous_desktop_number: Optional[int] = None
        self.frame = frame
        self.monitoring = False
        self.monitor_thread: Optional[threading.Thread] = None
        self.overlay_manager: Optional[DesktopNumberOverlayManager] = None
        self.toxita_yaml_path = pathlib.Path.home() / '.dodo' / 'toxita.yaml'
        self.toxita_error_cache: set[tuple[int, int]] = set()  # (hwnd, desktop_number) pairs

        # Auto-pinning feature
        self.always_pinned_path = pathlib.Path.home() / '.dodo' / 'always_pinned'
        self.always_pinned_patterns: list[re.Pattern] = []
        self.unpinned_windows: set[int] = set()  # Window handles that were manually unpinned
        self.auto_pin_monitoring = False
        self.auto_pin_thread: Optional[threading.Thread] = None

        # Desktop names feature
        self.names_yaml_path = pathlib.Path.home() / '.dodo' / 'names.yaml'
        self.desktop_names: dict[int, str] = {}  # desktop_number -> name

        try:
            # Test if pyvda is working
            current = pyvda.VirtualDesktop.current()
            self.current_desktop_number = current.number
            print(f'Virtual Desktop Manager initialized (current desktop: {current.number})')

            # Ensure we have 10 desktops
            self.ensure_ten_desktops()

            # Load always_pinned patterns
            self.load_always_pinned_patterns()

            # Load desktop names
            self.load_desktop_names()

        except Exception as e:
            print(f'Failed to initialize Virtual Desktop Manager: {e}')
            print('Note: This requires Windows 10/11 with virtual desktops enabled')

    def load_always_pinned_patterns(self) -> None:
        """Load regex patterns from ~/.dodo/always_pinned file"""
        self.always_pinned_patterns.clear()
        try:
            if self.always_pinned_path.exists():
                with open(self.always_pinned_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#'):
                            try:
                                pattern = re.compile(line)
                                self.always_pinned_patterns.append(pattern)
                                print(f'Loaded auto-pin pattern: {line}')
                            except re.error as e:
                                print(f'Invalid regex pattern "{line}": {e}')
                print(f'Loaded {len(self.always_pinned_patterns)} auto-pin patterns')
            else:
                print(f'No always_pinned file found at {self.always_pinned_path}')
        except Exception as e:
            print(f'Error loading always_pinned patterns: {e}')

    def load_desktop_names(self) -> None:
        """Load desktop names from ~/.dodo/names.yaml file"""
        self.desktop_names.clear()
        try:
            if self.names_yaml_path.exists():
                with open(self.names_yaml_path, 'r', encoding='utf-8') as f:
                    data = yaml.safe_load(f)
                    if data:
                        for key, value in data.items():
                            if isinstance(key, str) and key.startswith('desktop_'):
                                try:
                                    desktop_num = int(key.split('_')[1])
                                    if 1 <= desktop_num <= 10 and value:
                                        self.desktop_names[desktop_num] = str(value)
                                except (IndexError, ValueError):
                                    pass
                print(f'Loaded {len(self.desktop_names)} desktop names')
            else:
                print(f'No names.yaml file found at {self.names_yaml_path}')
        except Exception as e:
            print(f'Error loading desktop names: {e}')

    def ensure_ten_desktops(self) -> None:
        """Ensure there are at least 10 virtual desktops"""
        try:
            desktops = pyvda.get_virtual_desktops()
            current_count = len(desktops)

            if current_count < 10:
                print(f'Creating {10 - current_count} additional desktops '
                      f'(currently have {current_count})')
                for _ in range(10 - current_count):
                    pyvda.VirtualDesktop.create()
                print('Now have 10 virtual desktops')
            else:
                print(f'Already have {current_count} virtual desktops')

        except Exception as e:
            print(f'Error ensuring 10 desktops: {e}')

    def switch_desktop_by_number(self, desktop_number: int) -> None:
        """Switch to desktop by number (1-10)"""
        if desktop_number < 1 or desktop_number > 10:
            print(f'Invalid desktop number: {desktop_number}')
            return

        try:
            current = pyvda.VirtualDesktop.current()
            current_number = current.number
            self.current_desktop_number = current_number

            if current_number == desktop_number:
                print(f'Already on desktop {desktop_number}')
                return

            # pyvda uses 1-based indexing for VirtualDesktop constructor
            desktop = pyvda.VirtualDesktop(desktop_number)
            desktop.go()

            self.previous_desktop_number = current_number
            self.current_desktop_number = desktop_number

            # Show desktop number overlay
            if self.frame:
                wx.CallAfter(self._show_desktop_overlay, desktop_number)

        except Exception as e:
            print(f'Error switching to desktop {desktop_number}: {e}')

    def _show_desktop_overlay(self, desktop_number: int) -> None:
        """Show the desktop number overlay (called via CallAfter)."""
        try:
            # Clean up any existing overlays first
            if self.overlay_manager:
                self.overlay_manager.cleanup()
                self.overlay_manager = None

            # Get desktop name if available
            desktop_name = self.desktop_names.get(desktop_number)
            # Keep a reference to prevent garbage collection before timer fires
            self.overlay_manager = DesktopNumberOverlayManager(desktop_number, desktop_name)
        except Exception as e:
            print(f'Error showing desktop overlay: {e}')

    def switch_to_previous_desktop(self) -> None:
        """Switch back to the previously active desktop"""
        if self.previous_desktop_number is None:
            print('No previous desktop recorded')
            return

        target = self.previous_desktop_number
        self.switch_desktop_by_number(target)

    def move_window_to_desktop(self, desktop_number: int) -> None:
        """Move the active window to a specific desktop"""
        if desktop_number < 1 or desktop_number > 10:
            print(f'Invalid desktop number: {desktop_number}')
            return

        try:
            # Get the active window handle
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                print('No active window found')
                return

            window_title = win32gui.GetWindowText(hwnd)
            print(f'Moving window: {window_title}')

            # Create AppView for the current window
            app_view = pyvda.AppView(hwnd)

            # Get the target desktop (pyvda uses 1-based indexing)
            target_desktop = pyvda.VirtualDesktop(desktop_number)

            # Move window to desktop
            app_view.move(target_desktop)
            print(f'Moved window to desktop {desktop_number}')

        except Exception as e:
            print(f'Error moving window to desktop {desktop_number}: {e}')

    def pin_window(self) -> None:
        """Pin the active window to all desktops or unpin if already pinned"""
        try:
            # Get the active window handle
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                print('No active window found')
                return

            window_title = win32gui.GetWindowText(hwnd)

            # Create AppView for the current window
            app_view = pyvda.AppView(hwnd)

            # Toggle pin state
            if not app_view.is_pinned():
                app_view.pin()
                print(f'Pinned window to all desktops: {window_title}')
                # Remove from unpinned set if it was there
                self.unpinned_windows.discard(hwnd)
            else:
                app_view.unpin()
                print(f'Unpinned window: {window_title}')
                # Add to unpinned set so auto-pin won't re-pin it
                self.unpinned_windows.add(hwnd)

        except Exception as e:
            print(f'Error toggling pin state: {e}')

    def start_monitoring(self) -> None:
        """Start monitoring for automatic desktop changes"""
        if self.monitoring:
            return

        self.monitoring = True
        self.monitor_thread = threading.Thread(target=self._monitor_desktop_changes)
        self.monitor_thread.daemon = True
        self.monitor_thread.start()
        print('Started monitoring for desktop changes')

    def stop_monitoring(self) -> None:
        """Stop monitoring for desktop changes"""
        self.monitoring = False
        if self.monitor_thread:
            self.monitor_thread.join(timeout=2.0)

    def start_auto_pin_monitoring(self) -> None:
        """Start monitoring for windows that should be auto-pinned"""
        if self.auto_pin_monitoring:
            return

        if not self.always_pinned_patterns:
            print('No auto-pin patterns configured')
            return

        self.auto_pin_monitoring = True
        self.auto_pin_thread = threading.Thread(target=self._monitor_auto_pin)
        self.auto_pin_thread.daemon = True
        self.auto_pin_thread.start()
        print('Started auto-pin monitoring')

    def stop_auto_pin_monitoring(self) -> None:
        """Stop monitoring for auto-pinning"""
        self.auto_pin_monitoring = False
        if self.auto_pin_thread:
            self.auto_pin_thread.join(timeout=4.0)

    def _monitor_auto_pin(self) -> None:
        """Monitor thread that checks for windows to auto-pin every 3 seconds"""
        while self.auto_pin_monitoring:
            try:
                time.sleep(3)
                self._check_and_pin_windows()
            except Exception as e:
                print(f'Error in auto-pin monitoring: {e}')

    def _check_and_pin_windows(self) -> None:
        """Check all windows and pin those matching always_pinned patterns"""
        if not self.always_pinned_patterns:
            return

        def enum_windows_callback(hwnd, extra):
            try:
                # Skip invisible windows
                if not win32gui.IsWindowVisible(hwnd):
                    return True

                # Skip windows without titles
                title = win32gui.GetWindowText(hwnd)
                if not title:
                    return True

                # Check if window matches any pattern
                matches_pattern = False
                for pattern in self.always_pinned_patterns:
                    if pattern.search(title):
                        matches_pattern = True
                        break

                if not matches_pattern:
                    return True

                # Check if window was manually unpinned
                if hwnd in self.unpinned_windows:
                    return True

                # Check if window is already pinned
                try:
                    app_view = pyvda.AppView(hwnd)
                    if app_view.is_pinned():
                        return True

                    # Pin the window
                    app_view.pin()
                    print(f'Auto-pinned window: {title}')

                except Exception as e:
                    # Silently skip problematic windows
                    pass

            except Exception as e:
                # Silently skip problematic windows
                pass

            return True

        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            print(f'Error enumerating windows for auto-pin: {e}')

    def _monitor_desktop_changes(self) -> None:
        """Monitor thread that checks for desktop changes every second"""
        while self.monitoring:
            try:
                time.sleep(1)

                # Get the current desktop number
                current = pyvda.VirtualDesktop.current()
                actual_desktop = current.number

                # If desktop has changed externally (not by us)
                if actual_desktop != self.current_desktop_number:
                    # Update our tracking
                    self.previous_desktop_number = self.current_desktop_number
                    self.current_desktop_number = actual_desktop

                    # Show the overlay
                    if self.frame:
                        wx.CallAfter(self._show_desktop_overlay, actual_desktop)

            except Exception as e:
                print(f'Error in desktop monitoring: {e}')

    def get_all_windows_by_desktop(self) -> dict[str, list[tuple[int, str]]]:
        """Get all windows organized by desktop.

        Returns:
            Dictionary with keys 'pinned', 'desktop_1' through 'desktop_10'
            Each value is a list of (hwnd, title) tuples
        """
        result = {
            'pinned': [],
            **{f'desktop_{i}': [] for i in range(1, 11)}
        }

        def enum_windows_callback(hwnd, extra):
            try:
                # Skip invisible windows
                if not win32gui.IsWindowVisible(hwnd):
                    return True

                # Skip windows without titles
                title = win32gui.GetWindowText(hwnd)
                if not title:
                    return True

                # Filter out \r and \n characters from title
                title = title.replace('\r', '').replace('\n', '')

                # Check if window is pinned
                try:
                    app_view = pyvda.AppView(hwnd)
                    if app_view.is_pinned():
                        result['pinned'].append((hwnd, title))
                        return True
                except:
                    pass

                # Get window's desktop
                try:
                    app_view = pyvda.AppView(hwnd)
                    desktop = app_view.desktop
                    desktop_num = desktop.number

                    if 1 <= desktop_num <= 10:
                        result[f'desktop_{desktop_num}'].append((hwnd, title))
                except:
                    pass

            except Exception as e:
                # Silently skip problematic windows
                pass

            return True

        try:
            win32gui.EnumWindows(enum_windows_callback, None)
        except Exception as e:
            print(f'Error enumerating windows: {e}')

        return result

    def generate_toxita_yaml(self) -> None:
        """Generate toxita.yaml file with current window layout and open it."""
        try:
            # Get all windows by desktop
            windows_by_desktop = self.get_all_windows_by_desktop()

            # Build YAML structure
            yaml_data = {}

            # Add pinned windows first
            if windows_by_desktop['pinned']:
                yaml_data['pinned'] = {hwnd: title for hwnd, title in windows_by_desktop['pinned']}
            else:
                yaml_data['pinned'] = {'placeholder': 'placeholder'}

            # Add desktops 1-9
            for i in range(1, 10):
                key = f'desktop_{i}'
                if windows_by_desktop[key]:
                    yaml_data[key] = {hwnd: title for hwnd, title in windows_by_desktop[key]}
                else:
                    yaml_data[key] = {'placeholder': 'placeholder'}

            # Add desktop 10 last
            if windows_by_desktop['desktop_10']:
                yaml_data['desktop_10'] = {hwnd: title for hwnd, title in windows_by_desktop['desktop_10']}
            else:
                yaml_data['desktop_10'] = {'placeholder': 'placeholder'}

            # Ensure directory exists
            self.toxita_yaml_path.parent.mkdir(parents=True, exist_ok=True)

            # Write YAML file
            with open(self.toxita_yaml_path, 'w', encoding='utf-8') as f:
                yaml.dump(yaml_data, f, default_flow_style=False, allow_unicode=True, sort_keys=False, width=float('inf'))

            print(f'Generated toxita.yaml with {sum(len(v) for v in windows_by_desktop.values())} windows')

            # Open file with default application
            try:
                os.startfile(str(self.toxita_yaml_path))
            except Exception as e:
                print(f'Error opening toxita.yaml: {e}')

        except Exception as e:
            print(f'Error generating toxita.yaml: {e}')

    def apply_toxita_yaml(self) -> None:
        """Apply window layout from toxita.yaml file."""
        try:
            # Read YAML file
            if not self.toxita_yaml_path.exists():
                print('toxita.yaml does not exist')
                return

            with open(self.toxita_yaml_path, 'r', encoding='utf-8') as f:
                yaml_data = yaml.safe_load(f)

            if not yaml_data:
                print('toxita.yaml is empty')
                return

            # Clear error cache for new application
            self.toxita_error_cache.clear()

            # Process each section
            for section, windows in yaml_data.items():
                if not isinstance(windows, dict):
                    continue

                # Determine target desktop/action
                if section == 'pinned':
                    target_desktop = None  # Will pin
                elif section.startswith('desktop_'):
                    try:
                        desktop_num = int(section.split('_')[1])
                        if 1 <= desktop_num <= 10:
                            target_desktop = desktop_num
                        else:
                            continue
                    except (IndexError, ValueError):
                        continue
                else:
                    continue

                # Process each window
                for hwnd, title in windows.items():
                    try:
                        # Skip placeholder entries
                        if hwnd == 'placeholder':
                            continue

                        hwnd = int(hwnd)

                        # Check if window still exists
                        if not win32gui.IsWindow(hwnd):
                            continue

                        # Create error cache key
                        error_key = (hwnd, target_desktop if target_desktop else -1)

                        if section == 'pinned':
                            # Pin the window
                            try:
                                app_view = pyvda.AppView(hwnd)
                                if not app_view.is_pinned():
                                    app_view.pin()
                                    print(f'Pinned window {hwnd}: {win32gui.GetWindowText(hwnd)}')
                            except Exception as e:
                                if error_key not in self.toxita_error_cache:
                                    print(f'Error pinning window {hwnd}: {e}')
                                    self.toxita_error_cache.add(error_key)
                        else:
                            # Move to specific desktop
                            try:
                                app_view = pyvda.AppView(hwnd)

                                # Unpin if currently pinned
                                if app_view.is_pinned():
                                    app_view.unpin()

                                # Move to target desktop
                                target_desktop_obj = pyvda.VirtualDesktop(target_desktop)
                                app_view.move(target_desktop_obj)
                                print(f'Moved window {hwnd} to desktop {target_desktop}: {win32gui.GetWindowText(hwnd)}')

                            except Exception as e:
                                if error_key not in self.toxita_error_cache:
                                    print(f'Error moving window {hwnd} to desktop {target_desktop}: {e}')
                                    self.toxita_error_cache.add(error_key)

                    except (ValueError, TypeError):
                        # Invalid hwnd format
                        continue

            print('Applied toxita.yaml')

        except yaml.YAMLError as e:
            print(f'Error parsing toxita.yaml: {e}')
        except Exception as e:
            print(f'Error applying toxita.yaml: {e}')

class ToxitaFileHandler(FileSystemEventHandler):
    """Handles file system events for toxita.yaml"""

    def __init__(self, vda: VirtualDesktopAccessor):
        super().__init__()
        self.vda = vda
        self.last_modified = 0
        self.debounce_seconds = 0.5

    def on_modified(self, event):
        """Called when toxita.yaml is modified"""
        if event.is_directory:
            return

        if pathlib.Path(event.src_path) == self.vda.toxita_yaml_path:
            # Debounce multiple rapid modifications
            current_time = time.time()
            if current_time - self.last_modified < self.debounce_seconds:
                return

            self.last_modified = current_time
            print('toxita.yaml modified, applying changes...')
            self.vda.apply_toxita_yaml()

class Dodo:
    def __init__(self, frame: Optional[DodoFrame] = None) -> None:
        self.running: bool = True
        self.vda = VirtualDesktopAccessor(frame)
        self.toxita_observer: Optional[Observer] = None

    def start_toxita_watcher(self) -> None:
        """Start watching toxita.yaml for changes"""
        try:
            # Ensure directory exists
            self.vda.toxita_yaml_path.parent.mkdir(parents=True, exist_ok=True)

            # Create and start observer
            self.toxita_observer = Observer()
            event_handler = ToxitaFileHandler(self.vda)
            self.toxita_observer.schedule(
                event_handler,
                str(self.vda.toxita_yaml_path.parent),
                recursive=False
            )
            self.toxita_observer.start()
            print(f'Started watching {self.vda.toxita_yaml_path}')

        except Exception as e:
            print(f'Error starting toxita watcher: {e}')

    def stop_toxita_watcher(self) -> None:
        """Stop watching toxita.yaml"""
        if self.toxita_observer:
            self.toxita_observer.stop()
            self.toxita_observer.join(timeout=2.0)
            print('Stopped toxita watcher')

    def run_loop(self) -> None:
        """Run the main loop to keep the program running."""
        print('Starting Dodo Desktop Switcher')
        print('Use the system tray icon to switch desktops')

        try:
            while self.running:
                time.sleep(1)
        except KeyboardInterrupt:
            print('Keyboard interrupt received, stopping...')
            self.running = False
        except Exception as e:
            sys.excepthook(type(e), e, e.__traceback__)
        finally:
            self.cleanup()

    def cleanup(self) -> None:
        """Cleanup."""
        self.stop_toxita_watcher()
        print('Dodo shutting down')

class DodoTaskBarIcon(wx.adv.TaskBarIcon):
    def __init__(self, frame: DodoFrame) -> None:
        super(DodoTaskBarIcon, self).__init__()
        self.frame = frame

        # Create a custom icon with 'DD' in blue color
        icon_size = 16
        bmp = wx.Bitmap(icon_size, icon_size)
        dc = wx.MemoryDC(bmp)

        # Set white background
        dc.SetBackground(wx.Brush(wx.Colour(255, 255, 255)))
        dc.Clear()

        # Draw 'DD' in blue color
        blue_color = wx.Colour(0, 100, 200)
        dc.SetTextForeground(blue_color)
        dc.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL,
                           wx.FONTWEIGHT_BOLD))

        # Draw text centered
        text = 'DD'
        text_width, text_height = dc.GetTextExtent(text)
        x = (icon_size - text_width) // 2
        y = (icon_size - text_height) // 2
        dc.DrawText(text, x, y)

        dc.SelectObject(wx.NullBitmap)

        # Create icon from bitmap
        icon = wx.Icon()
        icon.CopyFromBitmap(bmp)

        self.SetIcon(icon, 'Dodo Desktop Switcher')
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self.on_left_down)

    def on_left_down(self, event: wx.Event) -> None:
        self.PopupMenu(self.CreatePopupMenu())

    def CreatePopupMenu(self) -> wx.Menu:
        menu = wx.Menu()

        # Get current desktop number
        try:
            current = pyvda.VirtualDesktop.current()
            current_desktop = current.number
        except:
            current_desktop = None

        # Show current desktop indicator
        if current_desktop is not None:
            current_item = menu.Append(wx.ID_ANY, f'Current Desktop: {current_desktop}')
            current_item.Enable(False)
            menu.AppendSeparator()

        # Add desktop switching options
        desktops_menu = wx.Menu()
        for i in range(1, 10):
            label = f'Desktop {i} (Alt+{i})'
            if current_desktop == i:
                label = f'Desktop {i} (Alt+{i}) ✓'
            item = desktops_menu.Append(wx.ID_ANY, label)
            self.Bind(wx.EVT_MENU,
                     lambda event, d=i: self.frame.dodo.vda.switch_desktop_by_number(d),
                     item)

        # Desktop 10
        label = f'Desktop 10 (Alt+0)'
        if current_desktop == 10:
            label = f'Desktop 10 (Alt+0) ✓'
        item = desktops_menu.Append(wx.ID_ANY, label)
        self.Bind(wx.EVT_MENU,
                 lambda event: self.frame.dodo.vda.switch_desktop_by_number(10),
                 item)

        menu.AppendSubMenu(desktops_menu, 'Switch to Desktop')
        menu.AppendSeparator()

        # Toxita button
        toxita_item = menu.Append(wx.ID_ANY, 'Generate Toxita (Ctrl+Alt+-)')
        self.Bind(wx.EVT_MENU,
                 lambda event: self.frame.dodo.vda.generate_toxita_yaml(),
                 toxita_item)
        menu.AppendSeparator()

        about_item = menu.Append(wx.ID_ANY, 'About')
        exit_item = menu.Append(wx.ID_EXIT, 'Exit')

        self.Bind(wx.EVT_MENU, self.on_about, about_item)
        self.Bind(wx.EVT_MENU, self.on_exit, exit_item)

        return menu

    def on_about(self, event: wx.Event) -> None:
        wx.MessageBox(
            'Dodo Desktop Switcher\n\n'
            'Keyboard shortcuts:\n'
            'Alt+1 to Alt+9: Switch to desktop 1-9\n'
            'Alt+0: Switch to desktop 10\n'
            'Alt+-: Switch to the previously active desktop\n'
            'Alt+Shift+1 to Alt+Shift+9: Move window to desktop 1-9\n'
            'Alt+Shift+0: Move window to desktop 10\n'
            'Alt+Shift+`: Pin window to all desktops\n'
            'Ctrl+Alt+-: Generate toxita.yaml with window layout\n'
            'Ctrl+Alt+Shift+-: Show desktop number overlay\n\n'
            'Toxita feature:\n'
            'Press Ctrl+Alt+- to create ~/.dodo/toxita.yaml with all\n'
            'windows organized by desktop. Edit and save the file to\n'
            'move windows between desktops automatically.\n\n'
            'Note: Requires Windows 10/11 with virtual desktops enabled',
            'About Dodo', wx.OK | wx.ICON_INFORMATION)

    def on_exit(self, event: wx.Event) -> None:
        self.frame.dodo.running = False
        wx.CallAfter(self.Destroy)
        self.frame.Close()


class DodoFrame(wx.Frame):
    def __init__(self) -> None:
        super(DodoFrame, self).__init__(None, title='Dodo Desktop Switcher', size=(1, 1))
        self.tbicon = DodoTaskBarIcon(self)
        self.dodo = Dodo(self)
        self.dodo_thread: Optional[threading.Thread] = None
        self.hotkey_ids: list[int] = []
        self.hotkey_desktop_map: dict[int, int] = {}
        self.hotkey_move_map: dict[int, int] = {}
        self.hotkey_previous_desktop_id: Optional[int] = None
        self.hotkey_pin_id: Optional[int] = None
        self.hotkey_toxita_id: Optional[int] = None
        self.hotkey_show_overlay_id: Optional[int] = None

        # Hide the frame
        self.Show(False)

        # Register hotkeys
        self.register_hotkeys()

        # Start Dodo in a separate thread
        self.dodo_thread = threading.Thread(target=self.dodo.run_loop)
        self.dodo_thread.daemon = True
        self.dodo_thread.start()

        # Start desktop monitoring
        self.dodo.vda.start_monitoring()

        # Start auto-pin monitoring
        self.dodo.vda.start_auto_pin_monitoring()

        # Start toxita file watcher
        self.dodo.start_toxita_watcher()

        # Bind the close event
        self.Bind(wx.EVT_CLOSE, self.on_close)

    def register_hotkeys(self) -> None:
        """Register system-wide hotkeys using wx"""
        try:
            # Start with ID 100
            hotkey_id = 100

            # Register Alt+1 through Alt+9 for desktops 1-9
            for i in range(1, 10):
                if self.RegisterHotKey(hotkey_id, win32con.MOD_ALT, ord(str(i))):
                    self.hotkey_desktop_map[hotkey_id] = i
                    self.hotkey_ids.append(hotkey_id)
                    print(f'Registered Alt+{i} for desktop {i}')
                hotkey_id += 1

            # Register Alt+0 for desktop 10
            if self.RegisterHotKey(hotkey_id, win32con.MOD_ALT, ord('0')):
                self.hotkey_desktop_map[hotkey_id] = 10
                self.hotkey_ids.append(hotkey_id)
                print(f'Registered Alt+0 for desktop 10')
            hotkey_id += 1

            # Register Alt+- for returning to the previous desktop
            if self.RegisterHotKey(hotkey_id, win32con.MOD_ALT, wx.WXK_F17):
                self.hotkey_previous_desktop_id = hotkey_id
                self.hotkey_ids.append(hotkey_id)
                print('Registered Alt+F17 for previous desktop (AHK should funnel Alt+- to this)')
            hotkey_id += 1

            # Register Alt+Shift+1 through Alt+Shift+9 for moving windows
            for i in range(1, 10):
                if self.RegisterHotKey(hotkey_id,
                                      win32con.MOD_ALT | win32con.MOD_SHIFT,
                                      ord(str(i))):
                    self.hotkey_move_map[hotkey_id] = i
                    self.hotkey_ids.append(hotkey_id)
                    print(f'Registered Alt+Shift+{i} for moving window to desktop {i}')
                hotkey_id += 1

            # Register Alt+Shift+0 for moving window to desktop 10
            if self.RegisterHotKey(hotkey_id,
                                  win32con.MOD_ALT | win32con.MOD_SHIFT,
                                  ord('0')):
                self.hotkey_move_map[hotkey_id] = 10
                self.hotkey_ids.append(hotkey_id)
                print(f'Registered Alt+Shift+0 for moving window to desktop 10')
            hotkey_id += 1

            # Register Alt+Shift+` for pinning/unpinning window
            # The tilde key (~) is VK code 192 (the key to the left of 1)
            if self.RegisterHotKey(hotkey_id,
                                  win32con.MOD_ALT | win32con.MOD_SHIFT,
                                  192):  # VK_OEM_3 (tilde/backtick key)
                self.hotkey_pin_id = hotkey_id
                self.hotkey_ids.append(hotkey_id)
                print('Registered Alt+Shift+` for pinning window')
            hotkey_id += 1

            # Register Ctrl+Alt+Minus for toxita
            # Minus key is VK code 189
            if self.RegisterHotKey(hotkey_id,
                                  win32con.MOD_CONTROL | win32con.MOD_ALT,
                                  189):  # VK_OEM_MINUS
                self.hotkey_toxita_id = hotkey_id
                self.hotkey_ids.append(hotkey_id)
                print('Registered Ctrl+Alt+- for toxita')
            hotkey_id += 1

            # Register Ctrl+Alt+Shift+Minus for showing desktop overlay
            if self.RegisterHotKey(hotkey_id,
                                  win32con.MOD_CONTROL | win32con.MOD_ALT | win32con.MOD_SHIFT,
                                  189):  # VK_OEM_MINUS
                self.hotkey_show_overlay_id = hotkey_id
                self.hotkey_ids.append(hotkey_id)
                print('Registered Ctrl+Alt+Shift+- for showing desktop overlay')

            # Bind the hotkey event handler
            self.Bind(wx.EVT_HOTKEY, self.on_hotkey)

            if self.hotkey_ids:
                print(f'Successfully registered {len(self.hotkey_ids)} hotkeys')
            else:
                print('Warning: No hotkeys were registered')

        except Exception as e:
            print(f'Error registering hotkeys: {e}')

    def on_hotkey(self, event: wx.Event) -> None:
        """Handle hotkey events"""
        hotkey_id = event.GetId()

        if hotkey_id in self.hotkey_desktop_map:
            desktop_num = self.hotkey_desktop_map[hotkey_id]
            self.dodo.vda.switch_desktop_by_number(desktop_num)
        elif hotkey_id in self.hotkey_move_map:
            desktop_num = self.hotkey_move_map[hotkey_id]
            self.dodo.vda.move_window_to_desktop(desktop_num)
        elif hotkey_id == self.hotkey_previous_desktop_id:
            self.dodo.vda.switch_to_previous_desktop()
        elif hotkey_id == self.hotkey_pin_id:
            self.dodo.vda.pin_window()
        elif hotkey_id == self.hotkey_toxita_id:
            self.dodo.vda.generate_toxita_yaml()
        elif hotkey_id == self.hotkey_show_overlay_id:
            # Show overlay for current desktop
            try:
                current = pyvda.VirtualDesktop.current()
                wx.CallAfter(self.dodo.vda._show_desktop_overlay, current.number)
            except Exception as e:
                print(f'Error showing desktop overlay: {e}')

    def on_close(self, event: wx.Event) -> None:
        # Stop desktop monitoring
        self.dodo.vda.stop_monitoring()

        # Stop auto-pin monitoring
        self.dodo.vda.stop_auto_pin_monitoring()

        # Unregister all hotkeys
        for hotkey_id in self.hotkey_ids:
            try:
                self.UnregisterHotKey(hotkey_id)
            except:
                pass

        self.dodo.running = False
        if self.dodo_thread and self.dodo_thread.is_alive():
            self.dodo_thread.join(1.0)
        self.dodo.cleanup()
        self.Destroy()


import click

def get_startup_folder() -> pathlib.Path:
    """Get the Windows Startup folder path"""
    return pathlib.Path(os.environ['APPDATA']) / 'Microsoft' / 'Windows' / 'Start Menu' / 'Programs' / 'Startup'

def get_shortcut_path() -> pathlib.Path:
    """Get the path where Dodo's startup shortcut would be"""
    return get_startup_folder() / 'Dodo.lnk'

def install_to_startup() -> None:
    """Add Dodo to Windows startup folder"""
    try:
        import win32com.client

        shortcut_path = get_shortcut_path()

        if shortcut_path.exists():
            print(f'Dodo is already installed to startup: {shortcut_path}')
            return

        # Get the path to pythonw.exe (windowless Python)
        python_exe = pathlib.Path(sys.executable)
        if python_exe.name.lower() == 'python.exe':
            pythonw_exe = python_exe.parent / 'pythonw.exe'
            # Fall back to python.exe if pythonw doesn't exist
            if not pythonw_exe.exists():
                pythonw_exe = python_exe
        else:
            # Already using pythonw or other variant
            pythonw_exe = python_exe

        # Create shortcut
        shell = win32com.client.Dispatch('WScript.Shell')
        shortcut = shell.CreateShortcut(str(shortcut_path))
        shortcut.TargetPath = str(pythonw_exe)
        shortcut.Arguments = '-m dodo'
        shortcut.WorkingDirectory = str(pathlib.Path.home())
        shortcut.Description = 'Dodo Desktop Switcher'
        shortcut.save()

        print(f'✓ Dodo installed to startup: {shortcut_path}')
        print('Dodo will now start automatically when you log in to Windows.')

    except Exception as e:
        print(f'Error installing to startup: {e}')
        sys.exit(1)

def uninstall_from_startup() -> None:
    """Remove Dodo from Windows startup folder"""
    try:
        shortcut_path = get_shortcut_path()

        if not shortcut_path.exists():
            print('Dodo is not currently installed to startup.')
            return

        shortcut_path.unlink()
        print(f'✓ Dodo removed from startup: {shortcut_path}')
        print('Dodo will no longer start automatically.')

    except Exception as e:
        print(f'Error removing from startup: {e}')
        sys.exit(1)

def check_startup_status() -> None:
    """Check if Dodo is installed to startup"""
    shortcut_path = get_shortcut_path()
    if shortcut_path.exists():
        print(f'✓ Dodo is installed to startup: {shortcut_path}')
    else:
        print('Dodo is not currently installed to startup.')
        print('Run "dodo --install" to add it to startup.')

@click.command()
@click.option('--cli', is_flag=True, help='Run in command-line mode without GUI')
@click.option('--install', is_flag=True, help='Install Dodo to Windows startup')
@click.option('--uninstall', is_flag=True, help='Remove Dodo from Windows startup')
@click.option('--status', is_flag=True, help='Check if Dodo is installed to startup')
def main(cli: bool, install: bool, uninstall: bool, status: bool) -> None:
    """Dodo Desktop Switcher - Switch Windows virtual desktops with shortcuts"""

    # Handle installation/uninstallation commands
    if install:
        install_to_startup()
        return

    if uninstall:
        uninstall_from_startup()
        return

    if status:
        check_startup_status()
        return

    # Normal operation
    if cli:
        # Command-line mode
        dodo = Dodo()
        try:
            dodo.run_loop()
        finally:
            dodo.cleanup()
    else:
        # GUI mode with system tray icon
        app = wx.App()
        frame = DodoFrame()
        app.MainLoop()

if __name__ == '__main__':
    main()
