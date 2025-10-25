#!/usr/bin/env python
from __future__ import annotations

import os
import sys
import time
import threading
import pathlib
from typing import Any, Optional
import wx
import wx.adv
import ctypes
from ctypes import wintypes
import pyvda
import win32gui
import win32con
import win32api

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
    def __init__(self, desktop_number: int, x: int, y: int):
        super().__init__(None, style=wx.FRAME_NO_TASKBAR | wx.STAY_ON_TOP | wx.NO_BORDER)

        self.desktop_number = desktop_number

        # Set font - large and bold
        font = wx.Font(72, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)

        # Create a temporary DC to measure text size
        temp_bmp = wx.Bitmap(1, 1)
        temp_dc = wx.MemoryDC(temp_bmp)
        temp_dc.SetFont(font)
        # Display "0" for desktop 10, otherwise show the desktop number
        text = "0" if desktop_number == 10 else str(desktop_number)
        text_width, text_height = temp_dc.GetTextExtent(text)
        temp_dc.SelectObject(wx.NullBitmap)

        # Add margin around the text (20px on each side)
        margin = 20
        window_width = text_width + margin * 2
        window_height = text_height + margin * 2

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

        # Set font and draw white text
        font = wx.Font(72, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
        dc.SetFont(font)
        dc.SetTextForeground(wx.Colour(255, 255, 255))

        # Display "0" for desktop 10, otherwise show the desktop number
        text = "0" if self.desktop_number == 10 else str(self.desktop_number)
        text_width, text_height = dc.GetTextExtent(text)

        # Center the text
        x = (width - text_width) // 2
        y = (height - text_height) // 2
        dc.DrawText(text, x, y)


class DesktopNumberOverlayManager:
    """Manages multiple overlay windows (one per monitor) and auto-closes them."""
    def __init__(self, desktop_number: int):
        self.overlays = []
        self.timer = None

        # Get all monitors
        monitors = Monitor.get_all()

        # Create one overlay per monitor
        for monitor in monitors:
            # Position at top-left of each monitor with 20px padding
            overlay = DesktopNumberOverlay(desktop_number, monitor.left + 20, monitor.top + 20)
            self.overlays.append(overlay)

        # Setup timer to close all overlays after 1.5 seconds
        if self.overlays:
            self.timer = wx.Timer()
            self.timer.Bind(wx.EVT_TIMER, self.on_timer)
            self.timer.Start(1500, wx.TIMER_ONE_SHOT)

    def on_timer(self, event):
        """Close all overlays when timer expires."""
        for overlay in self.overlays:
            try:
                overlay.Close()
            except:
                pass
        self.overlays.clear()

class VirtualDesktopAccessor:
    """Access Windows Virtual Desktop functionality using pyvda library"""

    def __init__(self, frame: Optional[DodoFrame] = None) -> None:
        self.current_desktop_number: Optional[int] = None
        self.previous_desktop_number: Optional[int] = None
        self.frame = frame

        try:
            # Test if pyvda is working
            current = pyvda.VirtualDesktop.current()
            self.current_desktop_number = current.number
            print(f'Virtual Desktop Manager initialized (current desktop: {current.number})')

            # Ensure we have 10 desktops
            self.ensure_ten_desktops()

        except Exception as e:
            print(f'Failed to initialize Virtual Desktop Manager: {e}')
            print('Note: This requires Windows 10/11 with virtual desktops enabled')

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
            overlay_manager = DesktopNumberOverlayManager(desktop_number)
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
        """Pin the active window to all desktops"""
        try:
            # Get the active window handle
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                print('No active window found')
                return

            window_title = win32gui.GetWindowText(hwnd)

            # Create AppView for the current window
            app_view = pyvda.AppView(hwnd)

            # Pin the window (only if not already pinned)
            if not app_view.is_pinned():
                app_view.pin()
                print(f'Pinned window to all desktops: {window_title}')
            else:
                print(f'Window already pinned: {window_title}')

        except Exception as e:
            print(f'Error pinning window: {e}')

class Dodo:
    def __init__(self, frame: Optional[DodoFrame] = None) -> None:
        self.running: bool = True
        self.vda = VirtualDesktopAccessor(frame)

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

        # Add desktop switching options
        desktops_menu = wx.Menu()
        for i in range(1, 10):
            item = desktops_menu.Append(wx.ID_ANY, f'Desktop {i} (Alt+{i})')
            self.Bind(wx.EVT_MENU,
                     lambda event, d=i: self.frame.dodo.vda.switch_desktop_by_number(d),
                     item)
        item = desktops_menu.Append(wx.ID_ANY, f'Desktop 10 (Alt+0)')
        self.Bind(wx.EVT_MENU,
                 lambda event: self.frame.dodo.vda.switch_desktop_by_number(10),
                 item)

        menu.AppendSubMenu(desktops_menu, 'Switch to Desktop')
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
            'Alt+Shift+`: Pin window to all desktops\n\n'
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

        # Hide the frame
        self.Show(False)

        # Register hotkeys
        self.register_hotkeys()

        # Start Dodo in a separate thread
        self.dodo_thread = threading.Thread(target=self.dodo.run_loop)
        self.dodo_thread.daemon = True
        self.dodo_thread.start()

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

    def on_close(self, event: wx.Event) -> None:
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
