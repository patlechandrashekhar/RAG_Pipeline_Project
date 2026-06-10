"""
Keep your computer awake while running long processes.
This script prevents the screen from turning off and system from sleeping.
"""

import time
import sys
import threading
from datetime import datetime

# Try different methods based on OS
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False
    print("Note: Install pyautogui for better keep-alive: pip install pyautogui")

# Windows-specific
if sys.platform == "win32":
    import ctypes

    # Constants for Windows
    ES_CONTINUOUS = 0x80000000
    ES_SYSTEM_REQUIRED = 0x00000001
    ES_DISPLAY_REQUIRED = 0x00000002
    ES_AWAYMODE_REQUIRED = 0x00000040

    def keep_awake_windows():
        """Prevent Windows from sleeping."""
        ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
        )
        print("[Windows] Sleep mode disabled")

    def allow_sleep_windows():
        """Allow Windows to sleep again."""
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        print("[Windows] Sleep mode re-enabled")

def mouse_jiggler():
    """Move mouse slightly every few minutes to keep screen active."""
    if not PYAUTOGUI_AVAILABLE:
        return

    print("Mouse jiggler started - move mouse to corner to stop")
    pyautogui.FAILSAFE = True  # Move mouse to corner to stop

    try:
        while True:
            # Get current position
            x, y = pyautogui.position()

            # Move mouse slightly
            pyautogui.moveTo(x + 1, y, duration=0.1)
            pyautogui.moveTo(x, y, duration=0.1)

            # Wait 2 minutes
            time.sleep(120)

            # Show activity
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Keeping system awake...")

    except pyautogui.FailSafeException:
        print("Mouse jiggler stopped (mouse moved to corner)")
    except KeyboardInterrupt:
        print("Mouse jiggler stopped")

def keyboard_press():
    """Simulate keyboard press to keep system awake."""
    if not PYAUTOGUI_AVAILABLE:
        return

    print("Keyboard keeper started - Press Ctrl+C to stop")

    try:
        while True:
            # Press Shift key (minimal impact)
            pyautogui.press('shift')

            # Wait 4 minutes
            time.sleep(240)

            print(f"[{datetime.now().strftime('%H:%M:%S')}] Keeping system awake...")

    except KeyboardInterrupt:
        print("Keyboard keeper stopped")

def main():
    """Main keep-alive function."""
    print("=" * 60)
    print("SYSTEM KEEP-ALIVE ACTIVE")
    print("=" * 60)
    print("Your screen and system will stay awake")
    print("Press Ctrl+C to stop and allow sleep again")
    print("-" * 60)

    # Windows-specific keep-awake
    if sys.platform == "win32":
        keep_awake_windows()

    # Start mouse jiggler in background if available
    if PYAUTOGUI_AVAILABLE:
        jiggler_thread = threading.Thread(target=mouse_jiggler, daemon=True)
        jiggler_thread.start()

    try:
        # Keep running
        while True:
            time.sleep(60)
            print(f"[{datetime.now().strftime('%H:%M:%S')}] System kept awake")

    except KeyboardInterrupt:
        print("\n" + "=" * 60)
        print("Stopping keep-alive...")

        # Re-enable sleep on Windows
        if sys.platform == "win32":
            allow_sleep_windows()

        print("System can now sleep normally")
        print("=" * 60)

if __name__ == "__main__":
    main()