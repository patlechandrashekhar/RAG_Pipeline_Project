"""
Simple and reliable keep-awake script for Windows.
Prevents screen timeout and system sleep.
"""

import ctypes
import time
from datetime import datetime

# Windows constants
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002

def keep_screen_on():
    """Prevent Windows from turning off the screen or sleeping."""
    print("=" * 60)
    print("SCREEN KEEP-ALIVE ACTIVE")
    print("=" * 60)
    print("Your screen will stay on while this runs")
    print("Press Ctrl+C to stop")
    print("-" * 60)

    # Call Windows API to prevent sleep
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )

    print("[OK] Screen timeout disabled")
    print("[OK] System sleep disabled")
    print("")
    print("Running... (showing heartbeat every minute)")

    start_time = datetime.now()
    counter = 0

    try:
        while True:
            # Wait 60 seconds
            time.sleep(60)

            # Update counter
            counter += 1
            elapsed = datetime.now() - start_time
            elapsed_str = str(elapsed).split('.')[0]  # Remove microseconds

            # Show we're still running
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Active for {elapsed_str} - Heartbeat #{counter}")

            # Refresh the keep-awake state every 10 minutes
            if counter % 10 == 0:
                ctypes.windll.kernel32.SetThreadExecutionState(
                    ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
                )
                print("  [Refreshed keep-awake state]")

    except KeyboardInterrupt:
        print("\n" + "=" * 60)
        print("Stopping keep-alive...")

    finally:
        # Re-enable sleep
        ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
        print("[OK] Screen timeout re-enabled")
        print("[OK] System sleep re-enabled")
        print("=" * 60)

if __name__ == "__main__":
    keep_screen_on()