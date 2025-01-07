import pyautogui
import time

print("Move your mouse to the desired position and wait...")
time.sleep(2)  # Gives you 2 seconds to position your mouse

x, y = pyautogui.position()
print(f"Mouse position: ({x}, {y})")