from pynput import mouse, keyboard
import time
import random

def on_click(x, y, button, pressed):
    print(f"Button {button} {'pressed' if pressed else 'released'}")
    # Add your script logic here based on the printed button information

# Start the listener
with mouse.Listener(on_click=on_click) as listener:
    listener.join()
