from pynput import mouse, keyboard
import time
import random

button_to_trigger = mouse.Button.x2  # Assuming the 4th mouse button is Button.x2
stop_key = keyboard.Key.esc  # Change this key if needed

is_script_running = False

def on_click(x, y, button, pressed):
    global is_script_running
    if button == button_to_trigger and pressed and not is_script_running:
        is_script_running = True
        # Your script logic here

        # Simulate keypress for 'v'
        keyboard.Controller().press('v')
        time.sleep(random.uniform(0.096, 0.1))
        keyboard.Controller().release('v')

        time.sleep(0.005)

        # Simulate keypress for '2'
        keyboard.Controller().press('2')
        keyboard.Controller().release('2')

        # Simulate keypress for '1'
        keyboard.Controller().press('1')
        time.sleep(random.uniform(0.096, 0.1))
        keyboard.Controller().release('1')

        # End of script logic

        is_script_running = False

# Start the listener
with mouse.Listener(on_click=on_click) as listener:
    try:
        listener.join()
    except KeyboardInterrupt:
        # Handle KeyboardInterrupt (Ctrl+C) to gracefully exit the script
        pass
