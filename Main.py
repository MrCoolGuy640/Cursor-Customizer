# Imports
import win32con
import win32api
import win32gui
import ctypes
import os
import json
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import atexit

# Dictionary containing the numbers win32con accepts for cursor modifications, along with text holding cursor type!
cursor_numbers = {
    32512: "ArrowCursor",
    32651: "HelpCursor",
    32514: "BackgroundWaitCursor",
    32650: "LoadingCursor",
    32515: "PrecisionSelectCursor",
    32513: "TextCursor",
    32649: "HandCursor",
    32648: "UnavailableCursor",
    32644: "HorizontalResizeCursor",
    32645: "VerticalResizeCursor",
    32642: "DiagonalResizeCursor1",
    32643: "DiagonalResizeCursor2",
    32646: "MoveCursor"
}

# Load Config.json
def load_cursor_config():
    if os.path.exists("Config.json"):
        with open("Config.json", "r") as config_file:
            data = json.load(config_file)
            cursor_types = {
                win32con.IDC_ARROW: str(data.get("ArrowCursor", "")),
                win32con.IDC_HELP: str(data.get("HelpCursor", "")),
                win32con.IDC_WAIT: str(data.get("BackgroundWaitCursor", "")),
                win32con.IDC_APPSTARTING: str(data.get("LoadingCursor", "")),
                win32con.IDC_CROSS: str(data.get("PrecisionSelectCursor", "")),
                win32con.IDC_IBEAM: str(data.get("TextCursor", "")),
                win32con.IDC_HAND: str(data.get("HandCursor", "")),
                win32con.IDC_NO: str(data.get("UnavailableCursor", "")),
                win32con.IDC_SIZEWE: str(data.get("HorizontalResizeCursor", "")),
                win32con.IDC_SIZENS: str(data.get("VerticalResizeCursor", "")),
                win32con.IDC_SIZENWSE: str(data.get("DiagonalResizeCursor1", "")),
                win32con.IDC_SIZENESW: str(data.get("DiagonalResizeCursor2", "")),
                win32con.IDC_SIZEALL: str(data.get("MoveCursor", "")),
            }

    return cursor_types


cursor_types = load_cursor_config()
original_cursors = {}

# Save original system cursor icons before changing it so it can be restored later.
for cursor_type, cursor_file in cursor_types.items():
    original_cursors[cursor_type] = ctypes.windll.user32.CopyImage(
        win32gui.LoadImage(0, cursor_type, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_SHARED),
        win32con.IMAGE_CURSOR, 0, 0, win32con.LR_COPYFROMRESOURCE
    )

# OLD SAVING (Was 1 line manually added per cursor)
#cursor = win32gui.LoadImage(0, win32con.IDC_ARROW, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_SHARED)
#save_system_cursor = ctypes.windll.user32.CopyImage(cursor, win32con.IMAGE_CURSOR, 0, 0, win32con.LR_COPYFROMRESOURCE)

# Function to restore the original system cursors
def restore_cursors():
    for cursor_type, cursor in original_cursors.items():
        ctypes.windll.user32.SetSystemCursor(cursor, cursor_type)
        ctypes.windll.user32.DestroyCursor(cursor)

# Make sure original system cursers restored when code is stopped
atexit.register(restore_cursors)

# Function to modify system cursor based on the given cursor type
def change_cursor(cursor_type, cursor_file):
    print(cursor_type)
    print(cursor_file)
    for number, cursor in cursor_numbers.items():
        if cursor == cursor_type:
            cursor_type = number
            break
        
    cursor = cursor_types.get(cursor_type)
    if cursor:
        cursor = win32gui.LoadImage(0, str(cursor_file), win32con.IMAGE_CURSOR, 0, 0, win32con.LR_LOADFROMFILE)
        ctypes.windll.user32.SetSystemCursor(cursor, cursor_type)

# Saving new data to Config.json
def save_cursor_config(cursor_type, cursor_file):
    if os.path.exists("Config.json"):
        with open("Config.json", "r") as config_file:
            data = json.load(config_file)
    
    data[str(cursor_type)] = str(cursor_file)
    with open("Config.json", "w") as config_file:
        json.dump(data, config_file)

# Display cursor previews in app
def show_cursor_preview(cursor_file, cursor_preview):
    img = Image.open(cursor_file)
    img.thumbnail((32, 32), Image.LANCZOS)
    img = ImageTk.PhotoImage(img)

    # Clear existing content
    cursor_preview.delete("all")

    # Create a dark gray background
    cursor_preview.create_rectangle(0, 0, 40, 40, fill="dark gray")

    # Calculate center and show the cursor preview
    x = (40 - img.width() / 2) / 1.5
    y = (40 - img.height() / 2) / 2
    cursor_preview.create_image(x, y, anchor=tk.NW, image=img)

    cursor_preview.image = img

# Function to handle selecting cursor file and changing the cursor
def select_cursor_file(cursor_type):
    cursor_file = filedialog.askopenfilename(initialdir=os.path.join(os.path.dirname(__file__), "Cursors"), title=f"Select Cursor for {cursor_type}", filetypes=(("Cursor files", "*.cur"),("Animation Files", "*.ani")))
    if cursor_file:
        change_cursor(cursor_type, cursor_file) #THIS
        save_cursor_config(cursor_type, cursor_file) #THIS
        cursor_name = os.path.basename(cursor_file)
        cursor_labels[cursor_type].config(text=f"Selected Cursor: {cursor_name}")
        show_cursor_preview(cursor_file, cursor_previews[cursor_type])

# Create Tkinter window for cursor selection
root = tk.Tk()
root.title("Cursor Selection")

# Create preview panes, labels, and buttons for each cursor type
columns = 3
rows = -(-len(cursor_types) // columns)  # ceiling division to calculate the number of rows
cursor_labels = {}
cursor_previews = {}
cursor_select_buttons = {}

i = 0

#print(cursor_types)

for cursor_type, cursor_file in cursor_types.items():
    #print(cursor_type)
    if cursor_type in cursor_numbers.keys():
        cursor_type = cursor_numbers[cursor_type]

    frame = tk.Frame(root)
    frame.grid(row=i // columns, column=i % columns)

    cursor_labels[cursor_type] = tk.Label(frame, text=f"Selected Cursor: None")
    cursor_labels[cursor_type].pack()

    cursor_preview = tk.Canvas(frame, width=40, height=40, bg="light gray")
    cursor_preview.pack()
    cursor_previews[cursor_type] = cursor_preview

    cursor_select_buttons[cursor_type] = tk.Button(frame, text=f"Select Cursor for {cursor_type}", command=lambda ct=cursor_type: select_cursor_file(ct))
    cursor_select_buttons[cursor_type].pack()

    i += 1

# Load previously selected cursors and show their previews
for cursor_type, cursor_file in cursor_types.items():
    cursor_name = cursor_numbers[cursor_type]  # Fetch the cursor text name from the numeric identifier
    if cursor_file:
        change_cursor(cursor_type, cursor_file)
        cursor_labels[cursor_name].config(text=f"Selected Cursor: {os.path.basename(cursor_file)}")
        show_cursor_preview(cursor_file, cursor_previews[cursor_name])

root.mainloop()
