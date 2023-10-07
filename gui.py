import tkinter as tk
from main import run_process  # Import the run_process function from your main.py

def start_process():
    run_process()
    status_label.config(text="Process completed")

# Create the main window
root = tk.Tk()
root.title("Pepsi Scheduler")

# Create a button to start the process
start_button = tk.Button(root, text="Start Process", command=start_process)
start_button.pack(pady=20)

# Create a label to display the status
status_label = tk.Label(root, text="Ready")
status_label.pack(pady=10)

# Run the GUI event loop
root.mainloop()
