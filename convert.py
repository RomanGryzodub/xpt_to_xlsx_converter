import tkinter as tk
from tkinter import filedialog, ttk
import pandas as pd
import pyreadstat

def reset_progress_bar():
    progress['value'] = 0
    progress_label.config(text="Progress: 0%")  # Reset progress text
    status_label.config(text="")  # Clear status label before starting a new operation
    progress.pack(pady=10)

def update_progress(step, total_steps):
    progress['value'] = step
    percentage = int((step / total_steps) * 100)
    progress_label.config(text=f"Progress: {percentage}%")  # Update text
    root.update_idletasks()

def select_files():
    file_paths = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel files", "*.xlsx")])
    if file_paths:
        reset_progress_bar()
        process_files(file_paths)

def process_files(file_paths):
    total_steps = len(file_paths) * 3
    progress['maximum'] = total_steps
    current_step = 0
    status_messages = []

    for file_path in file_paths:
        try:
            data = pd.read_excel(file_path)
            data.columns = [f'COL_{i+1}' for i in range(len(data.columns))]
            xpt_file_path = file_path.rsplit('.', 1)[0] + '.xpt'
            pyreadstat.write_xport(data, xpt_file_path)
            status_messages.append(f"Converted: {xpt_file_path}")
        except Exception as e:
            status_messages.append(f"Error with {file_path}: {e}")
        finally:
            current_step += 3
            update_progress(current_step, total_steps)

    status_label.config(text="\n".join(status_messages))
    progress_label.config(text="Progress: Completed")

def select_xpt_files():
    file_paths = filedialog.askopenfilenames(title="Select XPT Files", filetypes=[("XPT files", "*.xpt")])
    if file_paths:
        reset_progress_bar()
        convert_xpt_to_xlsx(file_paths)

def convert_xpt_to_xlsx(file_paths):
    total_steps = len(file_paths)
    progress['maximum'] = total_steps
    current_step = 0
    status_messages = []

    for file_path in file_paths:
        try:
            data, meta = pyreadstat.read_xport(file_path)
            xlsx_file_path = file_path.rsplit('.', 1)[0] + '.xlsx'
            data.to_excel(xlsx_file_path, index=False)
            status_messages.append(f"Converted: {xlsx_file_path}")
        except Exception as e:
            status_messages.append(f"Error with {file_path}: {e}")
        finally:
            current_step += 1
            update_progress(current_step, total_steps)

    status_label.config(text="\n".join(status_messages))
    progress_label.config(text="Progress: Completed")

# Set up the tkinter window
root = tk.Tk()
root.title("Excel/XPT Converter by Roman Gryzodub")
root.iconbitmap('icon.ico')  # Use raw string for paths on Windows

root.minsize(500, 300)

select_excel_button = tk.Button(root, text="Select Excel Files to Convert to XPT", command=select_files)
select_excel_button.pack(pady=10)

select_xpt_button = tk.Button(root, text="Select XPT Files to Convert to Excel", command=select_xpt_files)
select_xpt_button.pack(pady=10)

progress = ttk.Progressbar(root, orient=tk.HORIZONTAL, length=450, mode='determinate')
progress.pack(pady=10)

progress_label = tk.Label(root, text="Progress: 0%")
progress_label.pack(side=tk.BOTTOM, fill=tk.X)

status_label = tk.Label(root, text="", justify=tk.LEFT)
status_label.pack(pady=10, side=tk.BOTTOM)

root.mainloop()
