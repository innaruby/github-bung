import tkinter as tk
from tkinter import messagebox
import threading
import os

def run_processing():
    selected_directory = r"U:\rev"
    if not os.path.exists(selected_directory):
        messagebox.showerror("Error", f"Directory not found: {selected_directory}")
        return

    try:
        print(f"📁 Processing directory: {selected_directory}")
        process_excel_files(selected_directory)
        post_processing_with_vlookup(selected_directory)
        final_sum_pass(selected_directory)

        for file in os.listdir(selected_directory):
            if file.lower().startswith("kostenstelle") or not file.endswith((".xlsx", ".xlsm")):
                continue
            file_path = os.path.join(selected_directory, file)
            wb = openpyxl.load_workbook(file_path)
            process_sachaufwand_links(wb, file_path)
            for sheet_name in wb.sheetnames:
                if sheet_name.lower() == 'sachaufwand':
                    ws = wb[sheet_name]
                    end_row = find_end_row(ws, sheet_name)
                    apply_final_sums(ws, end_row)
                    print(f"📘 Zwischensumme and Summe logic applied to 'Sachaufwand' in file: {file}")
            wb.save(file_path)
            print(f"💾 Final update (Sachaufwand) saved in file: {file}")

        apply_light_grey_fill_final(selected_directory)

        messagebox.showinfo("Success", "✅ All files processed successfully!")

    except Exception as e:
        print(f"❌ Error: {e}")
        messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def on_process1_click():
    threading.Thread(target=run_processing).start()  # Run in a thread to avoid freezing UI

def launch_gui():
    root = tk.Tk()
    root.title("Excel Processor")

    label = tk.Label(root, text="Click below to start processing Excel files from U:\\rev", font=("Arial", 12))
    label.pack(pady=10)

    process_btn = tk.Button(root, text="Process1", font=("Arial", 12, "bold"), bg="green", fg="white",
                            width=20, height=2, command=on_process1_click)
    process_btn.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    launch_gui()
