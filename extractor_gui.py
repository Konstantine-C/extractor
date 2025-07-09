import os
import zipfile
import io
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

used_folder_names = {}

def fix_windows_path(path):
    if os.name == 'nt':
        path = os.path.normpath(path)
        if not path.startswith('\\\\?\\'):
            path = '\\\\?\\' + path
    return path


def get_unique_folder_name(base_name):
    clean_name = base_name
    counter = 1
    while clean_name in used_folder_names:
        clean_name = f"{base_name}_{counter}"
        counter += 1
    used_folder_names[clean_name] = True
    return clean_name


def clean_zip_name(zip_filename):
    import re
    name = os.path.splitext(zip_filename)[0]
    name = name.upper()
    name = re.sub(r'^(FW|SS|HO|SP)\d{2}_', '', name)
    name = re.sub(r'(_|\-)(1X1|RGB|CMYK|DIGITAL|PRINT|SOCIAL|EMAIL|HD|LOWRES|HR|LR)$', '', name)
    name = re.sub(r'[_\-]+', '_', name)
    parts = name.split('_')
    if len(parts) > 6:
        name = '_'.join(parts[:6])
    name = name.strip('_')
    return get_unique_folder_name(name)


def extract_from_zip(zip_file, output_root, zip_folder_name, allowed_exts):
    for member in zip_file.infolist():
        if member.is_dir():
            continue
        filename = member.filename
        _, ext = os.path.splitext(filename.lower())

        if ext in allowed_exts:
            ext_folder = ext.lstrip('.')
            short_path = os.path.basename(filename)
            final_dir = os.path.join(output_root, zip_folder_name, ext_folder)
            os.makedirs(final_dir, exist_ok=True)
            target_path = fix_windows_path(os.path.join(final_dir, short_path))

            try:
                with zip_file.open(member) as source, open(target_path, "wb") as target:
                    target.write(source.read())
            except Exception as e:
                print(f"Failed to extract {filename}: {e}")

        elif ext == '.zip':
            try:
                nested_data = zip_file.read(member)
                with zipfile.ZipFile(io.BytesIO(nested_data)) as nested_zip:
                    extract_from_zip(nested_zip, output_root, zip_folder_name, allowed_exts)
            except Exception as e:
                print(f"Failed to open nested zip {filename}: {e}")


def extract_selected_files(input_dir, output_dir, allowed_exts):
    if not os.path.exists(input_dir) or not os.path.exists(output_dir):
        messagebox.showerror("Error", "Please select valid folders.")
        return

    zip_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".zip")]
    if not zip_files:
        messagebox.showwarning("No ZIPs", "No ZIP files found in the selected input folder.")
        return

    for file in zip_files:
        zip_path = os.path.join(input_dir, file)
        try:
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_folder_name = clean_zip_name(file)
                extract_from_zip(zip_ref, output_dir, zip_folder_name, allowed_exts)
        except zipfile.BadZipFile:
            print(f"Bad ZIP: {zip_path}")

    messagebox.showinfo("Done", "Extraction complete!")


# --- GUI START ---
def launch_gui():
    import threading
    from tkinter import ttk

    root = tk.Tk()
    root.title("Smart ZIP Extractor")
    root.geometry("500x480")
    root.resizable(False, False)

    allowed_exts = set()
    zip_files = []

    def browse_input():
        path = filedialog.askdirectory()
        if path:
            input_var.set(path)

    def browse_output():
        path = filedialog.askdirectory()
        if path:
            output_var.set(path)

    def run_extraction_thread():
        t = threading.Thread(target=run_extraction)
        t.start()

    def run_extraction():
        nonlocal zip_files

        selected_exts = [ext for ext, var in ext_vars.items() if var.get()]
        if not selected_exts:
            messagebox.showwarning("No extensions", "Select at least one file extension to extract.")
            return

        input_dir = input_var.get()
        output_dir = output_var.get()

        if not os.path.exists(input_dir) or not os.path.exists(output_dir):
            messagebox.showerror("Error", "Please select valid input and output folders.")
            return

        zip_files = [f for f in os.listdir(input_dir) if f.lower().endswith(".zip")]
        if not zip_files:
            messagebox.showinfo("No ZIPs", "No ZIP files found in selected input folder.")
            return

        progress_label.config(text="Starting extraction...")
        progress_bar["maximum"] = len(zip_files)
        progress_bar["value"] = 0
        root.update_idletasks()

        selected_exts_set = set(f".{e}" for e in selected_exts)

        for idx, file in enumerate(zip_files):
            progress_label.config(text=f"Extracting: {file}")
            root.update_idletasks()

            zip_path = os.path.join(input_dir, file)
            try:
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_folder_name = clean_zip_name(file)
                    extract_from_zip(zip_ref, output_dir, zip_folder_name, selected_exts_set)
            except zipfile.BadZipFile:
                print(f"Bad ZIP: {zip_path}")

            progress_bar["value"] = idx + 1
            root.update_idletasks()

        progress_label.config(text="✅ Extraction complete!")
        messagebox.showinfo("Done", "Extraction completed successfully!")

    # Input/output selectors
    input_var = tk.StringVar()
    output_var = tk.StringVar()

    tk.Label(root, text="📁 Input ZIP Folder:").pack(pady=(10, 0))
    tk.Entry(root, textvariable=input_var, width=60).pack()
    tk.Button(root, text="Browse", command=browse_input).pack(pady=(0, 10))

    tk.Label(root, text="📂 Output Extraction Folder:").pack()
    tk.Entry(root, textvariable=output_var, width=60).pack()
    tk.Button(root, text="Browse", command=browse_output).pack(pady=(0, 10))

    # Extension checkboxes
    tk.Label(root, text="✅ File Types to Extract:").pack(pady=(10, 5))
    ext_frame = tk.Frame(root)
    ext_frame.pack()

    common_exts = ['jpg', 'jpeg', 'png', 'tiff', 'tif', 'pdf', 'idml', 'indd', 'otf', 'lst']
    ext_vars = {}
    for i, ext in enumerate(common_exts):
        var = tk.BooleanVar(value=(ext in ['jpg', 'jpeg', 'pdf']))  # default selected
        chk = tk.Checkbutton(ext_frame, text=ext.upper(), variable=var)
        chk.grid(row=i // 5, column=i % 5, padx=10, pady=5, sticky="w")
        ext_vars[ext] = var

    # Progress bar + status
    progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
    progress_bar.pack(pady=(20, 5))

    progress_label = tk.Label(root, text="Waiting for input...")
    progress_label.pack()

    # Extract button
    tk.Button(root, text="🚀 Extract Files", command=run_extraction_thread, bg="#0078D7", fg="white", font=("Arial", 12)).pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
	launch_gui()import os