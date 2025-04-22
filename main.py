import warnings
# ─ Ignore harmless cmap warnings from fpdf ─
warnings.filterwarnings(
    "ignore",
    message="cmap value too big/small:.*",
    module="fpdf.ttfonts"
)

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import platform
import pandas as pd
from fpdf import FPDF
import re
import threading
import subprocess

# Try to import the TkinterDnD wrapper for drag‑and‑drop
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

stop_requested = False  # signal to stop processing


def get_default_font_path():
    system = platform.system()
    if system == "Windows":
        return r"C:\Windows\Fonts\arial.ttf"
    elif system == "Linux":
        return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    elif system == "Darwin":
        return "/Library/Fonts/Arial.ttf"
    else:
        raise RuntimeError("Unsupported OS or cannot determine font path.")


FONT_PATH = get_default_font_path()


def sanitize_filename(filename: str) -> str:
    invalid_chars = r'\/:*?"<>|'
    return re.sub(f"[{re.escape(invalid_chars)}]", "_", filename)


class PDFGenerator(FPDF):
    def __init__(self, order_number, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.order_number = order_number
        try:
            self.add_font("DefaultFont", "", FONT_PATH, uni=True)
            self.add_font("DefaultFont", "B", FONT_PATH, uni=True)
            self.set_auto_page_break(False)
        except Exception as e:
            messagebox.showerror(
                "Font Error",
                f"Failed to load font from {FONT_PATH}.\nError: {e}"
            )
            raise

    def header(self):
        self.set_font("DefaultFont", "B", 12)
        title = f"Order Number {self.order_number}"
        self.cell(0, 10, title, align="C", ln=1)
        self.ln(3)


def print_column_fields(pdf, fields, data, start_x, start_y, col_width):
    current_y = start_y
    line_height = 6
    for label, key in fields:
        pdf.set_xy(start_x, current_y)
        pdf.set_font("DefaultFont", "B", 14)
        pdf.multi_cell(col_width, line_height, f"{label}:", align="L")
        current_y = pdf.get_y()
        pdf.set_xy(start_x, current_y)
        pdf.set_font("DefaultFont", "", 10)
        pdf.multi_cell(col_width, line_height, str(data.get(key, "")).strip(), align="L")
        current_y = pdf.get_y() + 4
    return current_y


def generate_pdf(data, output_path, order_number):
    try:
        pdf = PDFGenerator(order_number, format="Letter", orientation="P")
        pdf.add_page()
        full = f"{data.get('First Name (Billing)', '').strip()} {data.get('Last Name (Billing)', '').strip()}".strip()
        data["Full Name"] = full
        left_fields = [
            ("Name", "Full Name"),
            ("Email", "Email (Billing)"),
            ("Phone", "Phone (Billing)"),
            ("Shipping Method Title", "Shipping Method Title"),
            ("City", "City (Billing)"),
            ("Address 1&2", "Address 1&2 (Billing)"),
        ]
        right_fields = [
            ("Item Name", "Item Name"),
            ("SKU", "SKU"),
            ("Quantity", "Quantity (- Refund)"),
            ("Item Cost", "Item Cost"),
            ("Order Total Amount", "Order Total Amount"),
            ("Customer Note", "Customer Note"),
        ]
        available_w = pdf.w - pdf.l_margin - pdf.r_margin
        col_w = available_w / 2
        start_y = 25
        print_column_fields(pdf, left_fields, data, pdf.l_margin, start_y, col_w)
        print_column_fields(pdf, right_fields, data, pdf.l_margin + col_w, start_y, col_w)
        pdf.output(output_path)
    except Exception as e:
        messagebox.showerror(
            "PDF Generation Error",
            f"Failed for order {order_number}.\nDetails: {e}"
        )
        raise


def validate_csv(df):
    required = [
        "Order Number",
        "First Name (Billing)",
        "Last Name (Billing)",
        "Email (Billing)",
        "Phone (Billing)",
        "Item Name",
        "SKU",
        "Quantity (- Refund)",
        "Item Cost",
        "Order Total Amount",
    ]
    return [c for c in required if c not in df.columns]


def process_csv(csv_path, progress_label, progress_bar, open_folder_button):
    global stop_requested
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        messagebox.showerror("CSV Error", f"Could not read CSV:\n{e}")
        generate_button.config(state="normal")
        stop_button.config(state="disabled")
        return

    missing = validate_csv(df)
    if missing:
        messagebox.showerror("CSV Error", f"Missing columns: {', '.join(missing)}")
        generate_button.config(state="normal")
        stop_button.config(state="disabled")
        return

    out_dir = os.path.join(os.path.dirname(csv_path), "PDF Orders")
    os.makedirs(out_dir, exist_ok=True)

    total = len(df)
    progress_bar["maximum"] = total
    progress_bar["value"] = 0

    for idx, row in df.iterrows():
        if stop_requested:
            progress_label.config(text="Canceled by user.")
            messagebox.showinfo("Canceled", "Generation stopped.")
            break

        progress_label.config(text=f"Processing {idx+1}/{total}…")
        progress_bar["value"] = idx + 1

        filename = sanitize_filename(
            f"{row['Order Number']} - {row['First Name (Billing)']} {row['Last Name (Billing)']}.pdf"
        )
        out_path = os.path.join(out_dir, filename)
        try:
            generate_pdf(row.to_dict(), out_path, row["Order Number"])
        except Exception:
            pass

    else:
        progress_label.config(text="All done!")
        messagebox.showinfo("Success", "PDFs generated.")
        open_folder_button.config(state="normal")

    stop_requested = False
    generate_button.config(state="normal")
    stop_button.config(state="disabled")


def start_process_csv(csv_path, progress_label, progress_bar, open_folder_button):
    def worker():
        process_csv(csv_path, progress_label, progress_bar, open_folder_button)

    generate_button.config(state="disabled")
    stop_button.config(state="normal")
    open_folder_button.config(state="disabled")
    threading.Thread(target=worker, daemon=True).start()


def browse_file():
    file = filedialog.askopenfilename(
        parent=root, title="Select CSV", filetypes=[("CSV Files", "*.csv")]
    )
    if file:
        csv_path_var.set(file)


def generate():
    global stop_requested
    stop_requested = False
    path = csv_path_var.get()
    if not path:
        messagebox.showerror("Input Error", "Select a CSV first.")
        return
    start_process_csv(path, progress_label, progress_bar, open_folder_button)


def stop_process():
    global stop_requested
    stop_requested = True
    progress_label.config(text="Canceling…")


def open_output_folder():
    base = os.path.dirname(csv_path_var.get())
    folder = os.path.join(base, "PDF Orders")
    try:
        if platform.system() == "Windows":
            os.startfile(folder)
        elif platform.system() == "Darwin":
            subprocess.call(["open", folder])
        else:
            subprocess.call(["xdg-open", folder])
    except Exception as e:
        messagebox.showwarning("Open Folder Failed", f"{folder}\n{e}")


def show_instructions():
    messagebox.showinfo(
        "How to Use Order PDF Generator",
        "1. Prepare a CSV file with these headers:\n"
        "   • Order Number, First Name (Billing), Last Name (Billing)\n"
        "   • Email (Billing), Phone (Billing)\n"
        "   • Item Name, SKU, Quantity (- Refund), Item Cost\n"
        "   • Order Total Amount (optional: Customer Note)\n\n"
        "2. Launch this application.\n"
        "3. Drag & drop your CSV onto the entry field—or click Browse to select it.\n"
        "4. Click Generate PDFs.\n"
        "   • Progress will show below.\n"
        "   • Click Stop to cancel at any time.\n\n"
        "5. When complete, click Open Output Folder to view your PDFs.\n"
        "6. Use Quit to exit the program."
    )


def show_about():
    messagebox.showinfo(
        "About Order PDF Generator",
        "Order PDF Generator\n"
        "Version 1.0\n\n"
        "Creates two‑column PDFs from your CSV orders.\n"
        "© 2025"
    )


def quit_app():
    root.quit()


# ─── GUI SETUP ────────────────────────────────────────────
if DND_AVAILABLE:
    root = TkinterDnD.Tk()
else:
    root = tk.Tk()

root.title("Order PDF Generator")
root.geometry("800x340")

# Menu bar
menubar = tk.Menu(root)
helpmenu = tk.Menu(menubar, tearoff=0)
helpmenu.add_command(label="Instructions", command=show_instructions)
helpmenu.add_command(label="About", command=show_about)
menubar.add_cascade(label="Help", menu=helpmenu)
root.config(menu=menubar)

csv_path_var = tk.StringVar()
main = tk.Frame(root, padx=10, pady=10)
main.pack(fill="both", expand=True)

tk.Label(main, text="Select your CSV file:")\
    .grid(row=0, column=0, sticky="e", padx=5, pady=5)

csv_entry = tk.Entry(main, textvariable=csv_path_var, width=50)
csv_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

if DND_AVAILABLE:
    csv_entry.drop_target_register(DND_FILES)
    csv_entry.dnd_bind(
        '<<Drop>>',
        lambda e: csv_path_var.set(e.data.strip('{}'))
    )

tk.Button(main, text="Browse", command=browse_file)\
    .grid(row=0, column=2, padx=5, pady=5)

generate_button = tk.Button(main, text="Generate PDFs", command=generate)
generate_button.grid(row=1, column=0, pady=10, sticky="e")

stop_button = tk.Button(main, text="Stop", command=stop_process, state="disabled")
stop_button.grid(row=1, column=1, pady=10, sticky="w")

progress_label = tk.Label(main, text="")
progress_label.grid(row=2, column=0, columnspan=3, pady=5)

progress_bar = ttk.Progressbar(main, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=3, column=0, columnspan=3, pady=5)

open_folder_button = tk.Button(
    main, text="Open Output Folder", command=open_output_folder, state="disabled"
)
open_folder_button.grid(row=4, column=0, columnspan=3, pady=5)

tk.Button(main, text="Quit", command=quit_app)\
    .grid(row=5, column=0, columnspan=3, pady=5)

root.mainloop()
