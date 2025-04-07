import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import platform
import pandas as pd
from fpdf import FPDF
import re
import threading
import subprocess

# Global flag used to signal the worker thread to stop processing.
stop_requested = False


def get_default_font_path():
    """Return a default Unicode TTF font path based on the OS."""
    system = platform.system()
    if system == "Windows":
        return r"C:\Windows\Fonts\arial.ttf"
    elif system == "Linux":
        return "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
    elif system == "Darwin":
        return "/Library/Fonts/Arial.ttf"
    else:
        raise RuntimeError("Unsupported OS or cannot determine default font path.")


FONT_PATH = get_default_font_path()


def sanitize_filename(filename: str) -> str:
    """Replace invalid filename characters with underscores."""
    invalid_chars = r'\/:*?"<>|'
    return re.sub(f"[{re.escape(invalid_chars)}]", "_", filename)


class PDFGenerator(FPDF):
    """Custom FPDF class that accepts an 'order_number' to customize the header text."""

    def __init__(self, order_number, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.order_number = order_number  # We'll use this in the header

        try:
            # Attempt to register default font (regular + bold).
            self.add_font("DefaultFont", "", FONT_PATH, uni=True)
            self.add_font("DefaultFont", "B", FONT_PATH, uni=True)
            # Turn off auto-page-break so everything can stay on one page if it's not too large.
            self.set_auto_page_break(False)
        except Exception as e:
            messagebox.showerror(
                "Font Error",
                f"Failed to load font from {FONT_PATH}.\n"
                f"Please ensure this font file is installed.\nError: {e}"
            )
            raise

    def header(self):
        """
        Override the default header to display:
        "Order Number <order_number>"
        """
        self.set_font("DefaultFont", "B", 12)
        title_text = f"Order Number {self.order_number}"
        self.cell(0, 10, title_text, align="C", new_x="LMARGIN", new_y="NEXT")
        self.ln(3)


def print_column_fields(pdf, fields, data, start_x, start_y, col_width):
    """
    Prints each field as a label (14pt bold) followed by its value (10pt),
    wrapping text as needed. Returns the final Y position after printing all fields.
    """
    current_y = start_y
    line_height = 6  # line spacing

    for label, key in fields:
        # Label in 14pt bold
        pdf.set_xy(start_x, current_y)
        pdf.set_font("DefaultFont", "B", 14)
        pdf.multi_cell(col_width, line_height, f"{label}:", align='L')

        current_y = pdf.get_y()

        # Value in 10pt
        value = str(data.get(key, "")).strip()
        pdf.set_xy(start_x, current_y)
        pdf.set_font("DefaultFont", "", 10)
        pdf.multi_cell(col_width, line_height, value, align='L')

        current_y = pdf.get_y() + 4

    return current_y


def generate_pdf(data, output_path, order_number):
    """
    Generates a single-page PDF for one order. The 'order_number' is used in the header.

    New layout:
      Left column:
        - Name (combined)
        - Email
        - Phone
        - Shipping Method Title
        - City
        - Address 1&2

      Right column:
        - Item Name
        - SKU
        - Quantity
        - Item Cost
        - [Potential more item lines if desired]
        - Last line: Order Total Amount
    """
    try:
        pdf = PDFGenerator(order_number=order_number, format="Letter", orientation="P")
        pdf.add_page()

        # We'll store "Full Name" in data to combine First/Last for convenience.
        # The CSV still has "First Name (Billing)" and "Last Name (Billing)" columns,
        # but we want to display them together in one line: "Name".
        full_name = f"{data.get('First Name (Billing)', '').strip()} {data.get('Last Name (Billing)', '').strip()}"
        data["Full Name"] = full_name.strip()

        # Define fields in each column
        # Left column
        fields_col1 = [
            ("Name", "Full Name"),
            ("Email", "Email (Billing)"),
            ("Phone", "Phone (Billing)"),
            ("Shipping Method Title", "Shipping Method Title"),
            ("City", "City (Billing)"),
            ("Address 1&2", "Address 1&2 (Billing)"),
        ]

        # Right column (and we put "Order Total Amount" last)
        fields_col2 = [
            ("Item Name", "Item Name"),
            ("SKU", "SKU"),
            ("Quantity", "Quantity (- Refund)"),
            ("Item Cost", "Item Cost"),
            ("Order Total Amount", "Order Total Amount"),
        ]

        page_width = pdf.w - pdf.l_margin - pdf.r_margin
        col_width = page_width / 2
        start_y = 25

        # Left column
        col1_x = pdf.l_margin
        print_column_fields(pdf, fields_col1, data, col1_x, start_y, col_width)

        # Right column
        col2_x = pdf.l_margin + col_width
        print_column_fields(pdf, fields_col2, data, col2_x, start_y, col_width)

        pdf.output(output_path)
    except Exception as e:
        messagebox.showerror(
            "PDF Generation Error",
            f"Failed to generate PDF for order.\nDetails: {e}"
        )
        raise


def validate_csv(df):
    """
    Checks for required columns and returns a list of any missing columns.
    """
    # Example: if your script absolutely needs these columns to run.
    required_columns = [
        "Order Number",
        "First Name (Billing)",
        "Last Name (Billing)",
        "Email (Billing)",
        "Phone (Billing)",
        "Item Name",
        "SKU",
        "Quantity (- Refund)",
        "Item Cost",
        "Order Total Amount"
        # Add more if they're strictly required for your generation logic.
    ]
    missing = [col for col in required_columns if col not in df.columns]
    return missing


def process_csv(csv_path, progress_label, progress_bar, open_folder_btn):
    """
    Reads the CSV, checks for missing columns, then generates one PDF per row.
    Allows user to cancel via the global 'stop_requested' flag.
    """
    global stop_requested

    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        messagebox.showerror(
            "CSV Error",
            "Failed to read the CSV file.\n\n"
            "Please ensure it's not open in another program and is valid CSV.\n\n"
            f"Details: {e}"
        )
        # Re-enable UI
        generate_button.config(state="normal")
        stop_button.config(state="disabled")
        return

    # Check for missing columns
    missing_columns = validate_csv(df)
    if missing_columns:
        messagebox.showerror(
            "CSV Error",
            f"The CSV file is missing required columns: {', '.join(missing_columns)}.\n"
            "Please correct the file and try again."
        )
        generate_button.config(state="normal")
        stop_button.config(state="disabled")
        return

    base_dir = os.path.dirname(csv_path)
    output_folder = os.path.join(base_dir, "PDF Orders")
    os.makedirs(output_folder, exist_ok=True)

    total_rows = len(df)
    progress_bar["value"] = 0
    progress_bar["maximum"] = total_rows

    for index, row in df.iterrows():
        if stop_requested:
            progress_label.config(text="Operation canceled by user.")
            messagebox.showinfo("Canceled", "Order generation was canceled.")
            break

        progress_label.config(text=f"Processing record {index + 1}/{total_rows}...")
        progress_bar["value"] = index + 1

        first_name = str(row.get("First Name (Billing)", "")).strip()
        last_name = str(row.get("Last Name (Billing)", "")).strip()
        order_number = str(row.get("Order Number", "")).strip()

        raw_filename = f"{first_name} {last_name} {order_number}.pdf"
        safe_filename = sanitize_filename(raw_filename)
        output_path = os.path.join(output_folder, safe_filename)

        data = row.to_dict()

        try:
            generate_pdf(data, output_path, order_number)
        except Exception as e:
            messagebox.showerror(
                "Order Processing Error",
                f"Failed to generate PDF for order number "
                f"{row.get('Order Number', 'Unknown')}:\n\n{e}"
            )

    else:
        # If we didn't break (meaning stop wasn't triggered)
        progress_label.config(text="All PDFs generated successfully.")
        messagebox.showinfo("Success", "PDF generation complete.")
        open_folder_btn.config(state="normal")

    # Reset stop state & re-enable UI
    stop_requested = False
    generate_button.config(state="normal")
    stop_button.config(state="disabled")


def start_process_csv(csv_path, progress_label, progress_bar, open_folder_btn):
    """
    Launches process_csv in a background thread, so the UI doesn't freeze
    and the user can press 'Stop' to cancel mid-way.
    """

    def worker():
        process_csv(csv_path, progress_label, progress_bar, open_folder_btn)

    generate_button.config(state="disabled")
    stop_button.config(state="normal")
    open_folder_btn.config(state="disabled")

    t = threading.Thread(target=worker, daemon=True)
    t.start()


def browse_file():
    """Open a file dialog to let the user pick the CSV file."""
    file_path = filedialog.askopenfilename(
        parent=root,
        title="Select CSV File",
        filetypes=[("CSV Files", "*.csv")]
    )
    if file_path:
        csv_path_var.set(file_path)


def generate():
    """Starts the CSV processing in a background thread (after validating the CSV path)."""
    global stop_requested
    stop_requested = False

    csv_path = csv_path_var.get()
    if not csv_path:
        messagebox.showerror("Input Error", "Please select a CSV file first.")
        return

    start_process_csv(csv_path, progress_label, progress_bar, open_folder_button)


def stop_process():
    """Sets the global flag that signals the worker thread to stop."""
    global stop_requested
    stop_requested = True
    progress_label.config(text="Canceling...")


def open_output_folder():
    """Open 'PDF Orders' folder in the file manager."""
    csv_path = csv_path_var.get()
    if not csv_path:
        return

    base_dir = os.path.dirname(csv_path)
    output_folder = os.path.join(base_dir, "PDF Orders")

    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(output_folder)
        elif system == "Darwin":
            subprocess.call(["open", output_folder])
        else:
            subprocess.call(["xdg-open", output_folder])
    except Exception as e:
        messagebox.showwarning(
            "Open Folder Failed",
            f"Could not open folder automatically.\nPath: {output_folder}\nError: {e}"
        )


def quit_app():
    """Close the application."""
    root.quit()


# ---------------
#   GUI SETUP
# ---------------
root = tk.Tk()
root.title("Order PDF Generator")
root.geometry("700x300")

csv_path_var = tk.StringVar()

main_frame = tk.Frame(root, padx=10, pady=10)
main_frame.pack(fill="both", expand=True)

# Prompt label for file
prompt_label = tk.Label(main_frame, text="Select your CSV file:")
prompt_label.grid(row=0, column=0, sticky="e", padx=5, pady=5)

# Entry & Browse
csv_entry = tk.Entry(main_frame, textvariable=csv_path_var, width=50)
csv_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

browse_button = tk.Button(main_frame, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=5, pady=5)

# Generate & Stop
generate_button = tk.Button(main_frame, text="Generate PDFs", command=generate)
generate_button.grid(row=1, column=2, pady=10, sticky="e")

stop_button = tk.Button(main_frame, text="Stop", command=stop_process)
stop_button.grid(row=4, column=1, pady=10, sticky="w")
stop_button.config(state="disabled")

# Progress label
progress_label = tk.Label(main_frame, text="")
progress_label.grid(row=2, column=0, columnspan=3, pady=5)

# Progress bar
progress_bar = ttk.Progressbar(main_frame, orient="horizontal", length=400, mode="determinate")
progress_bar.grid(row=3, column=0, columnspan=3, pady=5)

# Open folder button (disabled until success)
open_folder_button = tk.Button(main_frame, text="Open Output Folder", command=open_output_folder)
open_folder_button.grid(row=5, column=0, columnspan=3, pady=5)
open_folder_button.config(state="disabled")

# Quit
quit_button = tk.Button(main_frame, text="Quit", command=quit_app)
quit_button.grid(row=4, column=1, pady=10)

root.mainloop()
