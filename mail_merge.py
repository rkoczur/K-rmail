import os
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from email.message import EmailMessage

try:
    import pandas as pd
except ImportError:
    pd = None


class DefaultDict(dict):
    def __missing__(self, key):
        return ""


class MailMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Merge Sender")
        self.df = None
        self.file_path = None

        self._create_widgets()

    def _create_widgets(self):
        # Configure root grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

        # Create canvas and scrollbar
        canvas = tk.Canvas(self.root, bg="white")
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        # Create main frame inside canvas
        main = ttk.Frame(canvas, padding=10)
        canvas_window = canvas.create_window(0, 0, window=main, anchor="nw")

        # Update scroll region when frame changes
        def on_frame_configure(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            # Make frame width match canvas width
            canvas.itemconfig(canvas_window, width=canvas.winfo_width() if canvas.winfo_width() > 1 else 860)

        main.bind("<Configure>", on_frame_configure)
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))

        # Bind mouse wheel to scroll
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", on_mousewheel)
        
        # For Linux
        def on_scroll_up(event):
            canvas.yview_scroll(-3, "units")
        
        def on_scroll_down(event):
            canvas.yview_scroll(3, "units")
        
        canvas.bind_all("<Button-4>", on_scroll_up)
        canvas.bind_all("<Button-5>", on_scroll_down)

        # SMTP Settings Section
        smtp_frame = ttk.LabelFrame(main, text="SMTP Settings", padding=10)
        smtp_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ttk.Label(smtp_frame, text="Server: ").grid(row=0, column=0, sticky="w")
        self.smtp_host = ttk.Entry(smtp_frame, width=30)
        self.smtp_host.grid(row=0, column=1, sticky="w", padx=5)
        self.smtp_host.insert(0, "smtp.example.com")

        ttk.Label(smtp_frame, text="Port: ").grid(row=0, column=2, sticky="w", padx=(15, 0))
        self.smtp_port = ttk.Entry(smtp_frame, width=8)
        self.smtp_port.grid(row=0, column=3, sticky="w", padx=5)
        self.smtp_port.insert(0, "587")

        ttk.Label(smtp_frame, text="Connection: ").grid(row=0, column=4, sticky="w", padx=(15, 0))
        self.connection_type = ttk.Combobox(smtp_frame, values=("None", "TLS", "SSL"), state="readonly", width=8)
        self.connection_type.grid(row=0, column=5, sticky="w", padx=5)
        self.connection_type.set("TLS")

        ttk.Label(smtp_frame, text="Username: ").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.smtp_user = ttk.Entry(smtp_frame, width=30)
        self.smtp_user.grid(row=1, column=1, sticky="w", padx=5, pady=(10, 0))

        ttk.Label(smtp_frame, text="Password: ").grid(row=1, column=2, sticky="w", padx=(15, 0), pady=(10, 0))
        self.smtp_pass = ttk.Entry(smtp_frame, width=20, show="*")
        self.smtp_pass.grid(row=1, column=3, sticky="w", padx=5, pady=(10, 0))

        ttk.Label(smtp_frame, text="From Email: ").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.from_email = ttk.Entry(smtp_frame, width=30)
        self.from_email.grid(row=2, column=1, sticky="w", padx=5, pady=(10, 0))

        ttk.Label(smtp_frame, text="Subject: ").grid(row=2, column=2, sticky="w", padx=(15, 0), pady=(10, 0))
        self.subject = ttk.Entry(smtp_frame, width=40)
        self.subject.grid(row=2, column=3, columnspan=2, sticky="ew", padx=5, pady=(10, 0))

        excel_frame = ttk.LabelFrame(main, text="Load Recipients", padding=10)
        excel_frame.grid(row=1, column=0, sticky="ew", pady=(0, 10))

        load_button = ttk.Button(excel_frame, text="Load Excel/CSV", command=self.load_excel)
        load_button.grid(row=0, column=0, sticky="w")

        self.file_label = ttk.Label(excel_frame, text="No file loaded.")
        self.file_label.grid(row=0, column=1, sticky="w", padx=10)

        ttk.Label(excel_frame, text="Email column: ").grid(row=1, column=0, sticky="w", pady=(10, 0))
        self.email_field = ttk.Combobox(excel_frame, state="readonly", width=28)
        self.email_field.grid(row=1, column=1, sticky="w", padx=10, pady=(10, 0))

        self.columns_list = tk.Listbox(excel_frame, height=6, selectmode="browse")
        self.columns_list.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        self.columns_list.bind("<<ListboxSelect>>", lambda event: None)

        self.sample_text = tk.Text(excel_frame, height=5, wrap="word", state="disabled")
        self.sample_text.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        template_frame = ttk.LabelFrame(main, text="Letter Template", padding=10)
        template_frame.grid(row=2, column=0, sticky="nsew")
        template_frame.columnconfigure(0, weight=1)
        template_frame.rowconfigure(1, weight=1)

        self.template_text = tk.Text(template_frame, wrap="word", height=16)
        self.template_text.grid(row=1, column=0, sticky="nsew")

        field_button = ttk.Button(template_frame, text="Insert Field", command=self.insert_field)
        field_button.grid(row=0, column=0, sticky="w")

        button_frame = ttk.Frame(main)
        button_frame.grid(row=3, column=0, sticky="e")

        send_button = ttk.Button(button_frame, text="Send", command=self.send_messages)
        send_button.grid(row=0, column=0, padx=5)

        preview_button = ttk.Button(button_frame, text="Preview", command=self.show_preview)
        preview_button.grid(row=0, column=1, padx=5)

        info_text = ttk.Label(main, text="Use field names in curly brackets in the template, e.g., {Name}.")
        info_text.grid(row=4, column=0, sticky="w", pady=(8, 0))

    def load_excel(self):
        if pd is None:
            messagebox.showerror("Missing library", "The pandas and openpyxl packages are required. Install them according to requirements.txt.")
            return

        file_path = filedialog.askopenfilename(
            title="Select an Excel or CSV file",
            filetypes=[("Excel file", "*.xlsx *.xls"), ("CSV file", "*.csv"), ("All files", "*")],
        )
        if not file_path:
            return

        try:
            if file_path.lower().endswith(".csv"):
                df = pd.read_csv(file_path, dtype=str)
            else:
                df = pd.read_excel(file_path, dtype=str)
        except Exception as exc:
            messagebox.showerror("Error opening file", f"Failed to load the file:\n{exc}")
            return

        if df.empty:
            messagebox.showwarning("Empty file", "The loaded file contains no data.")
            return

        self.df = df.fillna("")
        self.file_path = file_path
        self.file_label.config(text=os.path.basename(file_path))

        columns = list(self.df.columns)
        self.email_field.config(values=columns)
        self.email_field.set(columns[0] if columns else "")

        self.columns_list.delete(0, tk.END)
        for col in columns:
            self.columns_list.insert(tk.END, col)

        self._update_sample_text()

    def _update_sample_text(self):
        if self.df is None:
            return
        sample = self.df.iloc[0].to_dict()
        display = "Sample of first row:\n"
        display += "\n".join([f"{k}: {v}" for k, v in sample.items()])
        self.sample_text.config(state="normal")
        self.sample_text.delete("1.0", tk.END)
        self.sample_text.insert(tk.END, display)
        self.sample_text.config(state="disabled")

    def insert_field(self):
        selection = self.columns_list.curselection()
        if not selection:
            messagebox.showinfo("Choose field", "First select a column from the list.")
            return
        field = self.columns_list.get(selection[0])
        self.template_text.insert(tk.INSERT, f"{{{field}}}")

    def show_preview(self):
        if self.df is None:
            messagebox.showinfo("No data", "First load the file containing recipients.")
            return
        if not self.email_field.get():
            messagebox.showinfo("Email column missing", "Select which column contains the email addresses.")
            return

        text = self.template_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showinfo("No template", "Write a letter for the template.")
            return

        sample = self.df.iloc[0].to_dict()
        preview = self._render_template(text, sample)
        preview_subject = self._render_template(self.subject.get(), sample)
        messagebox.showinfo("Preview", f"Subject: {preview_subject}\n\n{preview}")

    def _render_template(self, template, row_values):
        try:
            return template.format_map(DefaultDict(row_values))
        except Exception as exc:
            return f"Error processing the template: {exc}"

    def send_messages(self):
        if self.df is None:
            messagebox.showinfo("No data", "First load the file containing recipients.")
            return
        email_column = self.email_field.get().strip()
        if not email_column:
            messagebox.showinfo("Email column missing", "Select which column contains the email addresses.")
            return
        if email_column not in self.df.columns:
            messagebox.showerror("Error", "The selected email column was not found in the table.")
            return

        template = self.template_text.get("1.0", tk.END).strip()
        subject_template = self.subject.get().strip()
        from_email = self.from_email.get().strip()
        if not from_email:
            messagebox.showinfo("Sender missing", "Enter the sender's email address.")
            return
        if not template:
            messagebox.showinfo("No template", "Write a letter for the template.")
            return

        if not messagebox.askyesno("Confirmation", f"Are you sure you want to send emails to {len(self.df)} recipients?"):
            return

        smtp_host = self.smtp_host.get().strip()
        smtp_port = self.smtp_port.get().strip()
        smtp_user = self.smtp_user.get().strip()
        smtp_pass = self.smtp_pass.get().strip()
        connection_type = self.connection_type.get()

        try:
            port = int(smtp_port)
        except ValueError:
            messagebox.showerror("Error", "Port must be a number.")
            return

        success = 0
        errors = []

        try:
            if connection_type == "SSL":
                server = smtplib.SMTP_SSL(smtp_host, port, timeout=20)
            else:
                server = smtplib.SMTP(smtp_host, port, timeout=20)
                if connection_type == "TLS":
                    server.starttls()
            if smtp_user and smtp_pass:
                server.login(smtp_user, smtp_pass)
        except Exception as exc:
            messagebox.showerror("SMTP error", f"Failed to connect to the SMTP server:\n{exc}")
            return

        for index, row in self.df.iterrows():
            recipient = str(row.get(email_column, "")).strip()
            if not recipient:
                errors.append((index + 1, "Empty email address"))
                continue
            message = EmailMessage()
            message["From"] = from_email
            message["To"] = recipient
            message["Subject"] = self._render_template(subject_template, row)
            body = self._render_template(template, row)
            message.set_content(body)

            try:
                server.send_message(message)
                success += 1
            except Exception as exc:
                errors.append((index + 1, str(exc)))

        server.quit()

        summary = f"Successfully sent: {success}\n"
        if errors:
            summary += f"Errors: {len(errors)}\n"
            summary += "Rows:\n"
            summary += "\n".join([f"{row}: {err}" for row, err in errors])
        messagebox.showinfo("Send result", summary)


def main():
    root = tk.Tk()
    root.geometry("860x700")
    app = MailMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
