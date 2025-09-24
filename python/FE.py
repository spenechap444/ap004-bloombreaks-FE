# python
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import tempfile
import webbrowser
from infrastructure.email import Email
import csv
from tkinter import simpledialog
from pathlib import Path
import json

class CSVApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CSV Attacher")
        self.geometry("900x600")
        self.configure(bg="#2b2b2b")  # charcoal background

        # keep separate lists per purpose for simpler mapping
        self.email_attachments = []   # list of {"path", "headers", "rows"}
        self.slot_attachments = []

        self._setup_style()
        self._create_widgets()

    def _setup_style(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Red.Treeview",
                        background="#2b2b2b",
                        fieldbackground="#2b2b2b",
                        foreground="#e53935",
                        rowheight=20)
        style.map("TButton",
                  background=[("active", "#2b0000")],
                  foreground=[("active", "#ffffff")])

    def _create_widgets(self):
        self._create_top_frame()
        self._create_notebook()
        self._create_email_tab()
        self._create_slots_tab()
        self._create_tools_tab()

    def _create_top_frame(self):
        top_frame = tk.Frame(self, bg="#2b2b2b")
        top_frame.pack(side="top", fill="x", padx=10, pady=10)
        self.status = tk.Label(top_frame, text="No attachments", bg="#2b2b2b", fg="#e53935")
        self.status.pack(side="right")

    def _create_notebook(self):
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=(0,10))

    def _create_email_tab(self):
        self.email_frame = tk.Frame(self.notebook, bg="#2b2b2b")
        self.notebook.add(self.email_frame, text="Emails")

        # top row: attach / clear for emails
        btn_row = tk.Frame(self.email_frame, bg="#2b2b2b")
        btn_row.pack(fill="x", padx=6, pady=(6,4))

        attach_emails_btn = tk.Button(btn_row, text="Attach Emails CSV", command=lambda: self.attach_csv("emails"),
                                      bg="#b71c1c", fg="white", relief="flat", padx=8, pady=6)
        attach_emails_btn.pack(side="left")

        clear_email_btn = tk.Button(btn_row, text="Clear Emails", command=self.clear_email_attachments,
                                    bg="#2b0000", fg="white", relief="flat", padx=8, pady=6)
        clear_email_btn.pack(side="left", padx=(8,0))

        # listbox at left and preview (tree) below/right
        body = tk.PanedWindow(self.email_frame, orient="horizontal", sashrelief="flat", bg="#2b2b2b")
        body.pack(fill="both", expand=True, padx=6, pady=6)

        # left: list of attached email CSVs
        left = tk.Frame(body, bg="#2b2b2b")
        body.add(left, minsize=220)

        tk.Label(left, text="Email CSVs", bg="#2b2b2b", fg="#e53935").pack(anchor="nw", padx=6, pady=(0,4))
        self.email_listbox = tk.Listbox(left, bg="#2b2b2b", fg="#e53935", selectbackground="#4a0000", highlightthickness=0)
        self.email_listbox.pack(fill="both", expand=True, padx=6, pady=(0,6))
        self.email_listbox.bind("<<ListboxSelect>>", self.on_email_select)

        # right: preview tree
        right = tk.Frame(body, bg="#2b2b2b")
        body.add(right, stretch="always")

        tk.Label(right, text="Preview", bg="#2b2b2b", fg="#e53935").pack(anchor="nw", padx=6, pady=(0,4))
        tree_frame = tk.Frame(right, bg="#2b2b2b")
        tree_frame.pack(fill="both", expand=True, padx=6, pady=(0,6))

        self.email_tree = ttk.Treeview(tree_frame, show="headings", style="Red.Treeview")
        self.email_tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.email_tree.yview)
        vsb.pack(side="right", fill="y")
        self.email_tree.configure(yscrollcommand=vsb.set)

    def _create_slots_tab(self):
        self.slots_frame = tk.Frame(self.notebook, bg="#2b2b2b")
        self.notebook.add(self.slots_frame, text="Purchase Slots")

        btn_row = tk.Frame(self.slots_frame, bg="#2b2b2b")
        btn_row.pack(fill="x", padx=6, pady=(6,4))

        attach_slots_btn = tk.Button(btn_row, text="Attach Slots CSV", command=lambda: self.attach_csv("slots"),
                                     bg="#b71c1c", fg="white", relief="flat", padx=8, pady=6)
        attach_slots_btn.pack(side="left")

        clear_slots_btn = tk.Button(btn_row, text="Clear Slots", command=self.clear_slot_attachments,
                                    bg="#2b0000", fg="white", relief="flat", padx=8, pady=6)
        clear_slots_btn.pack(side="left", padx=(8,0))

        body = tk.PanedWindow(self.slots_frame, orient="horizontal", sashrelief="flat", bg="#2b2b2b")
        body.pack(fill="both", expand=True, padx=6, pady=6)

        left = tk.Frame(body, bg="#2b2b2b")
        body.add(left, minsize=220)

        tk.Label(left, text="Slots CSVs", bg="#2b2b2b", fg="#e53935").pack(anchor="nw", padx=6, pady=(0,4))
        self.slots_listbox = tk.Listbox(left, bg="#2b2b2b", fg="#e53935", selectbackground="#4a0000", highlightthickness=0)
        self.slots_listbox.pack(fill="both", expand=True, padx=6, pady=(0,6))
        self.slots_listbox.bind("<<ListboxSelect>>", self.on_slot_select)

        right = tk.Frame(body, bg="#2b2b2b")
        body.add(right, stretch="always")

        tk.Label(right, text="Preview", bg="#2b2b2b", fg="#e53935").pack(anchor="nw", padx=6, pady=(0,4))
        tree_frame = tk.Frame(right, bg="#2b2b2b")
        tree_frame.pack(fill="both", expand=True, padx=6, pady=(0,6))

        self.slots_tree = ttk.Treeview(tree_frame, show="headings", style="Red.Treeview")
        self.slots_tree.pack(side="left", fill="both", expand=True)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.slots_tree.yview)
        vsb.pack(side="right", fill="y")
        self.slots_tree.configure(yscrollcommand=vsb.set)

    def _create_tools_tab(self):
        self.tools_frame = tk.Frame(self.notebook, bg="#2b2b2b")
        self.notebook.add(self.tools_frame, text="Email utility")

        # Email title/subject
        tk.Label(self.tools_frame, text="Email Title", bg="#2b2b2b", fg="#e53935").pack(anchor="nw", padx=8, pady=(8,0))
        self.email_title_entry = tk.Entry(self.tools_frame, bg="#1f1f1f", fg="white", insertbackground="white")
        self.email_title_entry.pack(fill="x", padx=8, pady=(0,6))

        tk.Label(self.tools_frame, text="Email Subject", bg="#2b2b2b", fg="#e53935").pack(anchor="nw", padx=8, pady=(0,0))
        self.email_subject_entry = tk.Entry(self.tools_frame, bg="#1f1f1f", fg="white", insertbackground="white")
        self.email_subject_entry.pack(fill="x", padx=8, pady=(0,8))

    # SMTP settings are handled in the Settings dialog; not shown here

        # Settings button to persist defaults
        settings_btn = tk.Button(self.tools_frame, text="Settings", bg="#4a4a4a", fg="white", relief="flat", command=self.open_settings)
        settings_btn.pack(padx=8, pady=(0,8), anchor="w")

        # Teams selection
        teams_label = tk.Label(self.tools_frame, text="Select teams to include in email:", bg="#2b2b2b", fg="#e53935")
        teams_label.pack(anchor="nw", padx=8, pady=(0,6))

        cb_wrap = tk.Frame(self.tools_frame, bg="#2b2b2b")
        cb_wrap.pack(fill="both", expand=False, padx=8, pady=(0,8))

        canvas = tk.Canvas(cb_wrap, bg="#2b2b2b", highlightthickness=0, height=220)
        scrollbar = ttk.Scrollbar(cb_wrap, orient="vertical", command=canvas.yview)
        checkbox_container = tk.Frame(canvas, bg="#2b2b2b")

        checkbox_container.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=checkbox_container, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        NHL_TEAMS = [
            "Anaheim Ducks","Arizona Coyotes","Boston Bruins","Buffalo Sabres","Calgary Flames",
            "Carolina Hurricanes","Chicago Blackhawks","Colorado Avalanche","Columbus Blue Jackets",
            "Dallas Stars","Detroit Red Wings","Edmonton Oilers","Florida Panthers","Los Angeles Kings",
            "Minnesota Wild","Montreal Canadiens","Nashville Predators","New Jersey Devils",
            "New York Islanders","New York Rangers","Ottawa Senators","Philadelphia Flyers",
            "Pittsburgh Penguins","San Jose Sharks","Seattle Kraken","St. Louis Blues",
            "Tampa Bay Lightning","Toronto Maple Leafs","Vancouver Canucks","Vegas Golden Knights",
            "Washington Capitals","Winnipeg Jets"
        ]

        self.team_vars = {}
        for team in NHL_TEAMS:
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(checkbox_container, text=team, variable=var,
                                bg="#2b2b2b", fg="#e53935", selectcolor="#2b0000",
                                activebackground="#2b2b2b", anchor="w")
            cb.pack(fill="x", anchor="w", padx=4, pady=1)
            self.team_vars[team] = var

        publish_btn = tk.Button(self.tools_frame, text="Publish Email (use remaining teams)", bg="#b71c1c", fg="white",
                                command=self.publish_email, relief="flat", padx=10, pady=6)
        publish_btn.pack(padx=8, pady=(0,12), anchor="e")

        # Load settings to prefill SMTP fields
        try:
            self._load_smtp_settings()
        except Exception:
            pass

    def _settings_file_path(self):
        return Path.home() / '.bloombreaks_email_settings.json'

    def _load_smtp_settings(self):
        p = self._settings_file_path()
        if not p.exists():
            # default values
            defaults = {
                'sender': '',
                'smtp_host': 'smtp.gmail.com',
                'smtp_port': 587,
                'smtp_user': '',
                'smtp_password': ''
            }
            settings = defaults
        else:
            with p.open('r', encoding='utf-8') as fh:
                settings = json.load(fh)
    # cache settings for programmatic use (UI entries were removed)
        self._cached_smtp_settings = settings
        return settings

    def _get_smtp_settings(self):
        """Return settings dict from file or defaults."""
        p = self._settings_file_path()
        if not p.exists():
            return {
                'sender': '',
                'smtp_host': 'smtp.gmail.com',
                'smtp_port': 587,
                'smtp_user': '',
                'smtp_password': ''
            }
        try:
            with p.open('r', encoding='utf-8') as fh:
                return json.load(fh)
        except Exception:
            return {
                'sender': '',
                'smtp_host': 'smtp.gmail.com',
                'smtp_port': 587,
                'smtp_user': '',
                'smtp_password': ''
            }

    def open_settings(self):
        # modal dialog to edit SMTP defaults
        dlg = tk.Toplevel(self)
        dlg.title('SMTP Settings')
        dlg.transient(self)
        dlg.grab_set()

        tk.Label(dlg, text='Sender email').grid(row=0, column=0, sticky='w', padx=8, pady=6)
        sender_e = tk.Entry(dlg)
        sender_e.grid(row=0, column=1, padx=8, pady=6)

        tk.Label(dlg, text='SMTP host').grid(row=1, column=0, sticky='w', padx=8, pady=6)
        host_e = tk.Entry(dlg)
        host_e.grid(row=1, column=1, padx=8, pady=6)

        tk.Label(dlg, text='SMTP port').grid(row=2, column=0, sticky='w', padx=8, pady=6)
        port_e = tk.Entry(dlg)
        port_e.grid(row=2, column=1, padx=8, pady=6)

        tk.Label(dlg, text='SMTP user').grid(row=3, column=0, sticky='w', padx=8, pady=6)
        user_e = tk.Entry(dlg)
        user_e.grid(row=3, column=1, padx=8, pady=6)

        tk.Label(dlg, text='SMTP password').grid(row=4, column=0, sticky='w', padx=8, pady=6)
        pass_e = tk.Entry(dlg, show='*')
        pass_e.grid(row=4, column=1, padx=8, pady=6)

        # prefill with current values
        try:
            p = self._settings_file_path()
            if p.exists():
                with p.open('r', encoding='utf-8') as fh:
                    s = json.load(fh)
            else:
                s = {'sender': '', 'smtp_host': 'smtp.gmail.com', 'smtp_port': 587, 'smtp_user': '', 'smtp_password': ''}
        except Exception:
            s = {'sender': '', 'smtp_host': 'smtp.gmail.com', 'smtp_port': 587, 'smtp_user': '', 'smtp_password': ''}

        sender_e.insert(0, s.get('sender', ''))
        host_e.insert(0, s.get('smtp_host', 'smtp.gmail.com'))
        port_e.insert(0, str(s.get('smtp_port', 587)))
        user_e.insert(0, s.get('smtp_user', ''))
        pass_e.insert(0, s.get('smtp_password', ''))

        def on_save():
            try:
                val = {
                    'sender': sender_e.get().strip(),
                    'smtp_host': host_e.get().strip() or 'smtp.gmail.com',
                    'smtp_port': int(port_e.get().strip() or 587),
                    'smtp_user': user_e.get().strip(),
                    'smtp_password': pass_e.get()
                }
            except Exception:
                messagebox.showerror('Invalid', 'Please enter a valid port number')
                return
            p = self._settings_file_path()
            try:
                with p.open('w', encoding='utf-8') as fh:
                    json.dump(val, fh)
            except Exception as e:
                messagebox.showerror('Error', f'Failed to save settings: {e}')
                return
            # apply to main entries
            self._load_smtp_settings()
            dlg.destroy()

        btn_frame = tk.Frame(dlg)
        btn_frame.grid(row=5, column=0, columnspan=2, pady=(6,8))
        tk.Button(btn_frame, text='Save', command=on_save).pack(side='left', padx=6)
        tk.Button(btn_frame, text='Cancel', command=dlg.destroy).pack(side='left', padx=6)

    def attach_csv(self, purpose="emails"):
        path = filedialog.askopenfilename(title=f"Select CSV file for {purpose}", filetypes=[("CSV files","*.csv"), ("All files","*.*")])
        if not path:
            return
        p = Path(path)
        try:
            with p.open(newline="", encoding="utf-8") as fh:
                reader = csv.reader(fh)
                rows = list(reader)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV:\n{e}")
            return

        if not rows:
            messagebox.showinfo("Empty", "CSV file is empty.")
            return

        headers = rows[0]
        data = rows[1:] if len(rows) > 1 else []
        attachment = {"path": p, "headers": headers, "rows": data}

        if purpose == "emails":
            self.email_attachments.append(attachment)
            display = f"{p.name}"
            self.email_listbox.insert("end", display)
            self.email_listbox.selection_clear(0, "end")
            self.email_listbox.selection_set("end")
            self.email_listbox.event_generate("<<ListboxSelect>>")
            self._show_preview_for_tree(attachment, self.email_tree)
        else:
            self.slot_attachments.append(attachment)
            display = f"{p.name}"
            self.slots_listbox.insert("end", display)
            self.slots_listbox.selection_clear(0, "end")
            self.slots_listbox.selection_set("end")
            self.slots_listbox.event_generate("<<ListboxSelect>>")
            self._show_preview_for_tree(attachment, self.slots_tree)

        self._update_status()

    def on_email_select(self, event):
        sel = self.email_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        attachment = self.email_attachments[idx]
        self._show_preview_for_tree(attachment, self.email_tree)

    def on_slot_select(self, event):
        sel = self.slots_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        attachment = self.slot_attachments[idx]
        self._show_preview_for_tree(attachment, self.slots_tree)

    def _show_preview_for_tree(self, attachment, tree, max_rows=200):
        # clear tree
        for c in tree.get_children():
            tree.delete(c)
        tree["columns"] = attachment["headers"]
        for h in attachment["headers"]:
            tree.heading(h, text=h, anchor="w")
            tree.column(h, width=140, anchor="w", stretch=True)
        for r in attachment["rows"][:max_rows]:
            row = list(r) + [""] * max(0, len(attachment["headers"]) - len(r))
            tree.insert("", "end", values=row)

    def clear_email_attachments(self):
        if not self.email_attachments:
            return
        if not messagebox.askyesno("Clear", "Remove all email CSV attachments?"):
            return
        self.email_attachments.clear()
        self.email_listbox.delete(0, "end")
        for c in self.email_tree.get_children():
            self.email_tree.delete(c)
        self.email_tree["columns"] = []
        self._update_status()

    def clear_slot_attachments(self):
        if not self.slot_attachments:
            return
        if not messagebox.askyesno("Clear", "Remove all slots CSV attachments?"):
            return
        self.slot_attachments.clear()
        self.slots_listbox.delete(0, "end")
        for c in self.slots_tree.get_children():
            self.slots_tree.delete(c)
        self.slots_tree["columns"] = []
        self._update_status()

    def _update_status(self):
        total = len(self.email_attachments) + len(self.slot_attachments)
        self.status.config(text=f"{total} attachments (emails: {len(self.email_attachments)}, slots: {len(self.slot_attachments)})")

    def publish_email(self):
        title = self.email_title_entry.get().strip()
        subject = self.email_subject_entry.get().strip()
        remaining = [t for t,v in self.team_vars.items() if not v.get()]
        if not title:
            messagebox.showwarning("Missing", "Please enter an email title.")
            return
        if not subject:
            messagebox.showwarning("Missing", "Please enter an email subject.")
            return
        if not remaining:
            messagebox.showinfo("No recipients", "No remaining teams selected — nothing to publish.")
            return
        # Build a simple HTML body summarizing the publish
        body_html = f"<h2>{title}</h2><h3>{subject}</h3><p>Teams to publish to ({len(remaining)}):<br>" + ", ".join(remaining) + "</p>"

        # Prepare recipients: try to collect from attached email CSVs (column 'email'), else ask user
        recipients = []
        for att in self.email_attachments:
            # rows are lists; try to find email column
            headers = [h.lower() for h in att['headers']]
            if 'email' in headers:
                idx = headers.index('email')
                for r in att['rows']:
                    if len(r) > idx and r[idx].strip():
                        recipients.append(r[idx].strip())

        if not recipients:
            # prompt for a single recipient (comma separated allowed)
            ans = filedialog.askopenfilename(title="No recipient emails found - select a file with recipients or Cancel to enter manually", filetypes=[("CSV files","*.csv"), ("All files","*.*")])
            # don't force; if user cancels continue to prompt via input dialog
            if ans:
                # try to read simple CSV and extract first column as emails
                try:
                    p = Path(ans)
                    with p.open(newline="", encoding='utf-8') as fh:
                        r = list(csv.reader(fh))
                    if r:
                        for row in r[1:]:
                            if row:
                                recipients.append(row[0].strip())
                except Exception:
                    pass

        if not recipients:
            # last resort: ask user to type comma-separated recipients
            recip_str = tk.simpledialog.askstring("Recipients", "Enter recipient email addresses (comma-separated):")
            if not recip_str:
                messagebox.showinfo("No recipients", "No recipients provided — cancelled.")
                return
            recipients = [r.strip() for r in recip_str.split(',') if r.strip()]

        # Format the message using Email helper
        # Load SMTP settings from persisted settings (Settings dialog)
        settings = self._get_smtp_settings()
        sender_addr = settings.get('sender') or 'no-reply@example.com'
        smtp_host = settings.get('smtp_host') or 'smtp.gmail.com'
        smtp_port = int(settings.get('smtp_port') or 587)
        smtp_user = settings.get('smtp_user') or None
        smtp_password = settings.get('smtp_password') or None

        email_client = Email(sender=sender_addr, smtp_host=smtp_host, smtp_port=smtp_port, smtp_user=smtp_user, smtp_password=smtp_password)
        msg_obj = email_client.format_message(title=title, subject=subject, body_html=body_html, recipients=recipients)

        # Create preview file in temp and open it
        tf = tempfile.NamedTemporaryFile(delete=False, suffix='.msg')
        tf.close()
        preview_path = email_client.create_msg_file(msg_obj, tf.name)
        # Try to open preview file for user
        try:
            os.startfile(preview_path)
        except Exception:
            try:
                webbrowser.open(preview_path)
            except Exception:
                messagebox.showinfo("Preview saved", f"Preview saved to: {preview_path}")

        # Ask user to confirm sending
        if not messagebox.askyesno("Send", f"Send this email to {len(recipients)} recipients?"):
            return

        # Send to each recipient (simple loop)
        for r in recipients:
            try:
                email_client.send_mail(r, subject, body_html)
            except Exception as e:
                logging = __import__('logging')
                logging.error(f"Failed to send to {r}: {e}")

        messagebox.showinfo("Done", "Emails queued for sending (see logs for errors).")

if __name__ == "__main__":
    app = CSVApp()
    app.mainloop()