#!/usr/bin/env python3
"""
eM Client Universal Migrator - Graphical User Interface
========================================================
Interactive Tkinter GUI allowing users to select individual accounts,
specific folders within accounts, calendar, contacts, and choose between
migrating to Mozilla Thunderbird, Microsoft Outlook (PST/Profile), or Standalone Export.
"""

import os
import sys
import threading
import queue
import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from migrator import (
    EMClientReader,
    ThunderbirdTarget,
    OutlookTarget,
    HAS_WIN32COM,
    clean_filename
)

class MigratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("eM Client Universal Migrator (to Thunderbird & Outlook)")
        self.root.geometry("900x720")
        self.root.minsize(800, 600)

        # Apply modern ttk theme
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        self.reader = None
        self.data_tree = {}
        self.log_queue = queue.Queue()
        self.is_migrating = False

        self.setup_ui()
        self.root.after(100, self.initial_load)
        self.root.after(100, self.process_log_queue)

    def setup_ui(self):
        # 1. Top Header Frame
        header_frame = ttk.Frame(self.root, padding="10")
        header_frame.pack(fill=tk.X)

        title_lbl = ttk.Label(header_frame, text="eM Client Universal Migrator", font=("Segoe UI", 16, "bold"))
        title_lbl.pack(anchor=tk.W)
        sub_lbl = ttk.Label(header_frame, text="Transfer Accounts, Emails, Folders, Calendars & Contacts to Thunderbird or Outlook", font=("Segoe UI", 10))
        sub_lbl.pack(anchor=tk.W)

        # 2. Source & Target Selection Frame
        config_frame = ttk.LabelFrame(self.root, text=" 1. Migration Paths & Destination ", padding="10")
        config_frame.pack(fill=tk.X, padx=10, pady=5)

        # eM Client Path
        ttk.Label(config_frame, text="eM Client Data:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.em_path_var = tk.StringVar()
        em_entry = ttk.Entry(config_frame, textvariable=self.em_path_var, width=65)
        em_entry.grid(row=0, column=1, sticky=tk.EW, padx=5, pady=2)
        ttk.Button(config_frame, text="Browse...", command=self.browse_em_path).grid(row=0, column=2, padx=2, pady=2)

        # Migration Target Options
        ttk.Label(config_frame, text="Target Application:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.target_var = tk.StringVar(value="thunderbird")
        
        target_btn_frame = ttk.Frame(config_frame)
        target_btn_frame.grid(row=1, column=1, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Radiobutton(target_btn_frame, text="Mozilla Thunderbird", variable=self.target_var, value="thunderbird", command=self.on_target_changed).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(target_btn_frame, text="Microsoft Outlook (PST)", variable=self.target_var, value="outlook", command=self.on_target_changed).pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(target_btn_frame, text="Standalone Export (MBOX/EML/ICS)", variable=self.target_var, value="export", command=self.on_target_changed).pack(side=tk.LEFT, padx=5)

        # Thunderbird Profile Path Frame
        self.tb_frame = ttk.Frame(config_frame)
        self.tb_frame.grid(row=2, column=0, columnspan=3, sticky=tk.EW, pady=2)
        ttk.Label(self.tb_frame, text="Thunderbird Profile:").grid(row=0, column=0, sticky=tk.W)
        self.tb_path_var = tk.StringVar()
        ttk.Entry(self.tb_frame, textvariable=self.tb_path_var, width=63).grid(row=0, column=1, sticky=tk.EW, padx=5)
        ttk.Button(self.tb_frame, text="Browse...", command=self.browse_tb_path).grid(row=0, column=2, padx=2)

        config_frame.columnconfigure(1, weight=1)

        # 3. Component Selection Checkboxes
        comp_frame = ttk.LabelFrame(self.root, text=" 2. Components to Transfer ", padding="10")
        comp_frame.pack(fill=tk.X, padx=10, pady=5)

        self.inc_accounts_var = tk.BooleanVar(value=True)
        self.inc_mails_var = tk.BooleanVar(value=True)
        self.inc_cal_var = tk.BooleanVar(value=True)
        self.inc_contacts_var = tk.BooleanVar(value=True)

        self.chk_accounts = ttk.Checkbutton(comp_frame, text="Account Settings & Identities (Servers, Ports, OAuth2)", variable=self.inc_accounts_var)
        self.chk_accounts.pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(comp_frame, text="Emails & Folders", variable=self.inc_mails_var).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(comp_frame, text="Calendar Events", variable=self.inc_cal_var).pack(side=tk.LEFT, padx=10)
        ttk.Checkbutton(comp_frame, text="Contacts (Address Book)", variable=self.inc_contacts_var).pack(side=tk.LEFT, padx=10)

        # 4. Accounts & Folders Selection Treeview
        tree_frame = ttk.LabelFrame(self.root, text=" 3. Select Accounts and Folders to Migrate ", padding="10")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        tree_toolbar = ttk.Frame(tree_frame)
        tree_toolbar.pack(fill=tk.X, pady=(0, 5))
        ttk.Button(tree_toolbar, text="Select All", command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(tree_toolbar, text="Deselect All", command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(tree_toolbar, text="Refresh Tree", command=self.reload_tree).pack(side=tk.LEFT, padx=2)

        tree_scroll = ttk.Scrollbar(tree_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(tree_frame, columns=("count",), yscrollcommand=tree_scroll.set, selectmode="none")
        self.tree.heading("#0", text="Account / Folder Name", anchor=tk.W)
        self.tree.heading("count", text="Items", anchor=tk.E)
        self.tree.column("#0", stretch=True, width=550)
        self.tree.column("count", stretch=False, width=100, anchor=tk.E)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree.yview)

        self.tree.bind("<Button-1>", self.on_tree_click)

        # Map item IDs to checked state: True / False
        self.item_checked = {}
        self.item_type = {} # 'account' or 'folder'
        self.item_data = {} # metadata

        # 5. Progress and Log Frame
        bottom_frame = ttk.Frame(self.root, padding="10")
        bottom_frame.pack(fill=tk.X)

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(bottom_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X, pady=2)

        self.status_lbl = ttk.Label(bottom_frame, text="Ready. Select accounts and folders, then click Start Migration.", font=("Segoe UI", 9))
        self.status_lbl.pack(anchor=tk.W, pady=2)

        # Log box (collapsible / compact)
        self.log_text = tk.Text(bottom_frame, height=5, wrap=tk.WORD, font=("Consolas", 8), bg="#f8f9fa")
        self.log_text.pack(fill=tk.X, pady=2)

        # Start Button
        self.start_btn = ttk.Button(bottom_frame, text="▶ Start Migration", command=self.start_migration)
        self.start_btn.pack(side=tk.RIGHT, pady=5)

    def on_target_changed(self):
        target = self.target_var.get()
        if target == 'thunderbird':
            self.tb_frame.grid()
            self.chk_accounts.config(state=tk.NORMAL)
        else:
            self.tb_frame.grid_remove()
            if target == 'outlook':
                self.chk_accounts.config(state=tk.DISABLED)
                self.inc_accounts_var.set(False)
            else:
                self.chk_accounts.config(state=tk.DISABLED)
                self.inc_accounts_var.set(False)

    def browse_em_path(self):
        p = filedialog.askdirectory(title="Select eM Client Data Folder")
        if p:
            self.em_path_var.set(p)
            self.reload_tree()

    def browse_tb_path(self):
        p = filedialog.askdirectory(title="Select Thunderbird Profile Folder")
        if p:
            self.tb_path_var.set(p)

    def initial_load(self):
        default_em = EMClientReader.find_default_path()
        if default_em:
            self.em_path_var.set(str(default_em))

        default_tb = ThunderbirdTarget.find_default_profile()
        if default_tb:
            self.tb_path_var.set(str(default_tb))

        self.reload_tree()

    def reload_tree(self):
        em_path = self.em_path_var.get()
        if not em_path or not os.path.exists(em_path):
            return

        self.tree.delete(*self.tree.get_children())
        self.item_checked.clear()
        self.item_type.clear()
        self.item_data.clear()

        try:
            self.reader = EMClientReader(em_path)
            self.data_tree = self.reader.scan_data_tree()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read eM Client data: {e}")
            return

        for acc_name, info in sorted(self.data_tree.items()):
            # Calculate total emails
            total_mails = sum(f['count'] for f in info['folders'].values())
            summary_parts = []
            if total_mails > 0: summary_parts.append(f"{total_mails:,} mails")
            if info['cal_count'] > 0: summary_parts.append(f"{info['cal_count']} events")
            if info['contact_count'] > 0: summary_parts.append(f"{info['contact_count']} contacts")
            summary_str = ", ".join(summary_parts) or "0 items"

            node_id = self.tree.insert("", tk.END, text=f"☑ {acc_name}", values=(summary_str,), open=True)
            self.item_checked[node_id] = True
            self.item_type[node_id] = 'account'
            self.item_data[node_id] = {'account': acc_name, 'info': info}

            # Insert subfolders
            for f_path, f_info in sorted(info['folders'].items()):
                f_node = self.tree.insert(node_id, tk.END, text=f"  ☑ {f_path}", values=(f"{f_info['count']:,}",))
                self.item_checked[f_node] = True
                self.item_type[f_node] = 'folder'
                self.item_data[f_node] = {'account': acc_name, 'folder_path': f_path, 'folder_info': f_info}

        self.log(f"Loaded {len(self.data_tree)} accounts from eM Client.")

    def on_tree_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item: return

        # Toggle check state
        curr = self.item_checked.get(item, True)
        new_state = not curr
        self.set_item_check(item, new_state)

        # If account node, toggle all children
        if self.item_type.get(item) == 'account':
            for child in self.tree.get_children(item):
                self.set_item_check(child, new_state)
        # If folder node, check parent if any child is checked
        elif self.item_type.get(item) == 'folder':
            parent = self.tree.parent(item)
            if parent:
                any_checked = any(self.item_checked.get(c, False) for c in self.tree.get_children(parent))
                self.set_item_check(parent, any_checked)

    def set_item_check(self, item, checked):
        self.item_checked[item] = checked
        raw_text = self.tree.item(item, "text")
        clean_text = raw_text.replace("☑ ", "").replace("☐ ", "").strip()
        mark = "☑" if checked else "☐"
        prefix = "  " if self.item_type.get(item) == 'folder' else ""
        self.tree.item(item, text=f"{prefix}{mark} {clean_text}")

    def select_all(self):
        for item in self.tree.get_children():
            self.set_item_check(item, True)
            for child in self.tree.get_children(item):
                self.set_item_check(child, True)

    def deselect_all(self):
        for item in self.tree.get_children():
            self.set_item_check(item, False)
            for child in self.tree.get_children(item):
                self.set_item_check(child, False)

    def log(self, msg):
        self.log_queue.put(msg)

    def process_log_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            t = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{t}] {msg}\n")
            self.log_text.see(tk.END)
        self.root.after(100, self.process_log_queue)

    def start_migration(self):
        if self.is_migrating:
            return

        target = self.target_var.get()
        if target == 'thunderbird':
            tb_path = self.tb_path_var.get()
            if not tb_path or not os.path.exists(tb_path):
                messagebox.showerror("Error", "Please select a valid Thunderbird profile directory.")
                return
            lock_f = Path(tb_path) / 'parent.lock'
            if lock_f.exists():
                res = messagebox.askyesno("Thunderbird Running", "Thunderbird appears to be currently running.\nWould you like the migrator to terminate Thunderbird to safely apply updates?")
                if res:
                    if sys.platform == 'win32':
                        import subprocess
                        subprocess.run(['powershell', '-Command', "Get-Process -Name thunderbird -ErrorAction SilentlyContinue | Stop-Process -Force"], capture_output=True)
                    try:
                        if lock_f.exists(): lock_f.unlink()
                    except Exception: pass
                else:
                    return
        elif target == 'outlook':
            if not HAS_WIN32COM:
                messagebox.showerror("Outlook Error", "Microsoft Outlook COM integration requires the pywin32 library.\nPlease install pywin32 (`pip install pywin32`) or select Standalone Export.")
                return

        # Gather selected items
        selected_folders = [] # (account, folder_path)
        selected_accounts = set()
        for item, checked in self.item_checked.items():
            if checked and self.item_type.get(item) == 'folder':
                data = self.item_data[item]
                selected_folders.append((data['account'], data['folder_path']))
                selected_accounts.add(data['account'])

        if not selected_folders and not (self.inc_cal_var.get() or self.inc_contacts_var.get()):
            messagebox.showwarning("No Selection", "No accounts or folders are selected for migration.")
            return

        self.is_migrating = True
        self.start_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)

        # Launch migration in background thread
        thread = threading.Thread(target=self.run_migration_thread, args=(target, selected_accounts, selected_folders), daemon=True)
        thread.start()

    def run_migration_thread(self, target, selected_accounts, selected_folders):
        try:
            self.log(f"Starting migration to target: {target.upper()}")
            total_steps = len(selected_folders) + 2
            completed_steps = 0

            # 1. Accounts Settings (Thunderbird)
            if target == 'thunderbird' and self.inc_accounts_var.get():
                self.status_lbl.config(text="Configuring account settings in Thunderbird...")
                self.log("Updating Thunderbird prefs.js...")
                tb = ThunderbirdTarget(self.tb_path_var.get())
                tb.backup_prefs()
                # Filter accounts matching selection
                accs_to_add = [a for a in self.reader.accounts_config if a['email'] in selected_accounts]
                tb.deploy_accounts(accs_to_add)
                self.log(f"Configured {len(accs_to_add)} accounts in Thunderbird.")

            completed_steps += 1
            self.progress_var.set((completed_steps / total_steps) * 100)

            # 2. Email Migration
            if self.inc_mails_var.get() and selected_folders:
                if target == 'thunderbird':
                    tb = ThunderbirdTarget(self.tb_path_var.get())
                    # Write MBOX files
                    for acc_name, folder_path in selected_folders:
                        self.status_lbl.config(text=f"Migrating {acc_name} -> {folder_path}...")
                        info = self.data_tree.get(acc_name, {})
                        target_sbd = tb.profile / 'Mail' / 'Local Folders' / 'eM Client Migration.sbd' / f"{clean_filename(acc_name)}.sbd"
                        target_sbd.mkdir(parents=True, exist_ok=True)
                        
                        # Build .sbd path
                        parts = folder_path.split('/')
                        curr = target_sbd
                        for p in parts[:-1]:
                            (curr / p).touch(exist_ok=True)
                            curr = curr / f"{p}.sbd"
                            curr.mkdir(parents=True, exist_ok=True)

                        norm_name = parts[-1]
                        mbox_path = curr / norm_name
                        count = 0
                        with open(mbox_path, 'ab') as f_out:
                            for ad in info.get('acc_dirs', []):
                                for msg in self.reader.get_messages_generator(ad, folder_path):
                                    tb.write_mbox_message(f_out, msg)
                                    count += 1
                        
                        msf_path = curr / f"{norm_name}.msf"
                        if not msf_path.exists(): msf_path.touch()

                        self.log(f"  {acc_name} [{folder_path}]: {count} messages written to MBOX.")
                        completed_steps += 1
                        self.progress_var.set((completed_steps / total_steps) * 100)

                elif target == 'outlook':
                    ot = OutlookTarget(mode='pst')
                    ot.connect()
                    for acc_name, folder_path in selected_folders:
                        self.status_lbl.config(text=f"Importing into Outlook PST: {acc_name} -> {folder_path}...")
                        root_store = ot.get_or_create_store(acc_name)
                        parts = folder_path.split('/')
                        target_f = ot.get_target_folder(root_store, parts)
                        
                        info = self.data_tree.get(acc_name, {})
                        count = 0
                        for ad in info.get('acc_dirs', []):
                            for msg in self.reader.get_messages_generator(ad, folder_path):
                                ot.import_message(target_f, msg)
                                count += 1
                        self.log(f"  {acc_name} [{folder_path}]: {count} messages imported into Outlook.")
                        completed_steps += 1
                        self.progress_var.set((completed_steps / total_steps) * 100)

                elif target == 'export':
                    out_docs = Path(os.path.expanduser('~/Documents/Thunderbird Migration/Exported Mails'))
                    out_docs.mkdir(parents=True, exist_ok=True)
                    for acc_name, folder_path in selected_folders:
                        self.status_lbl.config(text=f"Exporting: {acc_name} -> {folder_path}...")
                        info = self.data_tree.get(acc_name, {})
                        target_dir = out_docs / clean_filename(acc_name) / folder_path.replace('/', os.sep)
                        target_dir.mkdir(parents=True, exist_ok=True)
                        count = 0
                        for ad in info.get('acc_dirs', []):
                            for msg in self.reader.get_messages_generator(ad, folder_path):
                                safe_subj = clean_filename(msg['subject'] or f"msg_{msg['id']}")[:50]
                                (target_dir / f"{msg['id']}_{safe_subj}.eml").write_bytes(msg['eml_bytes'])
                                count += 1
                        self.log(f"  {acc_name} [{folder_path}]: {count} .eml files exported.")
                        completed_steps += 1
                        self.progress_var.set((completed_steps / total_steps) * 100)

            # 3. Calendar & Contacts Export
            docs_dir = Path(os.path.expanduser('~/Documents/Thunderbird Migration'))
            docs_dir.mkdir(parents=True, exist_ok=True)

            if self.inc_cal_var.get():
                self.status_lbl.config(text="Exporting Calendars (.ics)...")
                self.log("Exporting calendars to standard iCalendar (.ics)...")
                # Scans all event_data.dat
                all_events = []
                for db in self.reader.data_path.glob('**/event_data.dat'):
                    try:
                        conn = sqlite3.connect(f'file:{db}?mode=ro', uri=True)
                        cur = conn.cursor()
                        cur.execute("SELECT properties FROM EventItemContents WHERE propertyName='calendar'")
                        for (props,) in cur.fetchall():
                            text = props.decode('utf-8', errors='replace') if isinstance(props, bytes) else str(props)
                            m = re.search(r'(BEGIN:VEVENT[\s\S]*?END:VEVENT)', text)
                            if m: all_events.append(m.group(1))
                        conn.close()
                    except Exception: pass

                if all_events:
                    out_ics = docs_dir / 'eMClient_All_Calendars.ics'
                    with open(out_ics, 'w', encoding='utf-8') as f:
                        f.write("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Universal Migrator//EN\r\n")
                        for ev in all_events: f.write(ev + "\r\n")
                        f.write("END:VCALENDAR\r\n")
                    self.log(f"Exported {len(all_events)} calendar events to {out_ics.name}")

            if self.inc_contacts_var.get():
                self.status_lbl.config(text="Exporting Contacts (.vcf)...")
                self.log("Exporting contacts to standard vCard (.vcf)...")
                all_vcards = []
                for db in self.reader.data_path.glob('**/contact_index.dat'):
                    try:
                        conn = sqlite3.connect(f'file:{db}?mode=ro', uri=True)
                        cur = conn.cursor()
                        cur.execute("SELECT id, firstName, lastName, displayName, organization, title FROM ContactItems")
                        for cid, fn, ln, dn, org, title in cur.fetchall():
                            v = ["BEGIN:VCARD", "VERSION:3.0", f"N:{ln or ''};{fn or ''};;;", f"FN:{dn or fn or ln or ''}"]
                            if org: v.append(f"ORG:{org}")
                            if title: v.append(f"TITLE:{title}")
                            try:
                                cur.execute("SELECT email FROM ContactEmails WHERE parentId=?", (cid,))
                                for (em,) in cur.fetchall():
                                    if em: v.append(f"EMAIL;TYPE=INTERNET:{em}")
                            except Exception: pass
                            try:
                                cur.execute("SELECT telephone FROM ContactTelephones WHERE parentId=?", (cid,))
                                for (tel,) in cur.fetchall():
                                    if tel: v.append(f"TEL;TYPE=VOICE:{tel}")
                            except Exception: pass
                            v.append("END:VCARD\r\n")
                            all_vcards.append("\r\n".join(v))
                        conn.close()
                    except Exception: pass

                if all_vcards:
                    out_vcf = docs_dir / 'eMClient_All_Contacts.vcf'
                    with open(out_vcf, 'w', encoding='utf-8') as f:
                        f.write("\r\n".join(all_vcards))
                    self.log(f"Exported {len(all_vcards)} contacts to {out_vcf.name}")

            self.progress_var.set(100)
            self.status_lbl.config(text="Migration successfully completed!")
            self.log("MIGRATION FINISHED SUCCESSFULLY!")
            messagebox.showinfo("Migration Complete", "All selected items have been successfully migrated!\n\nCalendar & Contacts are available in:\nDocuments/Thunderbird Migration/")

        except Exception as e:
            self.log(f"ERROR: {e}")
            messagebox.showerror("Migration Error", f"An error occurred during migration:\n{e}")
        finally:
            self.is_migrating = False
            self.start_btn.config(state=tk.NORMAL)

def main():
    root = tk.Tk()
    app = MigratorApp(root)
    root.mainloop()

if __name__ == '__main__':
    main()
