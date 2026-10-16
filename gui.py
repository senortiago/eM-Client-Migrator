#!/usr/bin/env python3
"""
eM Client Universal Migrator - Polished Graphical User Interface
=================================================================
A modern, high-DPI aware desktop UI for migrating Accounts, Emails, Subfolders,
Calendars, and Contacts from eM Client to Mozilla Thunderbird, Microsoft Outlook,
or Standalone Archives (.eml / .mbox / .ics / .vcf).
"""

import os
import sys
import re
import sqlite3
import threading
import queue
import datetime
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# Enable High-DPI scaling on Windows for razor-sharp typography
if sys.platform == 'win32':
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

from migrator import (
    EMClientReader,
    ThunderbirdTarget,
    OutlookTarget,
    HAS_WIN32COM,
    clean_filename
)

# Color Palette (Modern Slate & Royal Blue)
PALETTE = {
    "bg_canvas": "#F1F5F9",       # Slate 100
    "bg_card": "#FFFFFF",         # Pure White
    "bg_header": "#0F172A",       # Slate 900
    "bg_badge": "#1E293B",        # Slate 800
    "fg_header": "#F8FAFC",       # Slate 50
    "fg_sub": "#94A3B8",          # Slate 400
    "fg_primary": "#0F172A",      # Slate 900
    "fg_secondary": "#475569",    # Slate 600
    "fg_muted": "#94A3B8",        # Slate 400
    "border": "#CBD5E1",          # Slate 300
    "border_light": "#E2E8F0",    # Slate 200
    "accent": "#2563EB",          # Royal Blue 600
    "accent_hover": "#1D4ED8",    # Royal Blue 700
    "accent_soft": "#EFF6FF",     # Blue 50
    "accent_cyan": "#38BDF8",     # Sky 400
    "success": "#16A34A",         # Green 600
    "term_bg": "#0F172A",         # Terminal dark background
    "term_fg": "#E2E8F0",         # Terminal text
}


class TargetCard(tk.Frame):
    """A clickable modern selection card for choosing the migration destination."""
    def __init__(self, parent, value, icon, title, subtitle, variable, command):
        super().__init__(parent, bg=PALETTE["bg_card"], highlightthickness=2, cursor="hand2", padx=14, pady=10)
        self.value = value
        self.variable = variable
        self.command = command

        self.top_row = tk.Frame(self, bg=PALETTE["bg_card"])
        self.top_row.pack(fill=tk.X, anchor=tk.W)

        self.icon_lbl = tk.Label(self.top_row, text=icon, font=("Segoe UI Emoji", 14), bg=PALETTE["bg_card"], fg=PALETTE["accent"])
        self.icon_lbl.pack(side=tk.LEFT, padx=(0, 6))

        self.title_lbl = tk.Label(self.top_row, text=title, font=("Segoe UI", 10, "bold"), bg=PALETTE["bg_card"], fg=PALETTE["fg_primary"])
        self.title_lbl.pack(side=tk.LEFT)

        self.radio_indicator = tk.Label(self.top_row, text="●", font=("Segoe UI", 11), bg=PALETTE["bg_card"], fg=PALETTE["accent"])
        self.radio_indicator.pack(side=tk.RIGHT)

        self.sub_lbl = tk.Label(self, text=subtitle, font=("Segoe UI", 8), bg=PALETTE["bg_card"], fg=PALETTE["fg_secondary"], anchor=tk.W, justify=tk.LEFT)
        self.sub_lbl.pack(fill=tk.X, pady=(4, 0))

        for widget in (self, self.top_row, self.icon_lbl, self.title_lbl, self.radio_indicator, self.sub_lbl):
            widget.bind("<Button-1>", self.on_click)

    def on_click(self, _event=None):
        self.variable.set(self.value)
        if self.command:
            self.command()

    def update_visual(self):
        selected = (self.variable.get() == self.value)
        bg = PALETTE["accent_soft"] if selected else PALETTE["bg_card"]
        border = PALETTE["accent"] if selected else PALETTE["border_light"]
        ind_text = "◉" if selected else "○"
        ind_fg = PALETTE["accent"] if selected else PALETTE["fg_muted"]

        self.config(bg=bg, highlightbackground=border, highlightcolor=border)
        self.top_row.config(bg=bg)
        self.icon_lbl.config(bg=bg)
        self.title_lbl.config(bg=bg)
        self.sub_lbl.config(bg=bg)
        self.radio_indicator.config(bg=bg, text=ind_text, fg=ind_fg)


class MigratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("eM Client Universal Migrator")
        self.root.geometry("1020x800")
        self.root.minsize(900, 680)
        self.root.configure(bg=PALETTE["bg_canvas"])

        self.reader = None
        self.data_tree = {}
        self.account_types_map = {}
        self.log_queue = queue.Queue()
        self.is_migrating = False

        self.item_checked = {}
        self.item_type = {}
        self.item_data = {}

        self.configure_styles()
        self.setup_ui()

        self.root.after(80, self.initial_load)
        self.root.after(100, self.process_log_queue)

    def configure_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except Exception:
            pass

        # General Frame & Label styles
        self.style.configure("Card.TFrame", background=PALETTE["bg_card"])
        self.style.configure("Canvas.TFrame", background=PALETTE["bg_canvas"])
        self.style.configure("Card.TLabel", background=PALETTE["bg_card"], foreground=PALETTE["fg_primary"], font=("Segoe UI", 9))
        self.style.configure("Muted.TLabel", background=PALETTE["bg_card"], foreground=PALETTE["fg_secondary"], font=("Segoe UI", 9))
        self.style.configure("Section.TLabel", background=PALETTE["bg_card"], foreground=PALETTE["fg_primary"], font=("Segoe UI", 10, "bold"))

        # Checkbuttons on white card
        self.style.configure(
            "Card.TCheckbutton",
            background=PALETTE["bg_card"],
            foreground=PALETTE["fg_primary"],
            font=("Segoe UI", 9)
        )
        self.style.map("Card.TCheckbutton", background=[("active", PALETTE["bg_card"])])

        # Secondary Buttons
        self.style.configure(
            "Secondary.TButton",
            background="#F8FAFC",
            foreground=PALETTE["fg_primary"],
            bordercolor=PALETTE["border"],
            lightcolor="#F8FAFC",
            darkcolor="#F8FAFC",
            font=("Segoe UI", 9),
            padding=(10, 4)
        )
        self.style.map(
            "Secondary.TButton",
            background=[("active", "#E2E8F0"), ("disabled", "#F1F5F9")],
            foreground=[("disabled", PALETTE["fg_muted"])]
        )

        # Primary CTA Button
        self.style.configure(
            "Primary.TButton",
            background=PALETTE["accent"],
            foreground="#FFFFFF",
            bordercolor=PALETTE["accent"],
            lightcolor=PALETTE["accent"],
            darkcolor=PALETTE["accent"],
            font=("Segoe UI", 10, "bold"),
            padding=(18, 8)
        )
        self.style.map(
            "Primary.TButton",
            background=[("active", PALETTE["accent_hover"]), ("disabled", "#94A3B8")],
            foreground=[("disabled", "#E2E8F0")]
        )

        # Treeview Styling
        self.style.configure(
            "Migrator.Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=PALETTE["fg_primary"],
            rowheight=28,
            font=("Segoe UI", 9),
            borderwidth=1,
            relief="solid"
        )
        self.style.configure(
            "Migrator.Treeview.Heading",
            background="#F1F5F9",
            foreground="#334155",
            font=("Segoe UI", 9, "bold"),
            padding=(8, 6),
            relief="flat"
        )
        self.style.map(
            "Migrator.Treeview",
            background=[("selected", PALETTE["accent_soft"])],
            foreground=[("selected", PALETTE["accent"])]
        )

        # Progressbar
        self.style.configure(
            "Accent.Horizontal.TProgressbar",
            troughcolor="#E2E8F0",
            background=PALETTE["accent"],
            bordercolor="#E2E8F0",
            lightcolor=PALETTE["accent"],
            darkcolor=PALETTE["accent"],
            thickness=10
        )

    def setup_ui(self):
        # =====================================================================
        # 1. TOP DARK HEADER BANNER WITH LIVE KPI PILLS
        # =====================================================================
        header = tk.Frame(self.root, bg=PALETTE["bg_header"], padx=20, pady=14)
        header.pack(fill=tk.X)

        title_col = tk.Frame(header, bg=PALETTE["bg_header"])
        title_col.pack(side=tk.LEFT, fill=tk.Y)

        tk.Label(
            title_col,
            text="eM Client Universal Migrator",
            font=("Segoe UI", 16, "bold"),
            bg=PALETTE["bg_header"],
            fg=PALETTE["fg_header"]
        ).pack(anchor=tk.W)

        tk.Label(
            title_col,
            text="Seamless offline transfer of Accounts, Mails, Subfolders, Calendars & Contacts",
            font=("Segoe UI", 9),
            bg=PALETTE["bg_header"],
            fg=PALETTE["fg_sub"]
        ).pack(anchor=tk.W, pady=(2, 0))

        # KPI Metric Pills on the right of the header
        kpi_frame = tk.Frame(header, bg=PALETTE["bg_header"])
        kpi_frame.pack(side=tk.RIGHT)

        self.kpi_accounts_lbl = self.create_kpi_pill(kpi_frame, "ACCOUNTS", "0")
        self.kpi_mails_lbl = self.create_kpi_pill(kpi_frame, "EMAILS", "0")
        self.kpi_cal_lbl = self.create_kpi_pill(kpi_frame, "EVENTS", "0")
        self.kpi_contacts_lbl = self.create_kpi_pill(kpi_frame, "CONTACTS", "0")

        # Main Padded Container
        main_container = tk.Frame(self.root, bg=PALETTE["bg_canvas"], padx=16, pady=10)
        main_container.pack(fill=tk.BOTH, expand=True)

        # =====================================================================
        # 2. STEP 1: DESTINATION & PATHS CARD
        # =====================================================================
        dest_card = tk.Frame(main_container, bg=PALETTE["bg_card"], highlightbackground=PALETTE["border"], highlightthickness=1, padx=16, pady=12)
        dest_card.pack(fill=tk.X, pady=(0, 8))

        tk.Label(dest_card, text="1. Select Migration Target & Paths", font=("Segoe UI", 10, "bold"), bg=PALETTE["bg_card"], fg=PALETTE["fg_primary"]).pack(anchor=tk.W, pady=(0, 8))

        # 3 Interactive Target Cards
        cards_row = tk.Frame(dest_card, bg=PALETTE["bg_card"])
        cards_row.pack(fill=tk.X, pady=(0, 10))
        cards_row.columnconfigure(0, weight=1)
        cards_row.columnconfigure(1, weight=1)
        cards_row.columnconfigure(2, weight=1)

        self.target_var = tk.StringVar(value="thunderbird")
        self.target_cards = []

        c_tb = TargetCard(
            cards_row, "thunderbird", "🦅", "Mozilla Thunderbird",
            "Auto-configures accounts, MBOX mail stores & Unified Folders",
            self.target_var, self.on_target_changed
        )
        c_tb.grid(row=0, column=0, sticky=tk.EW, padx=(0, 6))
        self.target_cards.append(c_tb)

        c_ol = TargetCard(
            cards_row, "outlook", "📘", "Microsoft Outlook",
            "Creates top-level .PST data files per account via MAPI COM",
            self.target_var, self.on_target_changed
        )
        c_ol.grid(row=0, column=1, sticky=tk.EW, padx=6)
        self.target_cards.append(c_ol)

        c_ex = TargetCard(
            cards_row, "export", "📦", "Standalone Archive",
            "Exports standard .EML, .ICS & .VCF files to local disk",
            self.target_var, self.on_target_changed
        )
        c_ex.grid(row=0, column=2, sticky=tk.EW, padx=(6, 0))
        self.target_cards.append(c_ex)

        # Paths Grid
        paths_grid = tk.Frame(dest_card, bg=PALETTE["bg_card"])
        paths_grid.pack(fill=tk.X)
        paths_grid.columnconfigure(1, weight=1)

        tk.Label(paths_grid, text="Source eM Client Data:", font=("Segoe UI", 9, "bold"), bg=PALETTE["bg_card"], fg=PALETTE["fg_secondary"], width=20, anchor=tk.W).grid(row=0, column=0, sticky=tk.W, pady=3)
        self.em_path_var = tk.StringVar()
        ttk.Entry(paths_grid, textvariable=self.em_path_var, font=("Segoe UI", 9)).grid(row=0, column=1, sticky=tk.EW, padx=6, pady=3)
        ttk.Button(paths_grid, text="Browse...", style="Secondary.TButton", command=self.browse_em_path).grid(row=0, column=2, pady=3)

        self.dest_path_label = tk.Label(paths_grid, text="Thunderbird Profile:", font=("Segoe UI", 9, "bold"), bg=PALETTE["bg_card"], fg=PALETTE["fg_secondary"], width=20, anchor=tk.W)
        self.dest_path_label.grid(row=1, column=0, sticky=tk.W, pady=3)

        self.dest_path_var = tk.StringVar()
        self.dest_entry = ttk.Entry(paths_grid, textvariable=self.dest_path_var, font=("Segoe UI", 9))
        self.dest_entry.grid(row=1, column=1, sticky=tk.EW, padx=6, pady=3)
        self.dest_browse_btn = ttk.Button(paths_grid, text="Browse...", style="Secondary.TButton", command=self.browse_dest_path)
        self.dest_browse_btn.grid(row=1, column=2, pady=3)

        # =====================================================================
        # 3. STEP 2: COMPONENTS BAR
        # =====================================================================
        comp_card = tk.Frame(main_container, bg=PALETTE["bg_card"], highlightbackground=PALETTE["border"], highlightthickness=1, padx=16, pady=10)
        comp_card.pack(fill=tk.X, pady=(0, 8))

        comp_row = tk.Frame(comp_card, bg=PALETTE["bg_card"])
        comp_row.pack(fill=tk.X)

        tk.Label(comp_row, text="2. Components:", font=("Segoe UI", 10, "bold"), bg=PALETTE["bg_card"], fg=PALETTE["fg_primary"]).pack(side=tk.LEFT, padx=(0, 16))

        self.inc_accounts_var = tk.BooleanVar(value=True)
        self.inc_mails_var = tk.BooleanVar(value=True)
        self.inc_cal_var = tk.BooleanVar(value=True)
        self.inc_contacts_var = tk.BooleanVar(value=True)

        self.chk_accounts = ttk.Checkbutton(comp_row, text="Account Settings (IMAP/POP3/SMTP/OAuth2)", variable=self.inc_accounts_var, style="Card.TCheckbutton", command=self.update_selection_badge)
        self.chk_accounts.pack(side=tk.LEFT, padx=(0, 18))

        ttk.Checkbutton(comp_row, text="Emails & Folder Hierarchy", variable=self.inc_mails_var, style="Card.TCheckbutton", command=self.update_selection_badge).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Checkbutton(comp_row, text="Calendars (.ics)", variable=self.inc_cal_var, style="Card.TCheckbutton", command=self.update_selection_badge).pack(side=tk.LEFT, padx=(0, 18))
        ttk.Checkbutton(comp_row, text="Contacts (.vcf)", variable=self.inc_contacts_var, style="Card.TCheckbutton", command=self.update_selection_badge).pack(side=tk.LEFT)

        # =====================================================================
        # 4. STEP 3: INTERACTIVE ACCOUNT & FOLDER EXPLORER
        # =====================================================================
        tree_card = tk.Frame(main_container, bg=PALETTE["bg_card"], highlightbackground=PALETTE["border"], highlightthickness=1, padx=16, pady=12)
        tree_card.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Toolbar Row
        toolbar = tk.Frame(tree_card, bg=PALETTE["bg_card"])
        toolbar.pack(fill=tk.X, pady=(0, 8))

        tk.Label(toolbar, text="3. Accounts & Folders", font=("Segoe UI", 10, "bold"), bg=PALETTE["bg_card"], fg=PALETTE["fg_primary"]).pack(side=tk.LEFT, padx=(0, 12))

        ttk.Button(toolbar, text="☑ Select All", style="Secondary.TButton", command=self.select_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="☐ Deselect All", style="Secondary.TButton", command=self.deselect_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="⊞ Expand All", style="Secondary.TButton", command=self.expand_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="⊟ Collapse All", style="Secondary.TButton", command=self.collapse_all).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="↻ Refresh", style="Secondary.TButton", command=self.reload_tree).pack(side=tk.LEFT, padx=(2, 12))

        # Hide Empty Folders Toggle & Search Filter on the right of toolbar
        self.hide_empty_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            toolbar, text="Hide empty folders (0 mails)",
            variable=self.hide_empty_var, style="Card.TCheckbutton",
            command=self.populate_tree_nodes
        ).pack(side=tk.RIGHT, padx=(8, 0))

        search_frame = tk.Frame(toolbar, bg=PALETTE["bg_card"])
        search_frame.pack(side=tk.RIGHT)
        tk.Label(search_frame, text="🔍", bg=PALETTE["bg_card"], fg=PALETTE["fg_secondary"]).pack(side=tk.LEFT, padx=(0, 3))
        self.filter_var = tk.StringVar()
        self.filter_var.trace_add("write", lambda *_: self.populate_tree_nodes())
        filter_entry = ttk.Entry(search_frame, textvariable=self.filter_var, width=22, font=("Segoe UI", 9))
        filter_entry.pack(side=tk.LEFT)

        # Treeview + Scrollbar Container
        tree_container = tk.Frame(tree_card, bg=PALETTE["bg_card"])
        tree_container.pack(fill=tk.BOTH, expand=True)

        tree_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.tree = ttk.Treeview(
            tree_container,
            columns=("type", "count"),
            yscrollcommand=tree_scroll.set,
            selectmode="browse",
            style="Migrator.Treeview"
        )
        self.tree.heading("#0", text="Account / Folder Hierarchy", anchor=tk.W)
        self.tree.heading("type", text="Category", anchor=tk.W)
        self.tree.heading("count", text="Items", anchor=tk.E)

        self.tree.column("#0", stretch=True, width=540)
        self.tree.column("type", stretch=False, width=160, anchor=tk.W)
        self.tree.column("count", stretch=False, width=140, anchor=tk.E)

        self.tree.tag_configure("account", font=("Segoe UI", 9, "bold"), foreground="#0F172A", background="#F8FAFC")
        self.tree.tag_configure("folder", font=("Segoe UI", 9), foreground="#334155")
        self.tree.tag_configure("empty", font=("Segoe UI", 9), foreground="#94A3B8")

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.config(command=self.tree.yview)

        self.tree.bind("<Button-1>", self.on_tree_click)

        # =====================================================================
        # 5. FOOTER: PROGRESS, SELECTION TELEMETRY, CTA & DARK CONSOLE
        # =====================================================================
        footer_card = tk.Frame(main_container, bg=PALETTE["bg_card"], highlightbackground=PALETTE["border"], highlightthickness=1, padx=16, pady=10)
        footer_card.pack(fill=tk.X)

        # Action Bar (Status + Selection Counter + Start Button)
        action_row = tk.Frame(footer_card, bg=PALETTE["bg_card"])
        action_row.pack(fill=tk.X, pady=(0, 8))

        status_col = tk.Frame(action_row, bg=PALETTE["bg_card"])
        status_col.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.selection_badge_lbl = tk.Label(
            status_col,
            text="Selected: 0 Accounts · 0 Folders · 0 Emails",
            font=("Segoe UI", 9, "bold"),
            bg=PALETTE["bg_card"],
            fg=PALETTE["accent"]
        )
        self.selection_badge_lbl.pack(anchor=tk.W)

        self.status_lbl = tk.Label(
            status_col,
            text="Ready to migrate.",
            font=("Segoe UI", 9),
            bg=PALETTE["bg_card"],
            fg=PALETTE["fg_secondary"]
        )
        self.status_lbl.pack(anchor=tk.W)

        self.start_btn = ttk.Button(
            action_row,
            text="▶  Start Migration",
            style="Primary.TButton",
            command=self.start_migration
        )
        self.start_btn.pack(side=tk.RIGHT, padx=(12, 0))

        # Progress Bar
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            footer_card,
            variable=self.progress_var,
            maximum=100,
            style="Accent.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 8))

        # Dark Terminal Console
        term_frame = tk.Frame(footer_card, bg=PALETTE["term_bg"], padx=8, pady=6)
        term_frame.pack(fill=tk.X)

        self.log_text = tk.Text(
            term_frame,
            height=5,
            wrap=tk.WORD,
            font=("Consolas", 9),
            bg=PALETTE["term_bg"],
            fg=PALETTE["term_fg"],
            bd=0,
            highlightthickness=0,
            insertbackground=PALETTE["term_fg"]
        )
        self.log_text.pack(fill=tk.X)
        self.log_text.tag_configure("time", foreground="#38BDF8")
        self.log_text.tag_configure("success", foreground="#4ADE80")
        self.log_text.tag_configure("warn", foreground="#FBBF24")
        self.log_text.tag_configure("error", foreground="#F87171")

        # Initialize Target Cards visual state
        for card in self.target_cards:
            card.update_visual()

    def create_kpi_pill(self, parent, label_text, value_text):
        pill = tk.Frame(parent, bg=PALETTE["bg_badge"], padx=12, pady=5, highlightbackground="#334155", highlightthickness=1)
        pill.pack(side=tk.LEFT, padx=(8, 0))
        val_lbl = tk.Label(pill, text=value_text, font=("Segoe UI", 12, "bold"), bg=PALETTE["bg_badge"], fg=PALETTE["accent_cyan"])
        val_lbl.pack(anchor=tk.E)
        tk.Label(pill, text=label_text, font=("Segoe UI", 7, "bold"), bg=PALETTE["bg_badge"], fg=PALETTE["fg_sub"]).pack(anchor=tk.E)
        return val_lbl

    def on_target_changed(self):
        for card in self.target_cards:
            card.update_visual()

        target = self.target_var.get()
        if target == "thunderbird":
            self.dest_path_label.config(text="Thunderbird Profile:")
            default_tb = ThunderbirdTarget.find_default_profile()
            if default_tb:
                self.dest_path_var.set(str(default_tb))
            self.chk_accounts.config(state=tk.NORMAL)
            self.inc_accounts_var.set(True)
        elif target == "outlook":
            self.dest_path_label.config(text="Outlook PST Folder:")
            self.dest_path_var.set(str(Path(os.path.expanduser("~/Documents/Outlook Files"))))
            self.chk_accounts.config(state=tk.DISABLED)
            self.inc_accounts_var.set(False)
        else:
            self.dest_path_label.config(text="Archive Output Folder:")
            self.dest_path_var.set(str(Path(os.path.expanduser("~/Documents/Thunderbird Migration"))))
            self.chk_accounts.config(state=tk.DISABLED)
            self.inc_accounts_var.set(False)

        self.update_selection_badge()

    def browse_em_path(self):
        p = filedialog.askdirectory(title="Select eM Client Data Directory")
        if p:
            self.em_path_var.set(p)
            self.reload_tree()

    def browse_dest_path(self):
        p = filedialog.askdirectory(title="Select Destination Directory")
        if p:
            self.dest_path_var.set(p)

    def initial_load(self):
        default_em = EMClientReader.find_default_path()
        if default_em:
            self.em_path_var.set(str(default_em))

        default_tb = ThunderbirdTarget.find_default_profile()
        if default_tb:
            self.dest_path_var.set(str(default_tb))

        self.reload_tree()

    def reload_tree(self):
        em_path = self.em_path_var.get()
        if not em_path or not os.path.exists(em_path):
            return

        try:
            self.reader = EMClientReader(em_path)
            self.data_tree = self.reader.scan_data_tree()
            self.account_types_map = {
                a["email"]: f"{a['inc_type'].upper()} ({a['inc_host']})"
                for a in self.reader.accounts_config
            }
        except Exception as e:
            messagebox.showerror("Error Reading Data", f"Failed to scan eM Client databases:\n{e}")
            return

        # Update Top KPI Pills
        total_accs = len(self.data_tree)
        total_mails = sum(sum(f["count"] for f in info["folders"].values()) for info in self.data_tree.values())
        total_cal = sum(info["cal_count"] for info in self.data_tree.values())
        total_con = sum(info["contact_count"] for info in self.data_tree.values())

        self.kpi_accounts_lbl.config(text=f"{total_accs:,}")
        self.kpi_mails_lbl.config(text=f"{total_mails:,}")
        self.kpi_cal_lbl.config(text=f"{total_cal:,}")
        self.kpi_contacts_lbl.config(text=f"{total_con:,}")

        self.populate_tree_nodes()
        self.log(f"[✓] Discovered {total_accs} accounts, {total_mails:,} emails, {total_cal:,} events, and {total_con:,} contacts.")

    def populate_tree_nodes(self):
        """Populates the treeview applying current search filter and empty-folder visibility."""
        self.tree.delete(*self.tree.get_children())
        self.item_checked.clear()
        self.item_type.clear()
        self.item_data.clear()

        query = self.filter_var.get().strip().lower()
        hide_empty = self.hide_empty_var.get()

        for acc_name, info in sorted(self.data_tree.items()):
            total_mails = sum(f["count"] for f in info["folders"].values())
            if hide_empty and total_mails == 0 and info["cal_count"] == 0 and info["contact_count"] == 0:
                if acc_name not in self.account_types_map:
                    continue

            # Filter matching folders
            matching_folders = []
            for f_path, f_info in sorted(info["folders"].items()):
                if hide_empty and f_info["count"] == 0:
                    continue
                if query and (query not in acc_name.lower() and query not in f_path.lower()):
                    continue
                matching_folders.append((f_path, f_info))

            if query and not matching_folders and query not in acc_name.lower():
                continue

            summary_parts = []
            if total_mails > 0:
                summary_parts.append(f"{total_mails:,} mails")
            if info["cal_count"] > 0:
                summary_parts.append(f"{info['cal_count']:,} events")
            if info["contact_count"] > 0:
                summary_parts.append(f"{info['contact_count']:,} contacts")
            summary_str = " · ".join(summary_parts) or "Empty (Settings only)"

            acc_cat = self.account_types_map.get(acc_name, "Local Store")
            node_id = self.tree.insert(
                "", tk.END,
                text=f" ☑  ✉  {acc_name}",
                values=(acc_cat, summary_str),
                open=True,
                tags=("account",)
            )
            self.item_checked[node_id] = True
            self.item_type[node_id] = "account"
            self.item_data[node_id] = {"account": acc_name, "info": info}

            for f_path, f_info in matching_folders:
                cnt = f_info["count"]
                tag = "folder" if cnt > 0 else "empty"
                f_node = self.tree.insert(
                    node_id, tk.END,
                    text=f"    ☑  📁 {f_path}",
                    values=("Mail Folder", f"{cnt:,} mails"),
                    tags=(tag,)
                )
                self.item_checked[f_node] = True
                self.item_type[f_node] = "folder"
                self.item_data[f_node] = {"account": acc_name, "folder_path": f_path, "folder_info": f_info}

        self.update_selection_badge()

    def on_tree_click(self, event):
        # Ignore clicks on the expand/collapse arrow so expanding doesn't toggle check state
        element = self.tree.identify_element(event.x, event.y)
        if "indicator" in element:
            return

        region = self.tree.identify_region(event.x, event.y)
        if region not in ("tree", "cell"):
            return

        item = self.tree.identify_row(event.y)
        if not item:
            return

        curr = self.item_checked.get(item, True)
        new_state = not curr
        self.set_item_check(item, new_state)

        if self.item_type.get(item) == "account":
            for child in self.tree.get_children(item):
                self.set_item_check(child, new_state)
        elif self.item_type.get(item) == "folder":
            parent = self.tree.parent(item)
            if parent:
                any_checked = any(self.item_checked.get(c, False) for c in self.tree.get_children(parent))
                self.set_item_check(parent, any_checked)

        self.update_selection_badge()

    def set_item_check(self, item, checked):
        self.item_checked[item] = checked
        raw_text = self.tree.item(item, "text")
        clean_text = raw_text.replace("☑", "").replace("☐", "").strip()
        mark = "☑" if checked else "☐"
        if self.item_type.get(item) == "folder":
            self.tree.item(item, text=f"    {mark}  {clean_text}")
        else:
            self.tree.item(item, text=f" {mark}  {clean_text}")

    def select_all(self):
        for item in self.tree.get_children():
            self.set_item_check(item, True)
            for child in self.tree.get_children(item):
                self.set_item_check(child, True)
        self.update_selection_badge()

    def deselect_all(self):
        for item in self.tree.get_children():
            self.set_item_check(item, False)
            for child in self.tree.get_children(item):
                self.set_item_check(child, False)
        self.update_selection_badge()

    def expand_all(self):
        for item in self.tree.get_children():
            self.tree.item(item, open=True)

    def collapse_all(self):
        for item in self.tree.get_children():
            self.tree.item(item, open=False)

    def update_selection_badge(self):
        sel_accs = set()
        sel_folders = 0
        sel_mails = 0

        for item, checked in self.item_checked.items():
            if not checked:
                continue
            itype = self.item_type.get(item)
            if itype == "account":
                sel_accs.add(self.item_data[item]["account"])
            elif itype == "folder":
                sel_folders += 1
                sel_mails += self.item_data[item]["folder_info"]["count"]
                sel_accs.add(self.item_data[item]["account"])

        self.selection_badge_lbl.config(
            text=f"Selected: {len(sel_accs)} Accounts · {sel_folders} Folders · {sel_mails:,} Emails"
        )

    def log(self, msg):
        self.log_queue.put(msg)

    def process_log_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            t = datetime.datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{t}] ", "time")

            tag = None
            if "[✓]" in msg or "SUCCESS" in msg:
                tag = "success"
            elif "[!]" in msg or "Warning" in msg:
                tag = "warn"
            elif "ERROR" in msg or "[✗]" in msg:
                tag = "error"

            if tag:
                self.log_text.insert(tk.END, f"{msg}\n", tag)
            else:
                self.log_text.insert(tk.END, f"{msg}\n")
            self.log_text.see(tk.END)
        self.root.after(100, self.process_log_queue)

    def start_migration(self):
        if self.is_migrating:
            return

        target = self.target_var.get()
        dest_path = self.dest_path_var.get().strip()

        if target == "thunderbird":
            if not dest_path or not os.path.exists(dest_path):
                messagebox.showerror("Invalid Path", "Please select a valid Mozilla Thunderbird profile directory.")
                return
            lock_f = Path(dest_path) / "parent.lock"
            if lock_f.exists():
                res = messagebox.askyesno(
                    "Thunderbird Running",
                    "Mozilla Thunderbird appears to be open.\n\n"
                    "Would you like to automatically close Thunderbird now so accounts and mail stores can be safely updated?"
                )
                if res:
                    if sys.platform == "win32":
                        subprocess.run(["powershell", "-Command", "Get-Process -Name thunderbird -ErrorAction SilentlyContinue | Stop-Process -Force"], capture_output=True)
                    try:
                        if lock_f.exists():
                            lock_f.unlink()
                    except Exception:
                        pass
                else:
                    return
        elif target == "outlook":
            if not HAS_WIN32COM:
                messagebox.showerror(
                    "Outlook Dependency Missing",
                    "Microsoft Outlook MAPI automation requires 'pywin32'.\n\n"
                    "Run: pip install pywin32\nOr choose Standalone Archive."
                )
                return

        selected_folders = []
        selected_accounts = set()
        for item, checked in self.item_checked.items():
            if not checked:
                continue
            if self.item_type.get(item) == "account":
                selected_accounts.add(self.item_data[item]["account"])
            elif self.item_type.get(item) == "folder":
                data = self.item_data[item]
                selected_folders.append((data["account"], data["folder_path"]))
                selected_accounts.add(data["account"])

        if not selected_accounts and not selected_folders and not (self.inc_cal_var.get() or self.inc_contacts_var.get()):
            messagebox.showwarning("Nothing Selected", "Please select at least one account, folder, calendar, or contact set to migrate.")
            return

        self.is_migrating = True
        self.start_btn.config(state=tk.DISABLED)
        self.progress_var.set(0)

        thread = threading.Thread(
            target=self.run_migration_thread,
            args=(target, dest_path, selected_accounts, selected_folders),
            daemon=True
        )
        thread.start()

    def run_migration_thread(self, target, dest_path, selected_accounts, selected_folders):
        try:
            self.log(f"Starting migration to target: {target.upper()}")
            total_steps = max(len(selected_folders) + 2, 1)
            completed_steps = 0

            # 1. Account Settings (Thunderbird)
            if target == "thunderbird" and self.inc_accounts_var.get():
                self.status_lbl.config(text="Configuring accounts & server settings in Thunderbird...")
                tb = ThunderbirdTarget(dest_path)
                bak = tb.backup_prefs()
                if bak:
                    self.log(f"Backed up prefs.js -> {bak.name}")
                accs_to_add = [a for a in self.reader.accounts_config if a["email"] in selected_accounts]
                tb.deploy_accounts(accs_to_add)
                self.log(f"[✓] Configured {len(accs_to_add)} accounts in Thunderbird prefs.js.")

            completed_steps += 1
            self.progress_var.set((completed_steps / total_steps) * 100)

            # 2. Emails & Folders Migration
            total_migrated_msgs = 0
            if self.inc_mails_var.get() and selected_folders:
                if target == "thunderbird":
                    tb = ThunderbirdTarget(dest_path)
                    for acc_name, folder_path in selected_folders:
                        self.status_lbl.config(text=f"Migrating {acc_name} -> {folder_path}...")
                        info = self.data_tree.get(acc_name, {})
                        target_dir = tb.get_account_mbox_dir(acc_name)

                        parts = folder_path.split("/")
                        curr = target_dir
                        for p in parts[:-1]:
                            (curr / p).touch(exist_ok=True)
                            curr = curr / f"{p}.sbd"
                            curr.mkdir(parents=True, exist_ok=True)

                        raw_name = parts[-1]
                        norm_name = raw_name
                        if raw_name.upper() in ("INBOX", "ΕΙΣΕΡΧΟΜΕΝΑ", "ΌΛΑ ΤΑ ΜΗΝΎΜΑΤΑ", "ALL MAIL"):
                            norm_name = "INBOX"
                        elif raw_name.upper() in ("SENT", "ΑΠΕΣΤΑΛΜΈΝΑ", "SENT MAIL"):
                            norm_name = "Sent"
                        elif raw_name.upper() in ("TRASH", "DELETED", "ΚΆΔΟΣ ΑΠΟΡΡΙΜΜΆΤΩΝ", "ΔΙΑΓΡΑΜΜΈΝΑ"):
                            norm_name = "Trash"
                        elif raw_name.upper() in ("JUNK", "SPAM", "BULK", "ΑΝΕΠΙΘΎΜΗΤΑ"):
                            norm_name = "Junk"

                        mbox_path = curr / norm_name
                        count = 0
                        with open(mbox_path, "ab") as f_out:
                            for ad in info.get("acc_dirs", []):
                                for msg in self.reader.get_messages_generator(ad, folder_path):
                                    tb.write_mbox_message(f_out, msg)
                                    count += 1
                                    total_migrated_msgs += 1

                        msf_path = curr / f"{norm_name}.msf"
                        if not msf_path.exists():
                            msf_path.touch()

                        self.log(f"[✓] {acc_name} [{folder_path}]: {count:,} emails written to MBOX.")
                        completed_steps += 1
                        self.progress_var.set((completed_steps / total_steps) * 100)

                    self.status_lbl.config(text="Updating Thunderbird Unified Folders & folderCache...")
                    tb.update_unified_folders()
                    self.log("[✓] Unified Folders (virtualFolders.dat & folderCache.json) synchronized.")

                elif target == "outlook":
                    ot = OutlookTarget(mode="pst", output_dir=dest_path)
                    ot.connect()
                    for acc_name, folder_path in selected_folders:
                        self.status_lbl.config(text=f"Importing into Outlook PST: {acc_name} -> {folder_path}...")
                        root_store = ot.get_or_create_store(acc_name)
                        parts = folder_path.split("/")
                        target_f = ot.get_target_folder(root_store, parts)

                        info = self.data_tree.get(acc_name, {})
                        count = 0
                        for ad in info.get("acc_dirs", []):
                            for msg in self.reader.get_messages_generator(ad, folder_path):
                                ot.import_message(target_f, msg)
                                count += 1
                                total_migrated_msgs += 1
                        self.log(f"[✓] {acc_name} [{folder_path}]: {count:,} emails imported into Outlook PST.")
                        completed_steps += 1
                        self.progress_var.set((completed_steps / total_steps) * 100)

                elif target == "export":
                    out_docs = Path(dest_path) / "Exported Mails"
                    out_docs.mkdir(parents=True, exist_ok=True)
                    for acc_name, folder_path in selected_folders:
                        self.status_lbl.config(text=f"Exporting EML: {acc_name} -> {folder_path}...")
                        info = self.data_tree.get(acc_name, {})
                        target_dir = out_docs / clean_filename(acc_name) / folder_path.replace("/", os.sep)
                        target_dir.mkdir(parents=True, exist_ok=True)
                        count = 0
                        for ad in info.get("acc_dirs", []):
                            for msg in self.reader.get_messages_generator(ad, folder_path):
                                safe_subj = clean_filename(msg["subject"] or f"msg_{msg['id']}")[:50]
                                (target_dir / f"{msg['id']}_{safe_subj}.eml").write_bytes(msg["eml_bytes"])
                                count += 1
                                total_migrated_msgs += 1
                        self.log(f"[✓] {acc_name} [{folder_path}]: {count:,} .eml files exported.")
                        completed_steps += 1
                        self.progress_var.set((completed_steps / total_steps) * 100)

            # 3. Calendar & Contacts Export
            docs_dir = Path(os.path.expanduser("~/Documents/Thunderbird Migration"))
            docs_dir.mkdir(parents=True, exist_ok=True)

            if self.inc_cal_var.get():
                self.status_lbl.config(text="Exporting Calendars (.ics)...")
                all_events = []
                for db in self.reader.data_path.glob("**/event_data.dat"):
                    try:
                        conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
                        cur = conn.cursor()
                        cur.execute("SELECT properties FROM EventItemContents WHERE propertyName='calendar'")
                        for (props,) in cur.fetchall():
                            text = props.decode("utf-8", errors="replace") if isinstance(props, bytes) else str(props)
                            m = re.search(r"(BEGIN:VEVENT[\s\S]*?END:VEVENT)", text)
                            if m:
                                all_events.append(m.group(1))
                        conn.close()
                    except Exception:
                        pass

                if all_events:
                    out_ics = docs_dir / "eMClient_All_Calendars.ics"
                    with open(out_ics, "w", encoding="utf-8") as f:
                        f.write("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Universal Migrator//EN\r\n")
                        for ev in all_events:
                            f.write(ev + "\r\n")
                        f.write("END:VCALENDAR\r\n")
                    self.log(f"[✓] Exported {len(all_events):,} calendar events -> {out_ics}")

            if self.inc_contacts_var.get():
                self.status_lbl.config(text="Exporting Contacts (.vcf)...")
                all_vcards = []
                for db in self.reader.data_path.glob("**/contact_index.dat"):
                    try:
                        conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
                        cur = conn.cursor()
                        cur.execute("SELECT id, firstName, lastName, displayName, organization, title FROM ContactItems")
                        for cid, fn, ln, dn, org, title in cur.fetchall():
                            v = ["BEGIN:VCARD", "VERSION:3.0", f"N:{ln or ''};{fn or ''};;;", f"FN:{dn or fn or ln or ''}"]
                            if org:
                                v.append(f"ORG:{org}")
                            if title:
                                v.append(f"TITLE:{title}")
                            try:
                                cur.execute("SELECT email FROM ContactEmails WHERE parentId=?", (cid,))
                                for (em,) in cur.fetchall():
                                    if em:
                                        v.append(f"EMAIL;TYPE=INTERNET:{em}")
                            except Exception:
                                pass
                            try:
                                cur.execute("SELECT telephone FROM ContactTelephones WHERE parentId=?", (cid,))
                                for (tel,) in cur.fetchall():
                                    if tel:
                                        v.append(f"TEL;TYPE=VOICE:{tel}")
                            except Exception:
                                pass
                            v.append("END:VCARD\r\n")
                            all_vcards.append("\r\n".join(v))
                        conn.close()
                    except Exception:
                        pass

                if all_vcards:
                    out_vcf = docs_dir / "eMClient_All_Contacts.vcf"
                    with open(out_vcf, "w", encoding="utf-8") as f:
                        f.write("\r\n".join(all_vcards))
                    self.log(f"[✓] Exported {len(all_vcards):,} contacts -> {out_vcf}")

            self.progress_var.set(100)
            self.status_lbl.config(text=f"Migration complete! ({total_migrated_msgs:,} emails transferred)")
            self.log("[✓] MIGRATION FINISHED SUCCESSFULLY!")
            messagebox.showinfo(
                "Migration Complete",
                f"Successfully migrated {total_migrated_msgs:,} emails across {len(selected_accounts)} accounts!\n\n"
                f"Calendar (.ics) & Contacts (.vcf) saved in:\n{docs_dir}"
            )

        except Exception as e:
            self.log(f"[✗] ERROR: {e}")
            messagebox.showerror("Migration Error", f"An error occurred during migration:\n{e}")
        finally:
            self.is_migrating = False
            self.start_btn.config(state=tk.NORMAL)


def main():
    root = tk.Tk()
    MigratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
