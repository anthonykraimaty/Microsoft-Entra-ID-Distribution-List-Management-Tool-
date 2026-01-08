#!/usr/bin/env python3
"""Graphical User Interface for Distribution List Manager."""

import customtkinter as ctk
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import sys
from typing import Optional, Callable
from dataclasses import dataclass

from distribution_list_manager import DistributionListManager, DistributionList, Member
from exchange_client import ExchangeClient

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S'
)
logger = logging.getLogger(__name__)

# Set appearance - Black theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("green")  # Green accent on black

# Custom black colors
COLORS = {
    "bg_dark": "#0a0a0a",
    "bg_medium": "#141414",
    "bg_light": "#1e1e1e",
    "accent": "#00aa55",
    "accent_hover": "#00cc66",
    "selection": "#2d5a3d",  # Lighter/muted green for row selection
    "text": "#ffffff",
    "text_dim": "#888888",
    "error": "#ff4444",
}

print("=" * 50)
print("  Distribution List Manager - Starting...")
print("=" * 50)


def set_title_bar_color(window, color):
    """Set the Windows title bar color (Windows 10/11 only)."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
        from ctypes import wintypes

        # Get the window handle
        hwnd = ctypes.windll.user32.GetParent(window.winfo_id())

        # DWMWA_CAPTION_COLOR = 35 (Windows 11) or use dark mode attribute
        DWMWA_CAPTION_COLOR = 35
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20

        # Convert hex color to BGR integer (Windows uses BGR, not RGB)
        # color should be like "#2b2b2b"
        if color.startswith("#"):
            r = int(color[1:3], 16)
            g = int(color[3:5], 16)
            b = int(color[5:7], 16)
            bgr_color = (b << 16) | (g << 8) | r
        else:
            bgr_color = 0x2b2b2b  # Default grey

        # Try to set caption color (Windows 11)
        result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_CAPTION_COLOR,
            ctypes.byref(ctypes.c_int(bgr_color)),
            ctypes.sizeof(ctypes.c_int)
        )

        # Also enable dark mode for title bar (Windows 10/11)
        ctypes.windll.dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(ctypes.c_int(1)),
            ctypes.sizeof(ctypes.c_int)
        )
    except Exception as e:
        logger.debug(f"Could not set title bar color: {e}")


class LoadingDialog(ctk.CTkToplevel):
    """Loading indicator dialog."""

    def __init__(self, parent, message: str = "Loading..."):
        super().__init__(parent)
        self.title("")
        self.geometry("300x100")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 300) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 100) // 2
        self.geometry(f"+{x}+{y}")

        self.label = ctk.CTkLabel(self, text=message, font=ctk.CTkFont(size=14))
        self.label.pack(expand=True)

        self.progress = ctk.CTkProgressBar(self, mode="indeterminate", width=250)
        self.progress.pack(pady=10)
        self.progress.start()

    def update_message(self, message: str):
        self.label.configure(text=message)
        self.update()


class ProgressDialog(ctk.CTkToplevel):
    """Progress dialog with determinate progress bar for bulk operations."""

    def __init__(self, parent, title: str = "Processing...", total: int = 100):
        super().__init__(parent)
        self.title("")
        self.geometry("400x140")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.total = total
        self.current = 0
        self._cancelled = False

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 140) // 2
        self.geometry(f"+{x}+{y}")

        # Title label
        self.title_label = ctk.CTkLabel(self, text=title, font=ctk.CTkFont(size=14, weight="bold"))
        self.title_label.pack(pady=(15, 5))

        # Status label (e.g., "Processing item 5 of 20")
        self.status_label = ctk.CTkLabel(self, text=f"0 / {total}", font=ctk.CTkFont(size=12))
        self.status_label.pack(pady=5)

        # Progress bar
        self.progress = ctk.CTkProgressBar(self, mode="determinate", width=350)
        self.progress.pack(pady=10)
        self.progress.set(0)

        # Cancel button
        self.cancel_btn = ctk.CTkButton(
            self, text="Cancel", width=100, command=self._on_cancel,
            fg_color="gray", hover_color="darkgray"
        )
        self.cancel_btn.pack(pady=5)

        # Prevent closing
        self.protocol("WM_DELETE_WINDOW", self._on_cancel)

    def _on_cancel(self):
        """Handle cancel button."""
        self._cancelled = True
        self.cancel_btn.configure(state="disabled", text="Cancelling...")

    def is_cancelled(self) -> bool:
        """Check if operation was cancelled."""
        return self._cancelled

    def update_progress(self, current: int, status_text: str = None):
        """Update progress bar and status text. Thread-safe via self.after()."""
        self.current = current
        progress_value = current / self.total if self.total > 0 else 0
        self.progress.set(progress_value)

        if status_text:
            self.status_label.configure(text=status_text)
        else:
            self.status_label.configure(text=f"{current} / {self.total}")

        # Use update_idletasks() instead of update() to avoid recursion
        # update_idletasks() only processes display updates, not new events
        self.update_idletasks()

    def set_title(self, title: str):
        """Update the title label."""
        self.title_label.configure(text=title)
        self.update_idletasks()


class AddMemberDialog(ctk.CTkToplevel):
    """Dialog for adding a member to a distribution list."""

    def __init__(self, parent, on_submit: Callable[[str], None]):
        super().__init__(parent)
        self.title("Add Member")
        self.geometry("400x150")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.on_submit = on_submit

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 150) // 2
        self.geometry(f"+{x}+{y}")

        # Email input
        ctk.CTkLabel(self, text="Email Address:", font=ctk.CTkFont(size=14)).pack(pady=(20, 5))
        self.email_entry = ctk.CTkEntry(self, width=350, placeholder_text="user@domain.com")
        self.email_entry.pack(pady=5)
        self.email_entry.focus()

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="Add", command=self._submit, width=100).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.bind("<Return>", lambda e: self._submit())
        self.bind("<Escape>", lambda e: self.destroy())

    def _submit(self):
        email = self.email_entry.get().strip()
        if email and "@" in email:
            self.on_submit(email)
            self.destroy()
        else:
            messagebox.showerror("Invalid Email", "Please enter a valid email address.")


class AddGroupDialog(ctk.CTkToplevel):
    """Dialog for adding an existing distribution group as a member."""

    def __init__(self, parent, groups: list, current_group_id: str, on_submit: Callable[[str], None]):
        super().__init__(parent)
        self.title("Add Group as Member")
        self.geometry("450x400")
        self.resizable(False, True)
        self.transient(parent)
        self.grab_set()

        self.on_submit = on_submit
        self.groups = groups
        self.current_group_id = current_group_id
        self.filtered_groups = []

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 400) // 2
        self.geometry(f"+{x}+{y}")

        # Search input
        ctk.CTkLabel(self, text="Search Groups:", font=ctk.CTkFont(size=14)).pack(pady=(15, 5))
        self.search_entry = ctk.CTkEntry(self, width=400, placeholder_text="Type to filter groups...")
        self.search_entry.pack(pady=5, padx=20)
        self.search_entry.bind("<KeyRelease>", self._on_search)
        self.search_entry.focus()

        # Groups list
        ctk.CTkLabel(self, text="Select a group to add:", font=ctk.CTkFont(size=12)).pack(pady=(10, 5))

        # Frame for listbox with scrollbar
        list_frame = ctk.CTkFrame(self, fg_color="transparent")
        list_frame.pack(pady=5, padx=20, fill="both", expand=True)

        # Use CTkScrollableFrame for the list
        self.scroll_frame = ctk.CTkScrollableFrame(list_frame, width=380, height=200)
        self.scroll_frame.pack(fill="both", expand=True)

        self.group_buttons = []
        self.selected_group = None

        self._populate_groups()

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=15)

        self.add_btn = ctk.CTkButton(btn_frame, text="Add Group", command=self._submit, width=100)
        self.add_btn.pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.bind("<Escape>", lambda e: self.destroy())

    def _populate_groups(self, filter_text: str = ""):
        """Populate the groups list, optionally filtered."""
        # Clear existing buttons
        for btn in self.group_buttons:
            btn.destroy()
        self.group_buttons = []

        filter_lower = filter_text.lower()
        self.filtered_groups = []

        for group in self.groups:
            # Skip the current group (can't add a group to itself)
            if group.id == self.current_group_id:
                continue

            # Apply filter
            if filter_text:
                if filter_lower not in group.display_name.lower() and filter_lower not in group.mail.lower():
                    continue

            self.filtered_groups.append(group)

            # Create a button for each group
            btn = ctk.CTkButton(
                self.scroll_frame,
                text=f"{group.display_name}\n{group.mail}",
                anchor="w",
                height=50,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray70", "gray30"),
                command=lambda g=group: self._select_group(g)
            )
            btn.pack(fill="x", pady=2, padx=5)
            self.group_buttons.append(btn)

        if not self.filtered_groups:
            no_results = ctk.CTkLabel(self.scroll_frame, text="No groups found", text_color="gray")
            no_results.pack(pady=20)
            self.group_buttons.append(no_results)

    def _select_group(self, group):
        """Handle group selection."""
        self.selected_group = group

        # Update button colors to show selection
        for btn in self.group_buttons:
            if isinstance(btn, ctk.CTkButton):
                if btn.cget("text").endswith(group.mail):
                    btn.configure(fg_color=("green", "darkgreen"))
                else:
                    btn.configure(fg_color="transparent")

    def _on_search(self, event=None):
        """Handle search input."""
        self.selected_group = None
        self._populate_groups(self.search_entry.get().strip())

    def _submit(self):
        if self.selected_group:
            self.on_submit(self.selected_group.mail)
            self.destroy()
        else:
            messagebox.showwarning("No Selection", "Please select a group to add.")


class BulkAddDialog(ctk.CTkToplevel):
    """Dialog for adding multiple members."""

    def __init__(self, parent, on_submit: Callable[[list], None]):
        super().__init__(parent)
        self.title("Add Multiple Members")
        self.geometry("500x400")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self.on_submit = on_submit

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 500) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 400) // 2
        self.geometry(f"+{x}+{y}")

        # Instructions
        ctk.CTkLabel(
            self, text="Enter email addresses (one per line):",
            font=ctk.CTkFont(size=14)
        ).pack(pady=(20, 5))

        # Text area
        self.text_area = ctk.CTkTextbox(self, width=450, height=250)
        self.text_area.pack(pady=10, padx=20)
        self.text_area.focus()

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        ctk.CTkButton(btn_frame, text="Import from File", command=self._import_file,
                      width=120).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Add All", command=self._submit,
                      width=100).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.bind("<Escape>", lambda e: self.destroy())

    def _import_file(self):
        file_path = filedialog.askopenfilename(
            title="Select file with emails",
            filetypes=[
                ("All supported", "*.txt *.csv *.xlsx"),
                ("Text files", "*.txt"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
            ]
        )
        if file_path:
            try:
                emails = []
                if file_path.endswith(".csv"):
                    import pandas as pd
                    df = pd.read_csv(file_path)
                    # Try common column names
                    for col in ["email", "Email", "EMAIL", "mail", "Mail"]:
                        if col in df.columns:
                            emails = df[col].dropna().tolist()
                            break
                    if not emails and len(df.columns) > 0:
                        emails = df.iloc[:, 0].dropna().tolist()
                elif file_path.endswith(".xlsx"):
                    import pandas as pd
                    df = pd.read_excel(file_path)
                    for col in ["email", "Email", "EMAIL", "mail", "Mail"]:
                        if col in df.columns:
                            emails = df[col].dropna().tolist()
                            break
                    if not emails and len(df.columns) > 0:
                        emails = df.iloc[:, 0].dropna().tolist()
                else:
                    with open(file_path, "r") as f:
                        emails = [line.strip() for line in f if "@" in line]

                self.text_area.delete("1.0", "end")
                self.text_area.insert("1.0", "\n".join(emails))
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import file: {e}")

    def _submit(self):
        text = self.text_area.get("1.0", "end")
        emails = [line.strip() for line in text.split("\n") if line.strip() and "@" in line.strip()]
        if emails:
            self.on_submit(emails)
            self.destroy()
        else:
            messagebox.showerror("No Emails", "Please enter at least one valid email address.")


class EditListDialog(ctk.CTkToplevel):
    """Dialog for editing distribution list properties."""

    def __init__(self, parent, dl: DistributionList, on_submit: Callable[[str, str, str], None]):
        super().__init__(parent)
        self.title("Edit Distribution List")
        self.geometry("450x340")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.on_submit = on_submit
        self.current_mail = dl.mail

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 340) // 2
        self.geometry(f"+{x}+{y}")

        # Name input
        ctk.CTkLabel(self, text="Display Name:", font=ctk.CTkFont(size=14)).pack(pady=(20, 5))
        self.name_entry = ctk.CTkEntry(self, width=400)
        self.name_entry.pack(pady=5)
        self.name_entry.insert(0, dl.display_name)

        # Email input
        ctk.CTkLabel(self, text="Email Address:", font=ctk.CTkFont(size=14)).pack(pady=(10, 5))
        self.email_entry = ctk.CTkEntry(self, width=400)
        self.email_entry.pack(pady=5)
        self.email_entry.insert(0, dl.mail)

        # Description input
        ctk.CTkLabel(self, text="Description:", font=ctk.CTkFont(size=14)).pack(pady=(10, 5))
        self.desc_entry = ctk.CTkEntry(self, width=400)
        self.desc_entry.pack(pady=5)
        if dl.description:
            self.desc_entry.insert(0, dl.description)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="Save", command=self._submit, width=100,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.bind("<Escape>", lambda e: self.destroy())

    def _validate(self):
        """Validate input fields."""
        name = self.name_entry.get().strip()
        email = self.email_entry.get().strip()

        if not name:
            messagebox.showerror("Invalid Name", "Please enter a display name.")
            return None

        if not email or "@" not in email:
            messagebox.showerror("Invalid Email", "Please enter a valid email address.")
            return None

        desc = self.desc_entry.get().strip()
        # Pass full email if changed, None otherwise
        new_email = email if email != self.current_mail else None

        return name, desc, new_email

    def _submit(self):
        """Save and close."""
        result = self._validate()
        if result:
            name, desc, mail_nickname = result
            self.on_submit(name, desc, mail_nickname)
            self.destroy()


class CreateListDialog(ctk.CTkToplevel):
    """Dialog for creating a new distribution list."""

    def __init__(self, parent, on_submit: Callable[[str, str, str], None]):
        super().__init__(parent)
        self.title("Create Distribution List")
        self.geometry("450x340")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.on_submit = on_submit

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 340) // 2
        self.geometry(f"+{x}+{y}")

        # Name input
        ctk.CTkLabel(self, text="Display Name:", font=ctk.CTkFont(size=14)).pack(pady=(20, 5))
        self.name_entry = ctk.CTkEntry(self, width=400, placeholder_text="e.g. Marketing Team")
        self.name_entry.pack(pady=5)

        # Email input
        ctk.CTkLabel(self, text="Email Alias (before @):", font=ctk.CTkFont(size=14)).pack(pady=(10, 5))
        self.email_entry = ctk.CTkEntry(self, width=400, placeholder_text="e.g. marketing-team")
        self.email_entry.pack(pady=5)

        # Description input
        ctk.CTkLabel(self, text="Description (optional):", font=ctk.CTkFont(size=14)).pack(pady=(10, 5))
        self.desc_entry = ctk.CTkEntry(self, width=400, placeholder_text="e.g. Marketing department mailing list")
        self.desc_entry.pack(pady=5)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="Create", command=self._submit, width=100,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.name_entry.focus()
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: self._submit())

    def _submit(self):
        name = self.name_entry.get().strip()
        email_alias = self.email_entry.get().strip()
        desc = self.desc_entry.get().strip()

        if not name:
            messagebox.showerror("Invalid Name", "Please enter a display name.")
            return

        if not email_alias:
            messagebox.showerror("Invalid Email", "Please enter an email alias.")
            return

        # Remove @ and domain if user entered full email
        if "@" in email_alias:
            email_alias = email_alias.split("@")[0]

        # Validate alias (no special chars except hyphen and underscore)
        import re
        if not re.match(r'^[a-zA-Z0-9._-]+$', email_alias):
            messagebox.showerror("Invalid Alias", "Email alias can only contain letters, numbers, dots, hyphens, and underscores.")
            return

        self.on_submit(name, email_alias, desc)
        self.destroy()


class ConfirmDeleteDialog(ctk.CTkToplevel):
    """Dialog requiring user to type DELETE to confirm destructive operation."""

    def __init__(self, parent, message: str):
        super().__init__(parent)
        self.title("Confirm Deletion")
        self.geometry("400x180")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.confirmed = False

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 400) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 180) // 2
        self.geometry(f"+{x}+{y}")

        # Warning message
        ctk.CTkLabel(self, text=message, font=ctk.CTkFont(size=13),
                     text_color="orange", wraplength=350).pack(pady=(20, 10))

        # Entry for typing DELETE
        self.entry = ctk.CTkEntry(self, width=200, placeholder_text="Type DELETE here")
        self.entry.pack(pady=10)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=15)

        ctk.CTkButton(btn_frame, text="Confirm", command=self._confirm, width=100,
                      fg_color="#c0392b", hover_color="#e74c3c").pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.entry.focus()
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: self._confirm())

    def _confirm(self):
        if self.entry.get().strip().upper() == "DELETE":
            self.confirmed = True
            self.destroy()
        else:
            messagebox.showerror("Invalid", "You must type DELETE to confirm.")


class ErrorLogDialog(ctk.CTkToplevel):
    """Dialog showing copyable error log."""

    def __init__(self, parent, title: str, message: str, errors: list):
        super().__init__(parent)
        self.title(title)
        self.geometry("600x400")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 400) // 2
        self.geometry(f"+{x}+{y}")

        # Summary message
        ctk.CTkLabel(self, text=message, font=ctk.CTkFont(size=13),
                     wraplength=550).pack(pady=(15, 10), padx=10)

        # Scrollable text area for errors (copyable)
        self.textbox = ctk.CTkTextbox(self, width=560, height=280)
        self.textbox.pack(pady=10, padx=20, fill="both", expand=True)

        # Insert all errors
        error_text = "\n".join(errors)
        self.textbox.insert("1.0", error_text)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        ctk.CTkButton(btn_frame, text="Copy to Clipboard", command=self._copy, width=130,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Close", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.bind("<Escape>", lambda e: self.destroy())

    def _copy(self):
        self.clipboard_clear()
        self.clipboard_append(self.textbox.get("1.0", "end-1c"))
        messagebox.showinfo("Copied", "Errors copied to clipboard.")


class EditMemberDialog(ctk.CTkToplevel):
    """Dialog for editing a member's display name."""

    def __init__(self, parent, member, on_submit: Callable[[str, str], None]):
        super().__init__(parent)
        self.title("Edit Member")
        self.geometry("450x280")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.on_submit = on_submit
        self.member = member
        self.original_email = member.email

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 450) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 280) // 2
        self.geometry(f"+{x}+{y}")

        # Info label
        info_label = ctk.CTkLabel(self, text="Note: Editing removes the old member and adds the new one.",
                                   font=ctk.CTkFont(size=11), text_color="orange")
        info_label.pack(pady=(15, 5))

        # Name input
        ctk.CTkLabel(self, text="Display Name:", font=ctk.CTkFont(size=14)).pack(pady=(10, 5))
        self.name_entry = ctk.CTkEntry(self, width=400)
        self.name_entry.pack(pady=5)
        self.name_entry.insert(0, member.display_name)

        # Email input
        ctk.CTkLabel(self, text="Email Address:", font=ctk.CTkFont(size=14)).pack(pady=(10, 5))
        self.email_entry = ctk.CTkEntry(self, width=400)
        self.email_entry.pack(pady=5)
        self.email_entry.insert(0, member.email)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20)

        ctk.CTkButton(btn_frame, text="Save", command=self._submit, width=100,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"]).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack(side="left", padx=5)

        self.name_entry.focus()
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: self._submit())

    def _validate(self):
        """Validate input and return new email if valid."""
        new_email = self.email_entry.get().strip()

        if not new_email or "@" not in new_email:
            messagebox.showerror("Invalid Email", "Please enter a valid email address.")
            return None

        return new_email

    def _submit(self):
        """Save and close."""
        new_email = self._validate()
        if new_email:
            if new_email.lower() != self.original_email.lower():
                self.on_submit(self.original_email, new_email)
            self.destroy()


class SearchEmailDialog(ctk.CTkToplevel):
    """Dialog for searching which distribution lists an email belongs to."""

    def __init__(self, parent, manager, main_app):
        super().__init__(parent)
        self.title("Search Email Memberships")
        self.geometry("650x500")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()

        self.manager = manager
        self.parent = parent
        self.main_app = main_app  # Reference to main app for cache

        # Center on parent
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - 650) // 2
        y = parent.winfo_y() + (parent.winfo_height() - 500) // 2
        self.geometry(f"+{x}+{y}")

        # Search input
        search_frame = ctk.CTkFrame(self, fg_color="transparent")
        search_frame.pack(fill="x", padx=20, pady=(20, 10))

        ctk.CTkLabel(search_frame, text="Search:", font=ctk.CTkFont(size=14)).pack(side="left", padx=(0, 10))
        self.email_entry = ctk.CTkEntry(search_frame, width=350, placeholder_text="email or partial (e.g. 'john' or 'john@')")
        self.email_entry.pack(side="left", padx=(0, 10))
        self.email_entry.focus()

        self.search_btn = ctk.CTkButton(search_frame, text="Search", command=self._search, width=100)
        self.search_btn.pack(side="left")

        # Options row
        options_frame = ctk.CTkFrame(self, fg_color="transparent")
        options_frame.pack(fill="x", padx=20, pady=(0, 5))

        self.partial_match_var = ctk.BooleanVar(value=True)
        self.partial_match_cb = ctk.CTkCheckBox(
            options_frame, text="Partial match",
            variable=self.partial_match_var
        )
        self.partial_match_cb.pack(side="left")

        # Cache status label
        total_members = self.main_app._get_cached_member_count()
        self.cache_label = ctk.CTkLabel(
            options_frame,
            text=f"Cached: {len(self.main_app.members_cache)} lists, {total_members} members",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        )
        self.cache_label.pack(side="right", padx=10)

        # Results label
        self.results_label = ctk.CTkLabel(self, text="Enter an email to search (instant from cache)",
                                          font=ctk.CTkFont(size=12), text_color="gray")
        self.results_label.pack(pady=10)

        # Results list
        results_frame = ctk.CTkFrame(self, fg_color=COLORS["bg_dark"])
        results_frame.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)

        # Style for treeview
        style = ttk.Style()
        style.configure("Search.Treeview",
                        background=COLORS["bg_dark"],
                        foreground=COLORS["text"],
                        fieldbackground=COLORS["bg_dark"],
                        rowheight=28)
        style.map("Search.Treeview",
                  background=[("selected", COLORS["selection"])],
                  foreground=[("selected", COLORS["text"])])

        self.results_tree = ttk.Treeview(
            results_frame,
            columns=("name", "list_email", "matched_email"),
            show="headings",
            style="Search.Treeview"
        )
        self.results_tree.heading("name", text="Distribution List")
        self.results_tree.heading("list_email", text="List Email")
        self.results_tree.heading("matched_email", text="Matched Member")
        self.results_tree.column("name", width=180)
        self.results_tree.column("list_email", width=220)
        self.results_tree.column("matched_email", width=220)
        self.results_tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.results_tree.configure(yscrollcommand=scrollbar.set)

        # Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=10)

        ctk.CTkButton(btn_frame, text="Close", command=self.destroy, width=100,
                      fg_color="gray", hover_color="darkgray").pack()

        self.bind("<Return>", lambda e: self._search())
        self.bind("<Escape>", lambda e: self.destroy())

    def _search(self):
        """Search for email memberships using cache (instant)."""
        search_term = self.email_entry.get().strip()
        if not search_term or len(search_term) < 2:
            messagebox.showerror("Invalid Search", "Please enter at least 2 characters to search.")
            return

        if not self.main_app.cache_loaded:
            messagebox.showwarning("Not Ready", "Please wait for data to load.")
            return

        partial_match = self.partial_match_var.get()
        search_lower = search_term.lower()

        self.results_tree.delete(*self.results_tree.get_children())
        results = []

        # Search in cache (instant!)
        for list_id, data in self.main_app.members_cache.items():
            dl = data["dl"]
            for member_email in data["members"]:
                member_lower = member_email.lower()

                if partial_match:
                    is_match = search_lower in member_lower
                else:
                    is_match = member_lower == search_lower

                if is_match:
                    results.append((dl, member_email))
                    self.results_tree.insert("", "end", values=(dl.display_name, dl.mail, member_email))
                    break  # Only show one match per list

        if results:
            self.results_label.configure(
                text=f"Found '{search_term}' in {len(results)} distribution list(s)",
                text_color=COLORS["accent"]
            )
        else:
            self.results_label.configure(
                text=f"'{search_term}' not found in any distribution list",
                text_color="orange"
            )


class DistributionListManagerGUI(ctk.CTk):
    """Main application window."""

    def __init__(self):
        super().__init__()

        self.title("Distribution List Manager")
        self.geometry("1200x700")
        self.minsize(900, 500)

        # Cache for all data - stores {list_id: {"dl": DistributionList, "members": [email strings]}}
        self.members_cache = {}
        self.cache_loaded = False

        # Apply black theme to window
        self.configure(fg_color=COLORS["bg_dark"])

        # State
        self.manager: Optional[DistributionListManager] = None
        self.distribution_lists: list[DistributionList] = []
        self.current_list: Optional[DistributionList] = None
        self.current_members: list[Member] = []

        # Sort state for members list
        self.sort_column: str = "name"  # Default sort by name
        self.sort_ascending: bool = True

        # Sort state for distribution lists
        self.list_sort_ascending: bool = True

        self._setup_ui()
        self._connect()

        # Set title bar to grey (Windows 10/11)
        self.after(100, lambda: set_title_bar_color(self, "#2b2b2b"))

    def _setup_ui(self):
        """Setup the user interface."""
        # Menu bar
        self._setup_menu_bar()

        # Main container
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # Left panel - Distribution Lists
        self._setup_left_panel()

        # Right panel - Members
        self._setup_right_panel()

    def _setup_menu_bar(self):
        """Setup the menu bar."""
        self.menu_bar = tk.Menu(self, bg=COLORS["bg_medium"], fg=COLORS["text"],
                                 activebackground=COLORS["accent"], activeforeground="white")
        self.configure(menu=self.menu_bar)

        # File menu
        file_menu = tk.Menu(self.menu_bar, tearoff=0, bg=COLORS["bg_medium"], fg=COLORS["text"],
                            activebackground=COLORS["accent"], activeforeground="white")
        self.menu_bar.add_cascade(label="File", menu=file_menu)

        file_menu.add_command(label="New Distribution List...", command=self._create_list)
        file_menu.add_separator()
        file_menu.add_command(label="Export All Lists to CSV...", command=self._export_all_lists)
        file_menu.add_separator()
        file_menu.add_command(label="Import from XLSX/CSV...", command=self._import_from_csv)
        file_menu.add_command(label="Delete All and Import from XLSX/CSV...", command=self._clear_and_import_from_csv)
        file_menu.add_separator()
        file_menu.add_command(label="Refresh", command=self._refresh_lists)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)

        # Search menu
        search_menu = tk.Menu(self.menu_bar, tearoff=0, bg=COLORS["bg_medium"], fg=COLORS["text"],
                              activebackground=COLORS["accent"], activeforeground="white")
        self.menu_bar.add_cascade(label="Search", menu=search_menu)

        search_menu.add_command(label="Find Email Memberships...", command=self._search_email_memberships)

    def _setup_left_panel(self):
        """Setup the distribution lists panel."""
        left_frame = ctk.CTkFrame(self, fg_color=COLORS["bg_medium"])
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        left_frame.grid_rowconfigure(2, weight=1)
        left_frame.grid_columnconfigure(0, weight=1)

        # Header
        header = ctk.CTkLabel(left_frame, text="Distribution Lists",
                              font=ctk.CTkFont(size=18, weight="bold"))
        header.grid(row=0, column=0, pady=(10, 5), padx=10, sticky="w")

        # Search bar
        search_frame = ctk.CTkFrame(left_frame, fg_color="transparent")
        search_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)
        search_frame.grid_columnconfigure(0, weight=1)

        self.search_entry = ctk.CTkEntry(search_frame, placeholder_text="Search lists...")
        self.search_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.search_entry.bind("<KeyRelease>", self._on_search)

        ctk.CTkButton(search_frame, text="+ New", width=70,
                      command=self._create_list,
                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"]).grid(row=0, column=1, padx=(0, 5))

        ctk.CTkButton(search_frame, text="Refresh", width=70,
                      command=self._refresh_lists,
                      fg_color=COLORS["bg_light"], hover_color="#333333").grid(row=0, column=2)

        # Distribution list treeview
        tree_frame = ctk.CTkFrame(left_frame, fg_color=COLORS["bg_dark"])
        tree_frame.grid(row=2, column=0, sticky="nsew", padx=10, pady=5)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Style for treeview - Black theme
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview",
                        background=COLORS["bg_dark"],
                        foreground=COLORS["text"],
                        fieldbackground=COLORS["bg_dark"],
                        rowheight=32)
        style.configure("Treeview.Heading",
                        background=COLORS["bg_medium"],
                        foreground=COLORS["accent"],
                        font=("Segoe UI", 10, "bold"))
        style.map("Treeview",
                  background=[("selected", COLORS["selection"])],
                  foreground=[("selected", COLORS["text"])])

        self.list_tree = ttk.Treeview(tree_frame, columns=("email",), show="headings", selectmode="browse")
        self.list_tree.heading("email", text="Distribution List ▲", command=self._sort_distribution_lists)
        self.list_tree.column("email", width=300)
        self.list_tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.list_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.list_tree.configure(yscrollcommand=scrollbar.set)

        # Double-click to load members
        self.list_tree.bind("<Double-1>", self._on_list_double_click)

        # Right-click context menu for distribution lists
        self.list_context_menu = tk.Menu(self, tearoff=0, bg=COLORS["bg_medium"], fg=COLORS["text"],
                                          activebackground=COLORS["accent"], activeforeground="white")
        self.list_context_menu.add_command(label="View members", command=self._load_selected_list)
        self.list_context_menu.add_command(label="Edit list", command=self._edit_list)
        self.list_context_menu.add_separator()
        self.list_context_menu.add_command(label="Copy email", command=self._copy_list_email)
        self.list_context_menu.add_separator()
        self.list_context_menu.add_command(label="Remove all members", command=self._empty_list, foreground="orange")
        self.list_context_menu.add_command(label="Delete list", command=self._delete_list, foreground=COLORS["error"])

        self.list_tree.bind("<Button-3>", self._show_list_context_menu)

        # Hide context menu when clicking elsewhere
        self.bind("<Button-1>", self._hide_context_menus)

        # List info
        self.list_info = ctk.CTkLabel(left_frame, text="Double-click to view members",
                                       font=ctk.CTkFont(size=12), text_color="gray")
        self.list_info.grid(row=3, column=0, pady=5, padx=10, sticky="w")

    def _setup_right_panel(self):
        """Setup the members panel."""
        right_frame = ctk.CTkFrame(self, fg_color=COLORS["bg_medium"])
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 10), pady=10)
        right_frame.grid_rowconfigure(3, weight=1)  # Treeview row gets weight
        right_frame.grid_columnconfigure(0, weight=1)

        # Header with list name
        header_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        header_frame.grid_columnconfigure(0, weight=1)

        self.members_header = ctk.CTkLabel(header_frame, text="Members",
                                           font=ctk.CTkFont(size=18, weight="bold"))
        self.members_header.grid(row=0, column=0, sticky="w")

        self.edit_list_btn = ctk.CTkButton(header_frame, text="Edit List", width=100,
                                            command=self._edit_list, state="disabled",
                                            fg_color=COLORS["bg_light"], hover_color="#333333")
        self.edit_list_btn.grid(row=0, column=1, padx=5)

        # Toolbar
        toolbar = ctk.CTkFrame(right_frame, fg_color="transparent")
        toolbar.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        self.add_btn = ctk.CTkButton(toolbar, text="+ Add Member", command=self._add_member,
                                      state="disabled", width=120,
                                      fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"])
        self.add_btn.pack(side="left", padx=(0, 5))

        self.bulk_add_btn = ctk.CTkButton(toolbar, text="+ Add Multiple", command=self._bulk_add,
                                           state="disabled", width=120,
                                           fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"])
        self.bulk_add_btn.pack(side="left", padx=5)

        self.add_group_btn = ctk.CTkButton(toolbar, text="+ Add Group", command=self._add_group,
                                            state="disabled", width=100,
                                            fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"])
        self.add_group_btn.pack(side="left", padx=5)

        self.remove_btn = ctk.CTkButton(toolbar, text="Remove Selected", command=self._remove_member,
                                         state="disabled", width=130,
                                         fg_color=COLORS["error"], hover_color="#ff6666")
        self.remove_btn.pack(side="left", padx=5)

        self.export_btn = ctk.CTkButton(toolbar, text="Export", command=self._export_members,
                                         state="disabled", width=80,
                                         fg_color=COLORS["bg_light"], hover_color="#333333")
        self.export_btn.pack(side="right", padx=5)

        # Search bar for filtering members
        search_frame = ctk.CTkFrame(right_frame, fg_color="transparent")
        search_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=5)
        search_frame.grid_columnconfigure(0, weight=1)

        self.member_search_var = tk.StringVar()
        self.member_search_var.trace_add("write", self._on_member_search_changed)

        self.member_search_entry = ctk.CTkEntry(
            search_frame,
            placeholder_text="Search emails...",
            textvariable=self.member_search_var,
            fg_color="#4a4a4a",
            border_color="#5a5a5a"
        )
        self.member_search_entry.grid(row=0, column=0, sticky="ew")

        self.clear_search_btn = ctk.CTkButton(
            search_frame, text="✕", width=30,
            command=self._clear_member_search,
            fg_color=COLORS["bg_light"], hover_color="#333333"
        )
        self.clear_search_btn.grid(row=0, column=1, padx=(5, 0))

        # Members treeview
        members_tree_frame = ctk.CTkFrame(right_frame, fg_color=COLORS["bg_dark"])
        members_tree_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=5)
        members_tree_frame.grid_rowconfigure(0, weight=1)
        members_tree_frame.grid_columnconfigure(0, weight=1)

        self.members_tree = ttk.Treeview(
            members_tree_frame,
            columns=("name", "email", "type"),
            show="headings",
            selectmode="extended"
        )
        self.members_tree.heading("name", text="Name ▲", command=lambda: self._sort_members("name"))
        self.members_tree.heading("email", text="Email", command=lambda: self._sort_members("email"))
        self.members_tree.heading("type", text="Type", command=lambda: self._sort_members("type"))
        self.members_tree.column("name", width=200)
        self.members_tree.column("email", width=300)
        self.members_tree.column("type", width=100)
        self.members_tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(members_tree_frame, orient="vertical", command=self.members_tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.members_tree.configure(yscrollcommand=scrollbar.set)

        self.members_tree.bind("<<TreeviewSelect>>", self._on_member_select)
        self.members_tree.bind("<Double-1>", self._on_member_double_click)

        # Right-click context menu for members
        self.member_context_menu = tk.Menu(self, tearoff=0, bg=COLORS["bg_medium"], fg=COLORS["text"],
                                            activebackground=COLORS["accent"], activeforeground="white")
        self.member_context_menu.add_command(label="Edit member...", command=self._edit_member)
        self.member_context_menu.add_separator()
        self.member_context_menu.add_command(label="Remove from list", command=self._remove_member)
        self.member_context_menu.add_separator()
        self.member_context_menu.add_command(label="Copy email", command=self._copy_member_email)

        self.members_tree.bind("<Button-3>", self._show_member_context_menu)

        # Status bar
        self.status_label = ctk.CTkLabel(right_frame, text="",
                                          font=ctk.CTkFont(size=12), text_color="gray")
        self.status_label.grid(row=4, column=0, pady=5, padx=10, sticky="w")

    def _connect(self):
        """Initialize connection to Microsoft Graph."""
        logger.info("Connecting to Microsoft 365...")

        def do_connect():
            try:
                self.manager = DistributionListManager()
                logger.info("Connection successful!")
                self.after(0, self._on_connected)
            except Exception as e:
                logger.error(f"Connection failed: {e}")
                self.after(0, lambda: self._on_connection_error(str(e)))

        loading = LoadingDialog(self, "Connecting to Microsoft 365...")
        thread = threading.Thread(target=do_connect)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _on_connected(self):
        """Handle successful connection."""
        logger.info("Loading all distribution lists and members...")
        self._load_full_cache(is_startup=True)

    def _on_connection_error(self, error: str):
        """Handle connection error."""
        logger.error(f"Connection error displayed to user: {error}")
        messagebox.showerror(
            "Connection Error",
            f"Failed to connect to Microsoft 365:\n\n{error}\n\n"
            "Please check your .env configuration."
        )

    def _load_full_cache(self, show_dialog=True, is_startup=False):
        """Load all distribution lists and their members into cache using parallel requests."""
        if not self.manager:
            return

        if show_dialog:
            progress_dialog = ProgressDialog(self, "Loading All Data", 100)
        else:
            progress_dialog = None

        def do_load():
            cancelled = False
            try:
                # Step 1: Get all distribution lists
                if progress_dialog:
                    self.after(0, lambda: progress_dialog.update_progress(0, "Fetching distribution lists..."))

                # Check for cancellation
                if progress_dialog and progress_dialog.is_cancelled():
                    cancelled = True
                    self.after(0, lambda: self._on_load_cancelled(progress_dialog, is_startup))
                    return

                logger.info("Fetching distribution lists...")
                all_lists = self.manager.list_all()
                logger.info(f"Found {len(all_lists)} distribution lists")

                # Check for cancellation
                if progress_dialog and progress_dialog.is_cancelled():
                    cancelled = True
                    self.after(0, lambda: self._on_load_cancelled(progress_dialog, is_startup))
                    return

                if progress_dialog:
                    self.after(0, lambda: setattr(progress_dialog, 'total', len(all_lists)))

                # Step 2: Load members for each list in parallel (5 threads)
                cache = {}
                completed_count = [0]  # Use list to allow mutation in nested function
                lock = threading.Lock()

                def load_members_for_list(dl):
                    """Load members for a single distribution list."""
                    # Check for cancellation
                    if progress_dialog and progress_dialog.is_cancelled():
                        return dl, [], "cancelled"
                    try:
                        members = self.manager.get_members(dl.id)
                        return dl, members, None
                    except Exception as e:
                        logger.warning(f"Failed to load members for {dl.display_name}: {e}")
                        return dl, [], e

                with ThreadPoolExecutor(max_workers=5) as executor:
                    # Submit all tasks
                    futures = {executor.submit(load_members_for_list, dl): dl for dl in all_lists}

                    # Process results as they complete
                    for future in as_completed(futures):
                        # Check for cancellation
                        if progress_dialog and progress_dialog.is_cancelled():
                            cancelled = True
                            # Cancel remaining futures
                            for f in futures:
                                f.cancel()
                            break

                        dl, members, error = future.result()

                        if error == "cancelled":
                            continue

                        with lock:
                            if error:
                                cache[dl.id] = {"dl": dl, "members": [], "member_objects": []}
                            else:
                                cache[dl.id] = {
                                    "dl": dl,
                                    "members": [m.email for m in members],
                                    "member_objects": members
                                }
                            completed_count[0] += 1
                            count = completed_count[0]

                        # Update progress
                        if progress_dialog and not progress_dialog.is_cancelled():
                            self.after(0, lambda c=count, name=dl.display_name: progress_dialog.update_progress(
                                c, f"Loaded: {name[:35]}..."
                            ))

                if cancelled:
                    self.after(0, lambda: self._on_load_cancelled(progress_dialog, is_startup))
                    return

                # Update app state
                self.distribution_lists = all_lists
                self.members_cache = cache
                self.cache_loaded = True

                total_members = sum(len(data["members"]) for data in cache.values())
                logger.info(f"Cached {len(cache)} lists with {total_members} total members")

                self.after(0, lambda: self._on_cache_loaded(progress_dialog))

            except Exception as e:
                logger.error(f"Failed to load cache: {e}")
                if progress_dialog:
                    self.after(0, lambda: progress_dialog.destroy())
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to load: {e}"))

        thread = threading.Thread(target=do_load)
        thread.start()

    def _on_cache_loaded(self, progress_dialog=None):
        """Called when full cache is loaded."""
        if progress_dialog:
            progress_dialog.destroy()
        self._update_list_tree()
        total_members = sum(len(data["members"]) for data in self.members_cache.values())
        self.list_info.configure(text=f"{len(self.distribution_lists)} lists, {total_members} members cached")

    def _on_load_cancelled(self, progress_dialog=None, is_startup=False):
        """Called when loading is cancelled."""
        logger.info("Loading cancelled by user")
        if progress_dialog:
            progress_dialog.destroy()

        if is_startup:
            # Exit the application if cancelled during startup
            logger.info("Exiting application due to startup cancellation")
            self.quit()
            self.destroy()
        else:
            # Just show a message if cancelled during refresh
            self.list_info.configure(text="Load cancelled")

    def _refresh_lists(self):
        """Refresh all distribution lists and cache."""
        self._load_full_cache()

    def _refresh_lists_quick(self):
        """Quick refresh just the list tree from cache (no API calls)."""
        self._update_list_tree()

    def _reload_cache(self):
        """Reload the entire cache (used after bulk operations)."""
        self._load_full_cache()

    def _update_cache_add_member(self, list_id: str, email: str, member_obj=None):
        """Update cache after adding a member."""
        if list_id in self.members_cache:
            if email not in self.members_cache[list_id]["members"]:
                self.members_cache[list_id]["members"].append(email)
                # Add member object (create a simple one if not provided)
                if member_obj:
                    self.members_cache[list_id]["member_objects"].append(member_obj)
                else:
                    from distribution_list_manager import Member
                    new_member = Member(id=email, display_name=email.split('@')[0], email=email, user_type="user")
                    self.members_cache[list_id]["member_objects"].append(new_member)

    def _update_cache_remove_member(self, list_id: str, email: str):
        """Update cache after removing a member."""
        if list_id in self.members_cache:
            email_lower = email.lower()
            # Remove from email list
            self.members_cache[list_id]["members"] = [
                m for m in self.members_cache[list_id]["members"]
                if m.lower() != email_lower
            ]
            # Remove from member objects list
            self.members_cache[list_id]["member_objects"] = [
                m for m in self.members_cache[list_id]["member_objects"]
                if m.email.lower() != email_lower
            ]

    def _get_cached_member_count(self):
        """Get total cached member count."""
        return sum(len(data["members"]) for data in self.members_cache.values())

    def _update_cache_status(self):
        """Update the status bar with current cache info."""
        total_members = self._get_cached_member_count()
        self.list_info.configure(text=f"{len(self.distribution_lists)} lists, {total_members} members cached")

    def _update_list_tree(self):
        """Update the distribution list treeview with sorting."""
        self.list_tree.delete(*self.list_tree.get_children())

        search_text = self.search_entry.get().lower()
        filtered = [
            dl for dl in self.distribution_lists
            if not search_text or search_text in dl.display_name.lower() or search_text in dl.mail.lower()
        ]

        # Sort the list
        filtered.sort(key=lambda dl: dl.display_name.lower(), reverse=not self.list_sort_ascending)

        for dl in filtered:
            self.list_tree.insert("", "end", iid=dl.id, values=(f"{dl.display_name} ({dl.mail})",))

        self.list_info.configure(text=f"{len(filtered)} distribution list(s)")

    def _sort_distribution_lists(self):
        """Toggle sort order for distribution lists."""
        self.list_sort_ascending = not self.list_sort_ascending

        # Update heading with sort indicator
        indicator = " ▲" if self.list_sort_ascending else " ▼"
        self.list_tree.heading("email", text="Distribution List" + indicator)

        # Refresh the tree
        self._update_list_tree()

    def _on_search(self, event=None):
        """Handle search input."""
        self._update_list_tree()

    def _hide_context_menus(self, event=None):
        """Hide all context menus."""
        try:
            self.list_context_menu.unpost()
            self.member_context_menu.unpost()
        except:
            pass

    def _on_list_double_click(self, event=None):
        """Handle double-click on distribution list."""
        self._load_selected_list()

    def _load_selected_list(self):
        """Load the currently selected list's members."""
        selection = self.list_tree.selection()
        if not selection:
            return

        list_id = selection[0]
        self.current_list = next((dl for dl in self.distribution_lists if dl.id == list_id), None)

        if self.current_list:
            # Clear search when switching lists
            self.member_search_var.set("")
            self._load_members()
            self.members_header.configure(text=f"Members - {self.current_list.display_name}")
            self.edit_list_btn.configure(state="normal")
            self.add_btn.configure(state="normal")
            self.bulk_add_btn.configure(state="normal")
            self.add_group_btn.configure(state="normal")
            self.export_btn.configure(state="normal")

    def _show_list_context_menu(self, event):
        """Show right-click context menu for distribution lists."""
        item = self.list_tree.identify_row(event.y)
        if item:
            self.list_tree.selection_set(item)
            # Update current_list for context menu actions (but don't load members)
            self.current_list = next((dl for dl in self.distribution_lists if dl.id == item), None)
            # Show menu and grab focus
            try:
                self.list_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.list_context_menu.grab_release()

    def _copy_list_email(self):
        """Copy distribution list email to clipboard."""
        if self.current_list:
            self.clipboard_clear()
            self.clipboard_append(self.current_list.mail)
            self.list_info.configure(text=f"Copied: {self.current_list.mail}")

    def _delete_list(self):
        """Delete the selected distribution list."""
        if not self.current_list or not self.manager:
            return

        # Confirm deletion
        confirm = messagebox.askyesno(
            "Delete Distribution List",
            f"Are you sure you want to DELETE this distribution list?\n\n"
            f"Name: {self.current_list.display_name}\n"
            f"Email: {self.current_list.mail}\n\n"
            f"This action cannot be undone!",
            icon="warning"
        )

        if not confirm:
            return

        def do_delete():
            try:
                self.manager.delete_list(self.current_list.id)
                self.after(0, self._on_list_deleted)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to delete list: {e}"))

        loading = LoadingDialog(self, f"Deleting {self.current_list.display_name}...")
        thread = threading.Thread(target=do_delete)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _on_list_deleted(self):
        """Handle list deletion completion."""
        messagebox.showinfo("Success", "Distribution list deleted successfully")
        self.current_list = None
        self.current_members = []
        self.members_header.configure(text="Members")
        self.members_tree.delete(*self.members_tree.get_children())
        self.edit_list_btn.configure(state="disabled")
        self.add_btn.configure(state="disabled")
        self.bulk_add_btn.configure(state="disabled")
        self.add_group_btn.configure(state="disabled")
        self.export_btn.configure(state="disabled")
        self.remove_btn.configure(state="disabled")
        self._refresh_lists()

    def _empty_list(self):
        """Remove all members from the selected distribution list."""
        if not self.current_list or not self.manager:
            return

        # First load members to know how many there are
        def do_get_members():
            try:
                members = self.manager.get_members(self.current_list.id)
                self.after(0, lambda: self._confirm_empty_list(members))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to get members: {e}"))

        loading = LoadingDialog(self, "Loading members...")
        thread = threading.Thread(target=do_get_members)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _confirm_empty_list(self, members):
        """Confirm and execute emptying the list."""
        if not members:
            messagebox.showinfo("Info", "This list has no members.")
            return

        confirm = messagebox.askyesno(
            "Remove All Members",
            f"Are you sure you want to remove ALL members from this list?\n\n"
            f"List: {self.current_list.display_name}\n"
            f"Members to remove: {len(members)}\n\n"
            f"This will remove all {len(members)} member(s)!",
            icon="warning"
        )

        if not confirm:
            return

        # Get all emails
        emails = [m.email for m in members]
        total = len(emails)
        progress_dialog = ProgressDialog(self, "Removing All Members", total)

        def do_empty():
            results = {"success": [], "failed": []}

            for i, email in enumerate(emails):
                if progress_dialog.is_cancelled():
                    break

                self.after(0, lambda idx=i, e=email: progress_dialog.update_progress(
                    idx + 1, f"Removing {e}..."
                ))

                try:
                    self.manager.remove_member(self.current_list.id, email)
                    results["success"].append(email)
                except Exception as e:
                    results["failed"].append({"email": email, "error": str(e)})

            self.after(0, lambda: self._on_empty_complete(results, progress_dialog))

        thread = threading.Thread(target=do_empty)
        thread.start()

    def _on_empty_complete(self, results, progress_dialog=None):
        """Handle empty list completion."""
        if progress_dialog:
            progress_dialog.destroy()

        # Update cache - clear all members from this list
        for email in results["success"]:
            self._update_cache_remove_member(self.current_list.id, email)

        success = len(results["success"])
        failed = len(results["failed"])

        if failed:
            message = f"Removed {success} members.\nFailed to remove {failed}:\n\n"
            for fail in results["failed"][:5]:
                message += f"- {fail['email']}: {fail['error']}\n"
            if failed > 5:
                message += f"...and {failed - 5} more"
            messagebox.showwarning("Partial Success", message)
        else:
            messagebox.showinfo("Success", f"Removed all {success} members from the list.")

        self._load_members()
        self._update_cache_status()

    def _load_members(self):
        """Load members of the current list from cache."""
        if not self.current_list:
            return

        # Load from cache (instant!)
        if self.current_list.id in self.members_cache:
            cached_data = self.members_cache[self.current_list.id]
            self.current_members = cached_data.get("member_objects", [])
            self._update_members_tree()
        else:
            # Fallback to API if not in cache (shouldn't happen normally)
            self._load_members_from_api()

    def _load_members_from_api(self):
        """Load members from API (fallback)."""
        if not self.manager or not self.current_list:
            return

        def do_load():
            try:
                members = self.manager.get_members(self.current_list.id)
                self.current_members = members
                # Update cache
                if self.current_list.id in self.members_cache:
                    self.members_cache[self.current_list.id]["member_objects"] = members
                    self.members_cache[self.current_list.id]["members"] = [m.email for m in members]
                self.after(0, self._update_members_tree)
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to load members: {e}"))

        loading = LoadingDialog(self, "Loading members...")
        thread = threading.Thread(target=do_load)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _update_members_tree(self):
        """Update the members treeview with optional filtering and sorting."""
        self.members_tree.delete(*self.members_tree.get_children())

        # Get search filter text
        search_text = self.member_search_var.get().lower().strip()

        # Filter members if search text is provided
        if search_text:
            filtered_members = [
                m for m in self.current_members
                if search_text in m.display_name.lower() or search_text in m.email.lower()
            ]
        else:
            filtered_members = list(self.current_members)

        # Sort members
        sort_key = {
            "name": lambda m: m.display_name.lower(),
            "email": lambda m: m.email.lower(),
            "type": lambda m: m.user_type.lower()
        }.get(self.sort_column, lambda m: m.display_name.lower())

        filtered_members.sort(key=sort_key, reverse=not self.sort_ascending)

        for member in filtered_members:
            self.members_tree.insert("", "end", iid=member.id,
                                     values=(member.display_name, member.email, member.user_type))

        # Show filtered count if filtering
        if search_text:
            self.status_label.configure(text=f"{len(filtered_members)} of {len(self.current_members)} member(s)")
        else:
            self.status_label.configure(text=f"{len(self.current_members)} member(s)")
        self.remove_btn.configure(state="disabled")

    def _on_member_search_changed(self, *args):
        """Handle member search text change."""
        self._update_members_tree()

    def _clear_member_search(self):
        """Clear the member search field."""
        self.member_search_var.set("")

    def _sort_members(self, column: str):
        """Sort members by the specified column."""
        if self.sort_column == column:
            # Toggle direction if same column clicked
            self.sort_ascending = not self.sort_ascending
        else:
            # New column, default to ascending
            self.sort_column = column
            self.sort_ascending = True

        # Update column headers to show sort indicator
        self._update_sort_headers()

        # Refresh the tree
        self._update_members_tree()

    def _update_sort_headers(self):
        """Update column headers to show sort indicator."""
        columns = {
            "name": "Name",
            "email": "Email",
            "type": "Type"
        }

        for col, text in columns.items():
            if col == self.sort_column:
                indicator = " ▲" if self.sort_ascending else " ▼"
                self.members_tree.heading(col, text=text + indicator)
            else:
                self.members_tree.heading(col, text=text)

    def _on_member_select(self, event=None):
        """Handle member selection."""
        selection = self.members_tree.selection()
        self.remove_btn.configure(state="normal" if selection else "disabled")

    def _on_member_double_click(self, event=None):
        """Handle double-click on member to edit."""
        item = self.members_tree.identify_row(event.y)
        if item:
            self.members_tree.selection_set(item)
            self._edit_member()

    def _edit_member(self):
        """Show edit member dialog."""
        if not self.current_list:
            return

        selection = self.members_tree.selection()
        if not selection:
            return

        member_id = selection[0]
        member = next((m for m in self.current_members if m.id == member_id), None)
        if not member:
            return

        def on_submit(old_email: str, new_email: str):
            if old_email.lower() == new_email.lower():
                # No change
                return
            self._do_edit_member(old_email, new_email)

        EditMemberDialog(self, member, on_submit)

    def _do_edit_member(self, old_email: str, new_email: str):
        """Edit member by removing old and adding new."""
        if not self.manager or not self.current_list:
            return

        def do_edit():
            try:
                # Remove old member
                self.manager.remove_member(self.current_list.id, old_email)
                # Add new member
                self.manager.add_member(self.current_list.id, new_email)
                self.after(0, lambda: self._on_member_edited(old_email, new_email))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Failed to edit member: {msg}"))

        loading = LoadingDialog(self, "Updating member...")
        thread = threading.Thread(target=do_edit)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _on_member_edited(self, old_email: str, new_email: str):
        """Handle member edited."""
        messagebox.showinfo("Success", f"Member updated:\n{old_email} → {new_email}")

        # Update cache
        self._update_cache_remove_member(self.current_list.id, old_email)
        self._update_cache_add_member(self.current_list.id, new_email)

        # Reload members from cache
        self._load_members()
        self._update_cache_status()

    def _show_member_context_menu(self, event):
        """Show right-click context menu for members."""
        # Select the item under cursor
        item = self.members_tree.identify_row(event.y)
        if item:
            self.members_tree.selection_set(item)
            try:
                self.member_context_menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.member_context_menu.grab_release()

    def _copy_member_email(self):
        """Copy selected member's email to clipboard."""
        selection = self.members_tree.selection()
        if selection:
            member_id = selection[0]
            member = next((m for m in self.current_members if m.id == member_id), None)
            if member:
                self.clipboard_clear()
                self.clipboard_append(member.email)
                self.status_label.configure(text=f"Copied: {member.email}")

    def _add_member(self):
        """Show add member dialog."""
        if not self.current_list:
            return

        def on_submit(email: str):
            self._do_add_member(email)

        AddMemberDialog(self, on_submit)

    def _do_add_member(self, email: str):
        """Add a member to the current list."""
        if not self.manager or not self.current_list:
            return

        def do_add():
            try:
                self.manager.add_member(self.current_list.id, email)
                self.after(0, lambda: self._on_member_added(email))
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Error", f"Failed to add member: {e}"))

        loading = LoadingDialog(self, f"Adding {email}...")
        thread = threading.Thread(target=do_add)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _on_member_added(self, email: str):
        """Handle member added."""
        self._update_cache_add_member(self.current_list.id, email)
        self._update_cache_status()
        messagebox.showinfo("Success", f"Added {email} to {self.current_list.display_name}")
        self._load_members()

    def _add_group(self):
        """Show dialog to add an existing group as a member."""
        if not self.current_list:
            return

        def on_submit(group_email: str):
            self._do_add_member(group_email)

        AddGroupDialog(self, self.distribution_lists, self.current_list.id, on_submit)

    def _bulk_add(self):
        """Show bulk add dialog."""
        if not self.current_list:
            return

        def on_submit(emails: list):
            self._do_bulk_add(emails)

        BulkAddDialog(self, on_submit)

    def _do_bulk_add(self, emails: list):
        """Add multiple members to the current list."""
        if not self.manager or not self.current_list:
            return

        total = len(emails)
        progress_dialog = ProgressDialog(self, "Adding Members", total)

        def do_add():
            results = {"success": [], "failed": []}

            for i, email in enumerate(emails):
                if progress_dialog.is_cancelled():
                    break

                self.after(0, lambda idx=i, e=email: progress_dialog.update_progress(
                    idx + 1, f"Adding {e}..."
                ))

                try:
                    self.manager.add_member(self.current_list.id, email)
                    results["success"].append(email)
                except Exception as e:
                    results["failed"].append({"email": email, "error": str(e)})

            self.after(0, lambda: self._on_bulk_add_complete(results, progress_dialog))

        thread = threading.Thread(target=do_add)
        thread.start()

    def _on_bulk_add_complete(self, results: dict, progress_dialog=None):
        """Handle bulk add completion."""
        if progress_dialog:
            progress_dialog.destroy()

        # Update cache with successful additions
        for email in results["success"]:
            self._update_cache_add_member(self.current_list.id, email)

        success = len(results["success"])
        failed = len(results["failed"])

        message = f"Added {success} member(s) successfully."
        if failed:
            message += f"\n\nFailed to add {failed} member(s):"
            for fail in results["failed"][:5]:
                message += f"\n  - {fail['email']}: {fail['error']}"
            if failed > 5:
                message += f"\n  ... and {failed - 5} more"

        if failed:
            messagebox.showwarning("Partial Success", message)
        else:
            messagebox.showinfo("Success", message)

        self._load_members()
        self._update_cache_status()

    def _remove_member(self):
        """Remove selected members."""
        if not self.current_list:
            return

        selection = self.members_tree.selection()
        if not selection:
            return

        # Get selected member emails
        emails = []
        for member_id in selection:
            member = next((m for m in self.current_members if m.id == member_id), None)
            if member:
                emails.append(member.email)

        if not emails:
            return

        confirm = messagebox.askyesno(
            "Confirm Removal",
            f"Remove {len(emails)} member(s) from {self.current_list.display_name}?\n\n" +
            "\n".join(emails[:5]) + ("\n..." if len(emails) > 5 else "")
        )

        if confirm:
            self._do_remove_members(emails)

    def _do_remove_members(self, emails: list):
        """Remove members from the current list."""
        if not self.manager or not self.current_list:
            return

        total = len(emails)
        progress_dialog = ProgressDialog(self, "Removing Members", total)

        def do_remove():
            results = {"success": [], "failed": []}

            for i, email in enumerate(emails):
                if progress_dialog.is_cancelled():
                    break

                self.after(0, lambda idx=i, e=email: progress_dialog.update_progress(
                    idx + 1, f"Removing {e}..."
                ))

                try:
                    self.manager.remove_member(self.current_list.id, email)
                    results["success"].append(email)
                except Exception as e:
                    results["failed"].append({"email": email, "error": str(e)})

            self.after(0, lambda: self._on_remove_complete(results, progress_dialog))

        thread = threading.Thread(target=do_remove)
        thread.start()

    def _on_remove_complete(self, results: dict, progress_dialog=None):
        """Handle remove completion."""
        if progress_dialog:
            progress_dialog.destroy()

        # Update cache with successful removals
        for email in results["success"]:
            self._update_cache_remove_member(self.current_list.id, email)

        success = len(results["success"])
        failed = len(results["failed"])

        if failed:
            message = f"Removed {success}, failed {failed}:\n\n"
            for fail in results["failed"]:
                message += f"- {fail['email']}:\n  {fail['error']}\n"
            messagebox.showwarning("Partial Success", message)
        else:
            messagebox.showinfo("Success", f"Removed {success} member(s)")

        self._load_members()
        self._update_cache_status()

    def _create_list(self):
        """Show create list dialog."""
        if not self.manager:
            messagebox.showerror("Error", "Not connected to Microsoft 365")
            return

        def on_submit(name: str, email_alias: str, description: str):
            self._do_create_list(name, email_alias, description)

        CreateListDialog(self, on_submit)

    def _do_create_list(self, name: str, email_alias: str, description: str):
        """Create a new distribution list."""
        if not self.manager:
            return

        def do_create():
            try:
                new_list = self.manager.create_list(name, email_alias, description or None)
                self.after(0, lambda: self._on_list_created(new_list))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Failed to create list: {msg}"))

        loading = LoadingDialog(self, "Creating distribution list...")
        thread = threading.Thread(target=do_create)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _on_list_created(self, new_list):
        """Handle list created."""
        messagebox.showinfo("Success", f"Distribution list created!\n\nEmail: {new_list.mail}")

        # Add to cache
        self.distribution_lists.append(new_list)
        self.members_cache[new_list.id] = {
            "dl": new_list,
            "members": [],
            "member_objects": []
        }

        # Refresh list tree and update status
        self._update_list_tree()
        self._update_cache_status()

        # Select the new list
        self.list_tree.selection_set(new_list.id)
        self.list_tree.see(new_list.id)

    def _edit_list(self):
        """Show edit list dialog."""
        if not self.current_list:
            return

        def on_submit(name: str, description: str, mail_nickname: str):
            self._do_update_list(name, description, mail_nickname)

        EditListDialog(self, self.current_list, on_submit)

    def _do_update_list(self, name: str, description: str, new_email: str = None):
        """Update the current list using Exchange PowerShell."""
        if not self.current_list:
            return

        def do_update():
            try:
                # Use Exchange PowerShell to update distribution group
                exchange = ExchangeClient()
                exchange.update_distribution_group(
                    identity=self.current_list.mail,
                    display_name=name,
                    primary_smtp=new_email
                )
                # Update description via Graph API if we have a manager (description not supported in Exchange PS easily)
                if self.manager and description is not None:
                    try:
                        self.manager.update_list(self.current_list.id, description=description)
                    except:
                        pass  # Ignore description update errors
                self.after(0, lambda: self._on_list_updated(name, new_email))
            except Exception as e:
                error_msg = str(e)
                self.after(0, lambda msg=error_msg: messagebox.showerror("Error", f"Failed to update: {msg}"))

        loading = LoadingDialog(self, "Updating list...")
        thread = threading.Thread(target=do_update)
        thread.start()

        def check_thread():
            if thread.is_alive():
                self.after(100, check_thread)
            else:
                loading.destroy()

        self.after(100, check_thread)

    def _on_list_updated(self, new_name: str, new_email: str = None):
        """Handle list updated."""
        old_email = self.current_list.mail if self.current_list else None

        if new_email:
            messagebox.showinfo("Success", f"Distribution list updated.\nNote: Email change may take a few minutes to propagate.")
        else:
            messagebox.showinfo("Success", "Distribution list updated")
        self.members_header.configure(text=f"Members - {new_name}")

        # Update cache with new name and email
        if self.current_list and self.current_list.id in self.members_cache:
            self.members_cache[self.current_list.id]["dl"].display_name = new_name
            if new_email:
                self.members_cache[self.current_list.id]["dl"].mail = new_email

        # If email changed, update all cached lists that had the old email as a member
        if new_email and old_email and new_email != old_email:
            self._update_cache_member_email(old_email, new_email)

        self._refresh_lists()

    def _update_cache_member_email(self, old_email: str, new_email: str):
        """Update all cached lists that have old_email as a member to use new_email."""
        old_email_lower = old_email.lower()
        updated_lists = []

        for list_id, data in self.members_cache.items():
            # Update in members list (email strings)
            for i, member_email in enumerate(data["members"]):
                if member_email.lower() == old_email_lower:
                    data["members"][i] = new_email
                    updated_lists.append(data["dl"].display_name if "dl" in data else list_id)

            # Update in member_objects list
            for member in data.get("member_objects", []):
                if member.email.lower() == old_email_lower:
                    member.email = new_email

        if updated_lists:
            logger.info(f"Updated member email from {old_email} to {new_email} in {len(updated_lists)} lists: {updated_lists}")

    def _export_members(self):
        """Export members to file."""
        if not self.current_list or not self.current_members:
            return

        file_path = filedialog.asksaveasfilename(
            title="Export Members",
            defaultextension=".csv",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx"),
                ("Text files", "*.txt"),
            ],
            initialfile=f"{self.current_list.mail.split('@')[0]}_members"
        )

        if not file_path:
            return

        try:
            import pandas as pd

            if file_path.endswith(".csv"):
                df = pd.DataFrame([
                    {"name": m.display_name, "email": m.email, "type": m.user_type}
                    for m in self.current_members
                ])
                df.to_csv(file_path, index=False)
            elif file_path.endswith(".xlsx"):
                df = pd.DataFrame([
                    {"name": m.display_name, "email": m.email, "type": m.user_type}
                    for m in self.current_members
                ])
                df.to_excel(file_path, index=False)
            else:
                with open(file_path, "w") as f:
                    for m in self.current_members:
                        f.write(f"{m.email}\n")

            messagebox.showinfo("Success", f"Exported {len(self.current_members)} members to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export: {e}")

    def _export_all_lists(self):
        """Export all distribution lists and their members to CSV."""
        if not self.manager or not self.distribution_lists:
            messagebox.showwarning("No Data", "No distribution lists loaded.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Export All Distribution Lists",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="all_distribution_lists"
        )

        if not file_path:
            return

        total = len(self.distribution_lists)
        progress_dialog = ProgressDialog(self, "Exporting All Lists", total)

        def do_export():
            try:
                import pandas as pd

                # Build data: each column is a distribution list
                all_data = {}
                max_members = 0

                for i, dl in enumerate(self.distribution_lists):
                    if progress_dialog.is_cancelled():
                        break

                    self.after(0, lambda idx=i, name=dl.display_name: progress_dialog.update_progress(
                        idx + 1, f"Exporting {name[:25]}..."
                    ))

                    members = self.manager.get_members(dl.id)
                    emails = [m.email for m in members]
                    all_data[dl.mail] = emails
                    if len(emails) > max_members:
                        max_members = len(emails)

                if not progress_dialog.is_cancelled():
                    # Pad shorter lists with empty strings
                    for key in all_data:
                        while len(all_data[key]) < max_members:
                            all_data[key].append("")

                    df = pd.DataFrame(all_data)
                    df.to_csv(file_path, index=False)

                    self.after(0, lambda: self._on_export_complete(len(all_data), file_path, progress_dialog))
                else:
                    self.after(0, lambda: progress_dialog.destroy())

            except Exception as e:
                self.after(0, lambda: progress_dialog.destroy())
                self.after(0, lambda: messagebox.showerror("Export Error", f"Failed to export: {e}"))

        thread = threading.Thread(target=do_export)
        thread.start()

    def _on_export_complete(self, count, file_path, progress_dialog=None):
        """Handle export completion."""
        if progress_dialog:
            progress_dialog.destroy()
        messagebox.showinfo(
            "Export Complete",
            f"Exported {count} distribution lists to:\n{file_path}"
        )

    def _import_from_csv(self):
        """Import members to distribution lists from CSV/XLSX file."""
        if not self.manager:
            messagebox.showwarning("Not Connected", "Please wait for connection.")
            return

        file_path = filedialog.askopenfilename(
            title="Import from XLSX/CSV",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            import pandas as pd
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)

            # Show preview
            lists_to_import = list(df.columns)
            total_emails = sum(df[col].notna().sum() for col in df.columns)

            confirm = messagebox.askyesno(
                "Confirm Import",
                f"Found {len(lists_to_import)} distribution list(s) with {total_emails} total emails.\n\n"
                f"Lists:\n" + "\n".join(f"  - {name}" for name in lists_to_import[:10]) +
                ("\n  ..." if len(lists_to_import) > 10 else "") +
                f"\n\nExisting members will be skipped.\nProceed with import?"
            )

            if not confirm:
                return

            self._do_import_csv(df)

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to read file: {e}")

    def _do_import_csv(self, df):
        """Execute the CSV/XLSX import with parallel threading."""
        import pandas as pd
        from concurrent.futures import ThreadPoolExecutor
        import threading as th

        # Count total emails to process
        total_emails = sum(df[col].notna().sum() for col in df.columns)
        progress_dialog = ProgressDialog(self, "Importing", total_emails)

        def do_import():
            results = {"success": 0, "skipped": 0, "failed": 0, "errors": []}
            results_lock = th.Lock()
            processed = [0]
            processed_lock = th.Lock()

            def update_processed():
                with processed_lock:
                    processed[0] += 1
                    return processed[0]

            # Build list of all tasks: (dl, email, current_emails_set)
            tasks = []
            list_members_cache = {}  # Cache current members per list

            for list_email in df.columns:
                # Find the distribution list
                dl = self.manager.get_by_email(list_email)
                if not dl:
                    with results_lock:
                        results["errors"].append(f"List not found: {list_email}")
                    continue

                # Get current members (cache to avoid duplicate API calls)
                if dl.id not in list_members_cache:
                    current_members = self.manager.get_members(dl.id)
                    list_members_cache[dl.id] = {m.email.lower() for m in current_members}

                current_emails = list_members_cache[dl.id]

                # Get emails to add from CSV
                emails_to_add = df[list_email].dropna().tolist()
                for email in emails_to_add:
                    email = str(email).strip()
                    if email and "@" in email:
                        tasks.append((dl, email, current_emails))

            # Determine thread count based on total emails
            num_threads = 15 if len(tasks) > 30 else 10

            def add_member_task(task):
                dl, email, current_emails = task
                if progress_dialog.is_cancelled():
                    return

                p = update_processed()
                self.after(0, lambda e=email, ln=dl.display_name: progress_dialog.update_progress(
                    p, f"Adding {e[:25]} to {ln[:15]}..."
                ))

                # Skip if already a member
                if email.lower() in current_emails:
                    with results_lock:
                        results["skipped"] += 1
                    return

                try:
                    self.manager.add_member(dl.id, email)
                    with results_lock:
                        results["success"] += 1
                except Exception as e:
                    with results_lock:
                        results["failed"] += 1
                        results["errors"].append(f"{email}: {str(e)}")

            with ThreadPoolExecutor(max_workers=num_threads) as executor:
                executor.map(add_member_task, tasks)

            self.after(0, lambda: self._on_import_complete(results, progress_dialog))

        thread = threading.Thread(target=do_import)
        thread.start()

    def _on_import_complete(self, results, progress_dialog=None):
        """Handle import completion."""
        if progress_dialog:
            progress_dialog.destroy()

        message = (
            f"Import Complete!\n\n"
            f"Added: {results['success']}\n"
            f"Skipped (existing): {results['skipped']}\n"
            f"Failed: {results['failed']}"
        )

        if results["errors"]:
            message += f"\n\nErrors ({len(results['errors'])}):\n"
            message += "\n".join(results["errors"][:10])
            if len(results["errors"]) > 10:
                message += f"\n...and {len(results['errors']) - 10} more"

        if results["failed"] > 0 or results["errors"]:
            messagebox.showwarning("Import Complete", message)
        else:
            messagebox.showinfo("Import Complete", message)

        # Reload full cache after import
        self._reload_cache()

    def _clear_and_import_from_csv(self):
        """Delete ALL distribution lists and import fresh from CSV/XLSX."""
        if not self.manager:
            messagebox.showwarning("Not Connected", "Please wait for connection.")
            return

        file_path = filedialog.askopenfilename(
            title="Delete All and Import from XLSX/CSV",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            import pandas as pd
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path)
            else:
                df = pd.read_csv(file_path)

            # Column headers are list emails, rows below are members
            lists_to_create = list(df.columns)
            total_members = sum(df[col].notna().sum() for col in df.columns)

            # Get current lists that will be deleted
            current_lists = self.manager.list_all()

            confirm = messagebox.askyesno(
                "WARNING: Delete All and Import",
                f"DESTRUCTIVE OPERATION!\n\n"
                f"This will DELETE ALL {len(current_lists)} existing distribution lists "
                f"and create {len(lists_to_create)} new ones from the file.\n\n"
                f"Lists to DELETE:\n" +
                "\n".join(f"  - {dl.mail}" for dl in current_lists[:5]) +
                (f"\n  ...and {len(current_lists) - 5} more" if len(current_lists) > 5 else "") +
                f"\n\nLists to CREATE:\n" +
                "\n".join(f"  - {name}" for name in lists_to_create[:5]) +
                (f"\n  ...and {len(lists_to_create) - 5} more" if len(lists_to_create) > 5 else "") +
                f"\n\nTotal members to add: {total_members}\n\n"
                f"This action CANNOT be undone!",
                icon="warning"
            )

            if not confirm:
                return

            # Require typing "DELETE" to confirm
            confirm_dialog = ConfirmDeleteDialog(
                self,
                f"You are about to delete {len(current_lists)} distribution lists.\n"
                f"Type DELETE to confirm:"
            )
            self.wait_window(confirm_dialog)

            if not confirm_dialog.confirmed:
                return

            self._do_clear_and_import_csv(df, current_lists)

        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to read file: {e}")

    def _do_clear_and_import_csv(self, df, current_lists):
        """Execute the delete all and import operation using Exchange PowerShell for creation."""
        import pandas as pd
        import time
        from concurrent.futures import ThreadPoolExecutor, as_completed
        import threading as th

        # Calculate total operations: delete existing + create new + add members
        total_members = sum(df[col].notna().sum() for col in df.columns)
        total_operations = len(current_lists) + len(df.columns) + total_members

        progress_dialog = ProgressDialog(self, "Delete and Import", total_operations)

        def do_clear_and_import():
            results = {
                "lists_deleted": 0,
                "lists_created": 0,
                "members_added": 0,
                "failed_delete": 0,
                "failed_create": 0,
                "failed_add": 0,
                "errors": []
            }
            results_lock = th.Lock()
            processed = [0]  # Use list for mutable counter in threads
            processed_lock = th.Lock()

            def update_processed(delta=1):
                with processed_lock:
                    processed[0] += delta
                    return processed[0]

            # Step 1: Delete ALL existing distribution lists via Graph API (parallel)
            # Thread count based on total member emails
            num_members = int(total_members)
            num_threads = 15 if num_members > 30 else 10

            def delete_list_task(dl):
                if progress_dialog.is_cancelled():
                    return
                try:
                    self.manager.delete_list(dl.id)
                    with results_lock:
                        results["lists_deleted"] += 1
                except Exception as e:
                    with results_lock:
                        results["failed_delete"] += 1
                        results["errors"].append(f"Delete {dl.mail}: {str(e)}")
                p = update_processed()
                self.after(0, lambda: progress_dialog.update_progress(p, f"Deleting lists..."))

            with ThreadPoolExecutor(max_workers=num_threads) as executor:
                executor.map(delete_list_task, current_lists)

            # Step 2: Create new distribution lists via Exchange PowerShell (parallel)
            # Column header = list email (e.g., sales@domain.com)
            # Rows below = member emails
            created_lists = {}  # Map email -> created successfully
            created_lock = th.Lock()

            def create_list_task(list_email):
                if progress_dialog.is_cancelled():
                    return

                list_email_str = str(list_email).strip()
                if not list_email_str or "@" not in list_email_str:
                    with results_lock:
                        results["errors"].append(f"Invalid list email: {list_email}")
                    with created_lock:
                        created_lists[list_email_str] = False
                    update_processed()
                    return

                # Extract mail nickname (part before @) and use as display name
                mail_nickname = list_email_str.split("@")[0]
                display_name = mail_nickname.replace("-", " ").replace("_", " ").replace(".", " ").title()

                p = update_processed()
                self.after(0, lambda n=list_email_str: progress_dialog.update_progress(
                    p, f"Creating {n[:25]}..."
                ))

                try:
                    # Each thread gets its own ExchangeClient instance
                    exchange = ExchangeClient()
                    exchange.create_distribution_group(display_name, mail_nickname, list_email_str)
                    with results_lock:
                        results["lists_created"] += 1
                    with created_lock:
                        created_lists[list_email_str] = True
                except Exception as e:
                    with results_lock:
                        results["failed_create"] += 1
                        results["errors"].append(f"Create {list_email_str}: {str(e)}")
                    with created_lock:
                        created_lists[list_email_str] = False

            with ThreadPoolExecutor(max_workers=num_threads) as executor:
                executor.map(create_list_task, df.columns)

            # Wait for Exchange to sync before adding members
            if results["lists_created"] > 0:
                self.after(0, lambda: progress_dialog.update_progress(
                    processed[0], "Waiting for Exchange sync (30s)..."
                ))
                time.sleep(30)

            # Step 3: Add members via Graph API (parallel)
            # Refresh the manager's view of lists
            try:
                refreshed_lists = self.manager.list_all()
                list_map = {dl.mail.lower(): dl for dl in refreshed_lists}
            except Exception as e:
                with results_lock:
                    results["errors"].append(f"Failed to refresh lists: {str(e)}")
                list_map = {}

            # Build list of all member additions
            member_tasks = []
            for list_email in df.columns:
                list_email_str = str(list_email).strip()
                list_email_lower = list_email_str.lower()

                # Skip if creation failed
                if not created_lists.get(list_email_str, False):
                    update_processed(df[list_email].notna().sum())
                    continue

                # Find the list in Graph API
                dl = list_map.get(list_email_lower)
                if not dl:
                    with results_lock:
                        results["errors"].append(f"List not found in Graph after creation: {list_email_str}")
                    update_processed(df[list_email].notna().sum())
                    continue

                # Queue member additions
                member_emails = df[list_email].dropna().tolist()
                for member_email in member_emails:
                    member_email = str(member_email).strip()
                    if member_email and "@" in member_email:
                        member_tasks.append((dl, member_email))

            def add_member_task(task):
                dl, member_email = task
                if progress_dialog.is_cancelled():
                    return

                p = update_processed()
                self.after(0, lambda e=member_email, ln=dl.display_name: progress_dialog.update_progress(
                    p, f"Adding {e[:20]} to {ln[:15]}..."
                ))

                try:
                    self.manager.add_member(dl.id, member_email)
                    with results_lock:
                        results["members_added"] += 1
                except Exception as e:
                    with results_lock:
                        results["failed_add"] += 1
                        results["errors"].append(f"Add {member_email} to {dl.mail}: {str(e)}")

            with ThreadPoolExecutor(max_workers=num_threads) as executor:
                executor.map(add_member_task, member_tasks)

            self.after(0, lambda: self._on_clear_and_import_complete(results, progress_dialog))

        thread = threading.Thread(target=do_clear_and_import)
        thread.start()

    def _on_clear_and_import_complete(self, results, progress_dialog=None):
        """Handle clear and import completion."""
        if progress_dialog:
            progress_dialog.destroy()

        summary = (
            f"Lists deleted: {results.get('lists_deleted', 0)}\n"
            f"Lists created: {results.get('lists_created', 0)}\n"
            f"Members added: {results.get('members_added', 0)}\n"
            f"Failed to delete: {results.get('failed_delete', 0)}\n"
            f"Failed to create: {results.get('failed_create', 0)}\n"
            f"Failed to add: {results.get('failed_add', 0)}"
        )

        has_failures = (results.get("failed_delete", 0) > 0 or
                       results.get("failed_create", 0) > 0 or
                       results.get("failed_add", 0) > 0)

        if has_failures and results.get("errors"):
            # Show copyable error dialog
            ErrorLogDialog(
                self,
                "Import Complete with Errors",
                summary,
                results["errors"]
            )
        else:
            messagebox.showinfo("Success", f"Delete and Import Complete!\n\n{summary}")

        # Reload full cache after import
        self._reload_cache()

    def _search_email_memberships(self):
        """Open the search email memberships dialog."""
        if not self.manager:
            messagebox.showwarning("Not Connected", "Please wait for connection.")
            return
        SearchEmailDialog(self, self.manager, self)


def main():
    app = DistributionListManagerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
